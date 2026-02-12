/**
 * Script to handle vacation collisions, pulling expected assignments for each date
 * across the full date range of the year, and cross-checking missed assignments
 * with manually overridden assignments.
 * 
 * Designed to work with Google Sheets and Google Apps Scripts.
 * 
 * *** NOTE ***
 * When running, please make sure that the "checkDates" function is listed as the target.
 * If this is not set, there can be unexpected behavior.
 * 
 * @version 2026-02-11
 * @author Rahul Dhodapkar
 */

//////////////////////////////////////////////////////////////////////
// LOAD DATA
//////////////////////////////////////////////////////////////////////

var ss = SpreadsheetApp.getActiveSpreadsheet();

// *** INPUT SHEETS ***
var schedulesData = ss.getSheetByName("Clinic Schedule DATA").getDataRange().getValues();
var residentRotationData = ss.getSheetByName("Rotation Assignments DATA").getDataRange().getValues();
var vacationData = ss.getSheetByName("Vacation Lookup DATA").getDataRange().getValues();
var overrideData = ss.getSheetByName("Assignment Override DATA").getDataRange().getValues();

// *** OUTPUT SHEETS ***
var blockOutputSheet = ss.getSheetByName("Block Schedule GEN")
blockOutputSheet.clear() // ***DANGEROUS COMMAND***

var adjustedAssignmentsOutputSheet = ss.getSheetByName("Adjusted Assignments GEN")
adjustedAssignmentsOutputSheet.clear() // ***DANGEROUS COMMAND***

//////////////////////////////////////////////////////////////////////
// DEFINE GLOBAL CONSTANTS
//////////////////////////////////////////////////////////////////////

// schedules headers
SCHED_RES_COL = 0
SCHED_DAYS_COL = 1
SCHED_WEEKS_COL = 2
SCHED_MONTHSMOD_COL = 3
SCHED_AMPM_COL = 4
SCHED_TEMPLATENAME_COL = 5

// rotation headers
ROT_RES_COL = 0
ROT_ROTATION_COL = 1
ROT_STARTDATE_COL = 2
ROT_ENDDATE_COL = 3

// vacation headers
VAC_RES_COL = 0
VAC_PGY_COL = 1
VAC_START_COL = 2
VAC_END_COL = 3

// override headers
OVER_RES_COL = 0
OVER_DATE_COL = 1
OVER_AMPM_COL = 2
OVER_ASSIGN_COL = 3

//////////////////////////////////////////////////////////////////////
// DEFINE UTILITY FUNCTIONS
//////////////////////////////////////////////////////////////////////

/**
 * Sets a value at a deeply nested path within an object, creating intermediate objects if they don't exist.
 * @param {object} obj - The object to modify.
 * @param {Array<string>} path - The path to the property as an array of keys.
 * @param {*} value - The value to set.
 */
function setNestedObjectOrAppend(obj, path, value) {
  if (path === undefined) {
    return
  }
  let current = obj;

  for (let i = 0; i < path.length - 1; i++) {
    const key = path[i];
    // Create an empty object if the intermediate property doesn't exist
    if (typeof current[key] !== 'object' || current[key] === null) {
      current[key] = {};
    }
    current = current[key];
  }
  // Assign the final value
  if (Array.isArray(current[path[path.length - 1]])) {
    current[path[path.length - 1]].push(value)
  } else {
    current[path[path.length - 1]] = [value];
  }
}

/**
 * Check if dates are equal, ignoring time components
 */
function datesEqual(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

//////////////////////////////////////////////////////////////////////
// GENERATE LOOKUP MAPS
//////////////////////////////////////////////////////////////////////

// AM = 0, PM = 1
// use the key pattern [day of week][week of month][month of year % 2][AM/PM]
dateToResident = {}
dateToClinic = {}

for (var i = 1; i < schedulesData.length; i++) {
  residentOnRotation = schedulesData[i][SCHED_RES_COL]
  clinicAffected = schedulesData[i][SCHED_TEMPLATENAME_COL]

  validDays = String(schedulesData[i][SCHED_DAYS_COL]).split(",")
  validWeeks = String(schedulesData[i][SCHED_WEEKS_COL]).split(",")
  validMonthsMod2 = String(schedulesData[i][SCHED_MONTHSMOD_COL]).split(",")
  validAMPM = String(schedulesData[i][SCHED_AMPM_COL]).split(",")

  for (var d = 0; d < validDays.length; d++) {
    for (var w = 0; w < validWeeks.length; w++) {
      for (var m = 0; m < validMonthsMod2.length; m++) {
        for (var a = 0; a < validAMPM.length; a++) {
          path = [validDays[d], validWeeks[w], validMonthsMod2[m], validAMPM[a]]
          setNestedObjectOrAppend(dateToResident, path, residentOnRotation);
          setNestedObjectOrAppend(dateToClinic, path, clinicAffected);
        }
      }
    }
  }
}

/**
 * Given a rotation and a date, return the resident names on rotation
 */
function getResidentForRotationAndDate(rotation, date) {
  resToReturn = []
  for (var i = 0; i < residentRotationData.length; i++) {
    if (residentRotationData[i][ROT_ROTATION_COL] != rotation) { continue; }
    if (new Date(residentRotationData[i][ROT_STARTDATE_COL]) < date 
        && new Date(residentRotationData[i][ROT_ENDDATE_COL]) >= date) {
          resToReturn.push(residentRotationData[i][ROT_RES_COL])
        }
  }
  return resToReturn
}

/**
 * Given a resident name and a date, return if the resident is on leave
 */
function isResidentOnLeave(residentName, date) {
  for (var i = 0; i < vacationData.length; i++) {
    if (vacationData[i][VAC_RES_COL] != residentName) { continue; }
    if (new Date(vacationData[i][VAC_START_COL]) <= date
        && new Date(vacationData[i][VAC_END_COL]) >= date) {
          return true
        }
  }
  return false
}

/**
 * Given a date, return all manual override assignments and residents
 */
function getAllOverrideAssignments(date, ampm) {
  assignments = []
  for (var i = 0; i < overrideData.length; i++) {
    ampm_value = []

    if (datesEqual(new Date(overrideData[i][OVER_DATE_COL]),date)
        && (ampm === String(overrideData[i][OVER_AMPM_COL])
            || ampm in String(overrideData[i][OVER_AMPM_COL]).split(","))) {
      assignments.push(overrideData[i])
    }
  }
  return assignments 
}


/**
 * Gather Staffing for Each Clinic given a list of clinics and rotation staffing
 */
function summarizeResidentStaffing(clinics, residents) {
  summary = {}

  for (var i = 0; i < clinics.length; i++) {
    var c = clinics[i]
    if (c in summary) {
      summary[c].push(residents[i])
    } else {
      summary[c] = [residents[i]]
    }
  }

  return summary
}

/**
 * Given expected staffing for a clinic and actual staffing, generate the
 * difference between the two, and implicit expected staffing
 */
function diffStaffingSummary(expected, actual) {
  var diffNum = {}
  var diffNames = {}
  
  var expectedClinics = Object.keys(expected)
  for (var i = 0; i < expectedClinics.length; i++) {
    var k = expectedClinics[i]
    var expectedNum = expected[k].length
    var actualNum = 0
    if (k in actual) { 
      actualNum = actual[k].length
    } else {
      actual[k] = []
    }
    if (expectedNum > actualNum) { // if there is not enough staffing
      diffNum[k] = expectedNum - actualNum
      diffNames[k] = {
        "expected": expected[k],
        "actual": actual[k]
      }
    }
  }

  return [diffNum, diffNames]  
}

function cloneDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


//////////////////////////////////////////////////////////////////////
// RUN LOGIC
//////////////////////////////////////////////////////////////////////

function checkDates() {

  // template for clinic blocking
  blockOut = [
    ["Date", "AM/PM", "ClinicToBlock", "Expected - Actual", "Diff Summary"]
  ]

  scheduleDataOut = [
    ["Date", "AM/PM", "Resident", "Clinic", "isOverride"]
  ]

  var startDate = new Date("2026-07-01T00:00:00");
  var endDate   = new Date("2027-06-30T00:00:00");

  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    Logger.log(cloneDate(d)); // important: clone if storing

    ////////////////////////////////////////////////////
    // extract date information
    ////////////////////////////////////////////////////

    var month = d.getMonth() + 1 // ADD 1 for 1-indexing and appropriate even-month calculation
    var dayOfWeek = d.getDay() // returns 0-6 for Sun-Sat
    if (dayOfWeek == 0) { dayOfWeek = 7 } // for compatibility

    var dayOfMonth = d.getDate()
    var weekNumOfMonth = Math.floor((dayOfMonth - 0.5) / 7) + 1

    ////////////////////////////////////////////////////
    // identify clinics on date, resident codes for date
    ////////////////////////////////////////////////////

    var ampm = ['am', 'pm']
    for (var a = 0; a < ampm.length; a++) {
      var res = dateToResident[dayOfWeek][weekNumOfMonth][month % 2][ampm[a]]
      var clinics = dateToClinic[dayOfWeek][weekNumOfMonth][month % 2][ampm[a]]
      var summary = summarizeResidentStaffing(clinics, res)
      //Logger.log("Templated staffing: ")
      //Logger.log(summary)

      // Now convert the resident rotations to names
      resNames = []
      resNamesMatchedClinics = []
      for (var i = 0; i < res.length; i++) {
        tmpResNames = getResidentForRotationAndDate(res[i], d)
        for (var j = 0; j < tmpResNames.length; j++) {
          if (clinics[i] == "<RESIDENT>") {
            resNames.push(tmpResNames[j])
            resNamesMatchedClinics.push(tmpResNames[j])
          } else {
            resNames.push(tmpResNames[j])
            resNamesMatchedClinics.push(clinics[i])
          }
        }
      }

      var expectedResidentStaffing = summarizeResidentStaffing(resNamesMatchedClinics, resNames)

      //Logger.log("Expected staffing: ")
      //Logger.log(expectedResidentStaffing)

      // now check vacations and manual override
      residentMissingNames = []
      residentMissingClinics = []
      for (var i = 0; i < resNames.length; i++) {
        if (isResidentOnLeave(resNames[i], d)) {
          residentMissingNames.push(resNames[i])
          residentMissingClinics.push(resNamesMatchedClinics[i])
        }
      }

      //Logger.log("Residents missing vacation: " + residentMissingNames)

      var overrideAssignments = getAllOverrideAssignments(d, ampm[a])
      var overrideResidentNames = []
      var overrideClinicNames = []
      for (var i = 0; i < overrideAssignments.length; i++) {
        overrideResidentNames.push(overrideAssignments[i][OVER_RES_COL])
        overrideClinicNames.push(overrideAssignments[i][OVER_ASSIGN_COL])
      }
      //Logger.log("Override assignments: ")
      //Logger.log(overrideAssignments)
      //Logger.log(overrideResidentNames)

      // remove overridden residents
      for (var i = 0; i < resNames.length; i++) {
        if (overrideResidentNames.includes(resNames[i])) {
          residentMissingNames.push(resNames[i])
          residentMissingClinics.push(resNamesMatchedClinics[i])
        }
      }

      // Logger.log("Residents missing: " + residentMissingNames)

      // capture a final list of all residents present and staffing for clinics
      var residentPresentNames = []
      var residentPresentClinics = []
      for (var i = 0; i < resNames.length; i++) {
        if (!residentMissingNames.includes(resNames[i])) {
          residentPresentNames.push(resNames[i])
          residentPresentClinics.push(resNamesMatchedClinics[i])

          scheduleDataOut.push([cloneDate(d), ampm[a], resNames[i], resNamesMatchedClinics[i], 'no'])
        }
      }

      // add back override assignments to present residents
      for (var i = 0; i < overrideAssignments.length; i++) {
        var tmpOverrideRes = overrideAssignments[i][OVER_RES_COL]
        var tmpOverrideClinic = overrideAssignments[i][OVER_ASSIGN_COL]

        residentPresentNames.push(tmpOverrideRes)
        residentPresentClinics.push(tmpOverrideClinic)

        scheduleDataOut.push([cloneDate(d), ampm[a], tmpOverrideRes, tmpOverrideClinic, 'yes'])
      }

      var residentPresentSummary = summarizeResidentStaffing(residentPresentClinics, residentPresentNames)
      var [diffNum, diffNames] = diffStaffingSummary(expectedResidentStaffing, residentPresentSummary)

      //Logger.log("Downbooking required : ")
      //Logger.log(diffNum)
      //Logger.log("Summary of Differences: ")
      //Logger.log(diffNames)

      // Downbooking Required
      var downbookClinics = Object.keys(diffNum)
      for (var i = 0; i < downbookClinics.length; i++) {
        blockOut.push([
          cloneDate(d), ampm[a], downbookClinics[i], 
          diffNum[downbookClinics[i]],
          JSON.stringify(diffNames[downbookClinics[i]], null)
        ])
      }
    }
  }

  // write full adjusted schedule data for later inspection
  adjustedAssignmentsOutputSheet.getRange(1,1,scheduleDataOut.length, scheduleDataOut[0].length).setValues(scheduleDataOut)

  // write full block requirements data for later inspection
  blockOutputSheet.getRange(1,1,blockOut.length, blockOut[0].length).setValues(blockOut)
}
