
function onOpen() {
  SpreadsheetApp.getUi().createMenu('LSR App')
    .addItem('⬇ Download from Intervals', 'fillActivities')
    .addItem('⬇ Download marked activity laps', 'fillLaps')
    .addSeparator()
    .addItem('📔 Update journal', 'fillJournal')
    .addItem('📔 Update wellness', 'fillWellnessJournal')
    .addItem('❌ Clear journals', 'clearJournals')
    .addToUi();
}

const ZONES = {
  z1: readSettings("Z1"),
  z2: readSettings("Z2"),
  z3: readSettings("Z3"),
  z4: readSettings("Z4")
};

const DEFAULTS = {
  INIT_WELLNESSDAYS_TO_DOWNLOAD: 300,
  INIT_DAYS_TO_DOWNLOAD: 50,
  MIN_INCLUDED_LAP_DISTANCE_METERS: 800 // if >1km
};
  
/**
 * laps - detailed / done
 * wellness - kartu / done
 * raw incl. wellness data
 */

function readSettings(key, fallbackValue = null) {
  const SETTINGS_SHEET = "Settings";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!settingsSheet) {
    ss.insertSheet().setName(SETTINGS_SHEET);
  }

  for(var i = 1; i < 20; i++) {
    const settingKey = settingsSheet.getRange(i, 1).getValue();
    if (settingKey === key) {
      return settingsSheet.getRange(i, 2).getValue();
    }
  }
  if (!fallbackValue === null) {
    throw Error("Setting " + key + " not found!");
  } else {
    return fallbackValue;
  }

}


function clearJournals() {
  var { rawSheet, lapsSheet } = _getNeededSheets();
  rawSheet.clear();
  lapsSheet.clear();
  _getOrCreateSheet("Wellness").clear();
}


function fillActivities() {
  fillActivityData();
  fillWellnessData();
}

function fillLaps () {
  const till = getDateString();
  const from = getDateString(21);
  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  const byDate = groupActivitiesByDay(
      fetchIntervalsActivityData(
        userId, apiKey,
        from, till
      )
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  const col = sheet.getActiveCell().getColumn();
      
  const dateRow = findDateRow(row, col);
  const dateToGet = sheet.getRange(dateRow, col).getValue();

  const activities = byDate[dateToGet];
  const activityToGet = activities.find(a => isRun(a));
  const activity = fetchIntervalsActivityIntervals(activityToGet.id, apiKey);

  const newRow = checkOrInsertNewRow(sheet, row, col);

  const groupLapsSetting = readSettings("group_laps", false);
  const laps = groupLapsSetting ? activity.icu_groups.reduce((a, i) => a + i.count + "x " + printLap(i), "") : activity.icu_intervals.reduce((a, i) => a + printLap(i), "");
  
  sheet.getRange(newRow, col).setValue(laps);
}

function fillWellnessJournal() {

    const sheet = _getOrCreateSheet("Wellness");


  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  var lastDate = sheet.getRange(2, 2).getValue();
  const lastId = sheet.getRange(2, 1).getValue();
  
  if (lastDate === "") {
    lastDate = getDateString(DEFAULTS.INIT_WELLNESSDAYS_TO_DOWNLOAD);
  } else {
    lastDate = new Date(lastDate).toISOString().split('T')[0];
  }

  const till = getDateString();
  const from = lastDate;

  const wellnessList = fetchIntervalsWellnessData(userId, apiKey, from, till);


  // Write column names
  WELLNESS_DATA.forEach((d, idx) => sheet.getRange(1, idx+1).setValue(_getFieldName(d))); 
  wellnessList.forEach(w => {
    sheet.insertRowAfter(1);
    WELLNESS_DATA.forEach((field, idx) => sheet.getRange(2, idx+1).setValue(
        _getFieldValue(field, w)
    ));
  })
}

function fillJournal() {
  var { rawSheet, lapsSheet } = _getNeededSheets();

  var lastDate = rawSheet.getRange(2, 2).getValue();
  const lastId = rawSheet.getRange(2, 1).getValue();
  
  if (lastDate === "") {
    lastDate = getDateString(DEFAULTS.INIT_DAYS_TO_DOWNLOAD);
  } else {
    lastDate = new Date(lastDate).toISOString().split('T')[0];
  }

  const till = getDateString();
  const from = lastDate;
  
  ZONES.z1 = readSettings("Z1");
  ZONES.z2 = readSettings("Z2");
  ZONES.z3 = readSettings("Z3");
  ZONES.z4 = readSettings("Z4");

  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  const activities = fetchIntervalsActivityData(userId, apiKey, from, till);
 
  /**
  * Fill Activity Journal
  */

  // Write column names
  RUN_DATA.forEach((d, idx) => rawSheet.getRange(1, idx+1).setValue(_getFieldName(d))); 
  RUN_LAPS_DATA.forEach((d, idx) => lapsSheet.getRange(1, idx+5).setValue(_getFieldName(d))); 

  for(var i = 0; i < activities.length; i++) {
    const activity = activities[i];
    if (lastId === activity.id) continue;

    if (_shouldIncludeInJournal(activity)) {
      rawSheet.insertRowAfter(1);
      RUN_DATA.forEach((field, idx) => rawSheet.getRange(2, idx+1).setValue(
        _getFieldValue(field, activity)
      ));

      if (_shouldGetLaps(activity)) {
        const laps = fetchIntervalsActivityIntervals(activity.id, apiKey);        
        
        laps.icu_intervals
          .filter(i=>_shouldIncludeInterval(i))
          .forEach((interval, interval_no) => {
              lapsSheet.insertRowAfter(1);
              lapsSheet.getRange(1, 1).setValue("reference activity id");
              lapsSheet.getRange(2, 1).setValue(activity.id);
              lapsSheet.getRange(1, 2).setValue("Activity Name");
              lapsSheet.getRange(2, 2).setValue(activity.name);
              lapsSheet.getRange(1, 3).setValue("Activity Date");
              lapsSheet.getRange(2, 3).setValue(activity.start_date_local);
              lapsSheet.getRange(1, 4).setValue("Interval Number");
              lapsSheet.getRange(2, 4).setValue(parseInt(interval_no) + 1);

              RUN_LAPS_DATA.forEach((d, idx) => lapsSheet.getRange(2, idx + 5).setValue( 
                _getFieldValue(d, interval) 
              ))
            })
      }
    }
  }
}

const WELLNESS_DATA = [
    "id",
    "weight",
    "restingHR",
    "hrv",
    {name: "sleep (hours)", fn: w => w.sleepSecs/3600 },
    "sleepScore",
    "sleepQuality",
    "avgSleepingHR",
    "soreness",
    "fatigue",
    "stress",
    "mood",
    "motivation",
    "injury",
    "comments"
];

// max_speed	average_speed   calories lthr	icu_resting_hr	icu_weight average_altitude	min_altitude	max_altitude  
const RUN_DATA = [
  "id", 
  "start_date_local",
  "type", 
  "name", 
  // { name: "elapsed_time", fn: a => getDuration(a.elapsed_time) }, 
  { name: "moving_time", fn: a => getDuration(a.moving_time) },
  { name: "distance", fn: a => getDistance(a.distance) },
  "pace",
  { name: "pace", fn: a => getPace(a.pace) },
  { name: "efficiency (kmh/bpm)", fn: a => a.average_speed * 3.6 / a.average_heartrate },
  
  "average_heartrate",
  "max_heartrate",
  "total_elevation_gain",
  "interval_summary",
  "perceived_exertion",
  "icu_intensity",
  "icu_training_load",
  "polarization_index",
  "average_cadence",
  "average_stride",
  { name: "zone", fn: a => _getHrZone(a.average_heartrate) },
  // "race"
];

const RUN_LAPS_DATA = [
  // "group_id", 
  "type",
  "start_time",
  { name: "distance", fn: a => getDistance(a.distance) },
  { name: "moving_time", fn: a => getDuration(a.moving_time) }, 
  { name: "average_speed", fn: a => getPace(a.average_speed) },
  { name: "efficiency (kmh/bpm)", fn: a => a.average_speed * 3.6 / a.average_heartrate },
  "intensity",
  "average_speed",
  "min_speed",
  "max_speed",
  "average_heartrate", 
  "max_heartrate",
  "average_cadence", 
  "total_elevation_gain", 
  "average_gradient",
  { name: "zone", fn: a => _getHrZone(a.average_heartrate) },
];

function _getHrZone(hr) {
  if (hr < ZONES.z1) return "Z1";
  if (hr < ZONES.z2) return "Z2";
  if (hr < ZONES.z3) return "Z3";
  return "Z4"
}

function _getFieldName(field) {
  if ("string" === typeof field) return field;
  return field.name; 
}
function _getFieldValue(field, activity) {
  if ("string" === typeof field) return activity[field];
  return field.fn(activity); 
}

function _getNeededSheets() {
  const RAW_SHEET = "Runs_Data";
  const LAPS_SHEET = "Laps_Data";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName(RAW_SHEET);
  var lapsSheet = ss.getSheetByName(LAPS_SHEET);

  if (!rawSheet) {
    ss.insertSheet().setName(RAW_SHEET);
    rawSheet = ss.getSheetByName(RAW_SHEET);
  }
  if (!lapsSheet) {
    ss.insertSheet().setName(LAPS_SHEET);
    lapsSheet = ss.getSheetByName(LAPS_SHEET);
  } 
  return {rawSheet, lapsSheet};
}

function _getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet().setName(sheetName);
  }
  return sheet;
}

// Filter activities to include in journal
function _shouldIncludeInJournal(activity) {
  return isRun(activity);
}
function _shouldIncludeInterval(interval) {
  return interval.distance >= DEFAULTS.MIN_INCLUDED_LAP_DISTANCE_METERS;// interval.type !== "RECOVERY";
}

// Check if need to download activity laps. 
// Other possible markers:
// icu_intensity > 80
// icu_rpe > 3
// icu_training_load
// icu_intervals_edited
function _shouldGetLaps(activity) {
  return true;
}







/** === COMMON === */

function isEmptyCell(row, col) {
  return !sheet.getRange(row, col).getValue();
}

function checkOrInsertNewRow(sheet, row, col) {
  if (!isEmptyCell(row, col)) {
    sheet.insertRowAfter(row);
    return row + 1;
  } else {
    return row;
  }
}

/** Returns date in ISO format with 00 clock time */
function getDateString(minusDays) {
  var today = new Date();
  if (minusDays) {
    today.setDate(today.getDate() - minusDays);
  }
  var month = (today.getMonth()+1);

  var date = today.getFullYear()+'-'+ formatNumber(month) +'-' + formatNumber(today.getDate());
  return date;
}

/** Is used to locate date row in case there is an additional row in between */
function findDateRow(curRow, curCol) {
  var rowToCheck = curRow - 1;
  var maybeDate = sheet.getRange(rowToCheck, curCol).getValue();
  if (!Number.isInteger(maybeDate)) {
  return findDateRow(rowToCheck, curCol);
  } else {
    return rowToCheck;
  }
}

/** Used to add 0 if single digit number */
function formatNumber(n){
    return n > 9 ? "" + n: "0" + n;
}

/** Clear user props */
function clearCache() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
}

/** Save user prop */
function saveProp(key, value) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(key, value);
}

/** Read user prop */
function readProp(key) {
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(key);
}


/** === Intervals === */

function fillActivityData() {
  const till = getDateString();
  const from = getDateString(21);
  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  printActivities(
    groupActivitiesByDay(
      fetchIntervalsActivityData(
        userId, apiKey,
        from, till
      )
    )
  );
}

function fillWellnessData() {
  const till = getDateString();
  const from = getDateString(21);
  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  writeToSheet(
    fetchIntervalsWellnessData(
      userId, apiKey,
      from, till
    )
  );
}

function fetchIntervalsActivityData(userId, apiKey, dateFrom, dateTo) {
      var endpoint = 'https://intervals.icu/api/v1/athlete/' + userId + '/activities';
      var params = '?oldest='+dateFrom+'&newest='+dateTo;
    
      var options = {
        method : 'GET',
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode("API_KEY:" + apiKey),
        }
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
      // Logger.log(response);
      return response.reverse(); 
}

function fetchIntervalsActivityIntervals(activityId, apiKey) {
      var endpoint = 'https://intervals.icu/api/v1/activity/'+activityId+'/intervals';
      
      var options = {
        method : 'GET',
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode("API_KEY:" + apiKey),
        }
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(endpoint, options));
      // Logger.log(response);
      return response; 
}

function fetchIntervalsWellnessData(userId, apiKey, dateFrom, dateTo) {
      var endpoint = 'https://intervals.icu/api/v1/athlete/' + userId + '/wellness';
      var params = '?oldest='+dateFrom+'&newest='+dateTo;
    
      var options = {
        method : 'GET',
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode("API_KEY:" + apiKey),
        }
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
      // Logger.log(response);

      return response; 
}

function writeToSheet(intervalsData) {
      var byDate = {};
      intervalsData.map(data => {
        byDate[data.id.slice(-2)] = data;
      });

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getActiveSheet();

      var row = sheet.getActiveCell().getRow();
      var col = sheet.getActiveCell().getColumn();
      
      const dateRow = findDateRow(row, col);
      const wellnessRow = dateRow + 1;
      sheet.insertRowAfter(dateRow);

      for(colIdx = col; colIdx < 20; colIdx++) {
          const day = ("0" + sheet.getRange(dateRow, colIdx).getValue()).slice(-2);
          const wellness = byDate[day];
          if (wellness) {
            sheet.getRange(wellnessRow, colIdx).setBackgroundRGB(255,255,0).setFontColor("black");
            sheet.getRange(wellnessRow, colIdx).setValue("hrv: " + wellness.hrv + " rhr: " + wellness.restingHR);  
          }
      } 
}


function groupActivitiesByDay(activities) {
  var byDate = {};
  activities.map(function(a) {
    var date = new Date(a.start_date).getDate();
    var currentDateActivities = byDate[date];
    if (currentDateActivities) {
      byDate[date] = [...currentDateActivities, a];
    } else {
      byDate[date] = [a];
    }
  });
  return byDate;
}

  function printActivities(byDate) {

      var totals = {
        swim_duration: 0,
        swim_dist: 0,
        
        bike_duration: 0,
        bike_dist: 0,
        
        run_duration: 0,
        run_dist: 0,
        run_elevation: 0,

        other_duration: 0,
        //remove
        duration: 0,
        distance: 0,
        elevation: 0
      };

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getActiveSheet();
      var row = sheet.getActiveCell().getRow();
      var col = sheet.getActiveCell().getColumn();
      
      var dateRow = findDateRow(row, col);
      row = checkOrInsertNewRow(sheet, row, col);

      for(colIdx = col; colIdx < 20; colIdx++) {
        var dateToGet = sheet.getRange(dateRow, colIdx).getValue();
        var currentCellValue = sheet.getRange(row, colIdx).getValue();
        
        if (!dateToGet) break;
        var activities = byDate[dateToGet];
        
        if (activities) {
          var data = "";
          activities.forEach(
            function(a){
              data = printActivityData(a, totals) + "\n" + data;
            })
          sheet.getRange(row, colIdx).setValue(data).setVerticalAlignment('top');  
        } else {
          sheet.getRange(row, colIdx).setValue("-").setVerticalAlignment('top');
        }
      }
      
      writeTotals(row, totals);
  }

  function isRun(a) {
    return a.type == "Run" || a.type == "TrailRun" || a.type == "VirtualRun";
  }

  function isWorkout(a) {
    return a.type == "Workout" || a.type == "WeightTraining" || a.type == "Yoga";
  }

  function isSwim(a) {
    return a.type == "Swim";
  }

  function isRide(a) {
    return a.type == "Ride" || a.type == "VirtualRide";
  }
  function isSki(a) {
    return a.type == "NordicSki";
  }

  function printActivityData(a, totals) {
    if (isRun(a)) {
      var laps = "";
      if (a.workout_type == 3) {
        laps = printLaps(a.id);
      }
      totals.run_dist = totals.run_dist + a.distance;
      totals.run_duration = totals.run_duration + a.moving_time;
      totals.run_elevation = totals.run_elevation + a.total_elevation_gain;
      
      return printRun(a) + laps; 
    }
    if (isWorkout(a)) {
      totals.other_duration = totals.other_duration + a.moving_time;
      
      return printWorkout(a);
    }
    if (isSwim(a)) {
      totals.swim_duration = totals.swim_duration + a.moving_time;
      
      return printSwim(a);
    }
    if (isRide(a)) {
      totals.bike_dist = totals.bike_dist + a.distance;
      totals.bike_duration = totals.bike_duration + a.moving_time;

      return printRide(a);
    }
    if (isSki(a)) {
      return printSki(a);
    }
    return printRun(a);
  }

  function updateTotals(totals, distance, time, elev) {
    return {
      duration: totals.duration + time,
      distance: totals.distance + distance,
      elevation: totals.elevation + elev
    };
  }

  function writeTotals(row, totals) {
    sheet.getRange(row, 9).setValue(getDistance(totals.run_dist)).setVerticalAlignment('top');  
    sheet.getRange(row, 10).setValue(getDuration(totals.run_duration)).setVerticalAlignment('top');  
    
    var run = "";
    if (totals.run_duration) {
      run = "run time: " + getDuration(totals.run_duration) + "\n";
    }
    var bike = "";
    if (totals.bike_duration){
      bike = "bike time: " + getDuration(totals.bike_duration) + "\n";
    }
    var swim = "";
    if (totals.swim_duration) {
     workout = "swim time: " + getDuration(totals.swim_duration) + "\n";
    }
    var workout = "";
    if (totals.other_duration) {
     workout = "gym time: " + getDuration(totals.other_duration) + "\n";
    }
    sheet.getRange(row, 12).setValue(
      run + bike + swim + workout
    ).setVerticalAlignment('top');  
  }

  function secondsToTime(totalSeconds) {
    var min = Math.floor(totalSeconds/60);
    var sec = Math.floor(totalSeconds-min*60);
    
    return min + ":" + (sec < 10 ? "0"+sec : sec);
  }

  // Convert m/s -> min/km
  function getPace(metersPerSec) {
    var secondsPerKm = parseInt(1/(metersPerSec/1000));
    return secondsToTime(secondsPerKm);
  }

  // Convert m/s -> min/100m
  function getSwimPace(metersPerSec) {
    var secondsPerKm = parseInt(1/(metersPerSec/100));
    return secondsToTime(secondsPerKm);
  }

  // Convert m/s -> km/h
  function getSpeed(metersPerSec) {
    return Number.parseFloat(metersPerSec * 3.6).toFixed(2);
  }

  function getDistance(stravaDistance)  {
    return Number.parseFloat(stravaDistance / 1000).toFixed(2);
  }

  function getDuration(activity_seconds) {
      var sec_num = parseInt(activity_seconds, 10);
      var hours   = Math.floor(sec_num / 3600);
      var minutes = Math.floor((sec_num - (hours * 3600)) / 60);
      var seconds = sec_num - (hours * 3600) - (minutes * 60);

      if (minutes < 10) { minutes = "0" + minutes;}
      if (seconds < 10) { seconds = "0" + seconds;}
      return hours+'h '+ minutes + 'm ' + seconds + "s";
  }

  function getHr(hr) {
    return hr ? Math.round(hr) : "--"; 
  }

