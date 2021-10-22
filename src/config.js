// Wish Tally Const
var WISH_TALLY_SHEET_SOURCE_ID = '1mTeEQs1nOViQ-_BVHkDSZgfKGsYiLATe1mFQxypZQWA';
/* Add URL here to avoid showing on Sheet */
var AUTO_IMPORT_URL_FOR_API_BYPASS = ""; // Optional

var WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME = "Character Event Wish History";
var WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME = "Weapon Event Wish History";
var WISH_TALLY_PERMANENT_WISH_SHEET_NAME = "Permanent Wish History";
var WISH_TALLY_NOVICE_WISH_SHEET_NAME = "Novice Wish History";
var WISH_TALLY_WISH_HISTORY_SHEET_NAME = "Wish History";
var WISH_TALLY_SETTINGS_SHEET_NAME = "Settings";
var WISH_TALLY_DASHBOARD_SHEET_NAME = "Dashboard";
var WISH_TALLY_CHANGELOG_SHEET_NAME = "Changelog";
var WISH_TALLY_PITY_CHECKER_SHEET_NAME = "Pity Checker";
var WISH_TALLY_EVENTS_SHEET_NAME = "Events";
var WISH_TALLY_CHARACTERS_OLD_SHEET_NAME = "Constellation";
var WISH_TALLY_CHARACTERS_SHEET_NAME = "Characters";
var WISH_TALLY_WEAPONS_SHEET_NAME = "Weapons";
var WISH_TALLY_RESULTS_SHEET_NAME = "Results";
var WISH_TALLY_README_SHEET_NAME = "README";
var WISH_TALLY_NAME_OF_WISH_HISTORY = [WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME, WISH_TALLY_PERMANENT_WISH_SHEET_NAME, WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME, WISH_TALLY_NOVICE_WISH_SHEET_NAME];

// AutoHotkey Const
var AUTOHOTKEY_SHEET_NAME = "AutoHotkey";
var AUTOHOTKEY_SCRIPT_SHEET_NAME = "AutoHotkey-Script";

// Import Const
var IMPORT_STATUS_COMPLETE = "COMPLETE";
var IMPORT_STATUS_FAILED = "FAILED";
var IMPORT_STATUS_WISH_HISTORY_COMPLETE = "DONE";
var IMPORT_STATUS_WISH_HISTORY_NOT_FOUND = "NOT FOUND";
var IMPORT_STATUS_WISH_HISTORY_EMPTY = "EMPTY";

// Scheduler Const
var SCHEDULER_TRIGGER_ON_TEXT = "ON";
var SCHEDULER_TRIGGER_OFF_TEXT = "OFF";
var SCHEDULER_RUN_AT_WHICH_DAY = {
    "Monday": ScriptApp.WeekDay.MONDAY,
    "Tuesday": ScriptApp.WeekDay.TUESDAY,
    "Wednesday": ScriptApp.WeekDay.WEDNESDAY,
    "Thursday": ScriptApp.WeekDay.THURSDAY,
    "Friday": ScriptApp.WeekDay.FRIDAY,
    "Saturday": ScriptApp.WeekDay.SATURDAY,
    "Sunday": ScriptApp.WeekDay.SUNDAY
  };
  var SCHEDULER_RUN_AT_HOUR = {
    "Run at 1:00": 1,
    "Run at 2:00": 2,
    "Run at 3:00": 3,
    "Run at 4:00": 4,
    "Run at 5:00": 5,
    "Run at 6:00": 6,
    "Run at 7:00": 7,
    "Run at 8:00": 8,
    "Run at 9:00": 9,
    "Run at 10:00": 10,
    "Run at 11:00": 11,
    "Run at 12:00": 12,
    "Run at 13:00": 13,
    "Run at 14:00": 14,
    "Run at 15:00": 15,
    "Run at 16:00": 16,
    "Run at 17:00": 17,
    "Run at 18:00": 18,
    "Run at 19:00": 19,
    "Run at 20:00": 20,
    "Run at 21:00": 21,
    "Run at 22:00": 22,
    "Run at 23:00": 23,
    "Run at Midnight": 0
  };
  var SCHEDULER_RUN_AT_EVERY_HOUR = {
    "Every hour": 1,
    "Every 2 hours": 2,
    "Every 3 hours": 3,
    "Every 4 hours": 4,
    "Every 5 hours": 5,
    "Every 6 hours": 6,
    "Every 7 hours": 7,
    "Every 8 hours": 8,
    "Every 9 hours": 9,
    "Every 10 hours": 10,
    "Every 11 hours": 11,
    "Every 12 hours": 12,
    "Every 13 hours": 13,
    "Every 14 hours": 14,
    "Every 15 hours": 15,
    "Every 16 hours": 16,
    "Every 17 hours": 17,
    "Every 18 hours": 18,
    "Every 19 hours": 19,
    "Every 20 hours": 20,
    "Every 21 hours": 21,
    "Every 22 hours": 22,
    "Every 23 hours": 23,
    "Every 24 hours": 24
  };