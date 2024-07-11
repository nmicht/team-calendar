
// Set the ID of the team calendar to add events to. You can find the calendar's
// ID on the settings page.
const TEAM_CALENDAR_ID = 'replace for your calendar id';

// Set the ID of the team calendar that contains the oncall rotation
const ONCALL_CALENDAR_ID = 'replace for your on-call rotation calendar id';

// Set the email address of the Google Group that contains everyone in the team.
const GROUP_EMAIL = 'replace for google group email';
const OTHERS_EMAIL = ['replace for other email', 'replace for other email'] 

const OOO_KEYWORDS = ['vacation', 'ooo', 'out of office', 'offline'];
const MONTHS_IN_ADVANCE = 3;

// Google Calendar Color Id for on call calendar events
const ONCALL_COLOR = 8;

/**
 * Sets up the script to run automatically every hour.
 */
function setup() {
  let triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    throw new Error('Triggers are already setup.');
  }
  ScriptApp.newTrigger('sync').timeBased().everyHours(1).create();
  // Runs the first sync immediately.
  sync();
}

/**
 * Sync calendars into team calendar
 * Looks through the group members' public calendars and adds any
 * 'vacation' or 'out of office' events to the team calendar.
 * Looks through the on call calendar and pull all the events
 * into the team calendar
 */
function sync() {
  // Defines the calendar event date range to search.
  let today = new Date();
  let maxDate = new Date();
  maxDate.setMonth(maxDate.getMonth() + MONTHS_IN_ADVANCE);

  // Determines the time the the script was last run.
  let lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  lastRun = lastRun ? new Date(lastRun) : null;

  // Gets the list of users emails in the Google Group.
  let users = getAllMembers(GROUP_EMAIL);
  users = users.concat(OTHERS_EMAIL);

  // For each user, finds events having one or more of the keywords in the event
  // summary in the specified date range. Imports each of those to the team
  // calendar.
  let countOOO = 0;
  users.forEach(function(user) {
    let OOOparams = {
      username: user.split('@')[0]
    }
    OOO_KEYWORDS.forEach(function(keyword) {
      let events = findOOOEvents(user, keyword, today, maxDate, lastRun);
      events.forEach(function(event) {
        importEvent(event, OOOparams);
        countOOO++;
      }); 
    }); 
  }); 


  // Find OnCall events. Imports each of those to the team calendar as the organizer
  let countOnCall = 0
  let OnCallParams = {
    color: ONCALL_COLOR
  }
  let events = findOnCallEvents(today, maxDate, lastRun);
  events.forEach(function(event) {
    importEvent(event, OnCallParams);
    countOnCall++;
  }); 

  PropertiesService.getScriptProperties().setProperty('lastRun', today);

  Logger.log('Imported ' + countOOO + ' team OoO events');
  Logger.log('Imported ' + countOnCall + ' OnCall events');
}

/**
 * Imports a given event into the shared team calendar.

 * @param {Calendar.Event} event The event to import.
 * @param {object} params  Extra params for sanitizing, like the team member that is attending the event.
 */
function importEvent(event, params) {

  event = sanitizeEvent(event, params) 

  // If the event is not of type 'default', it can't be imported, so it needs
  // to be changed.
  if (event.eventType != 'default') {
    event.eventType = 'default';
    delete event.outOfOfficeProperties;
    delete event.focusTimeProperties;
  }

  console.log('Importing: %s', event.summary);
  try {
    Calendar.Events.import(event, TEAM_CALENDAR_ID);
  } catch (e) {
    console.error('Error attempting to import event: %s. Skipping.',
        e.toString());
  }
}

/**
 * Sanitize and format an event before import into shared team calendar.
 * @param {object} params Extra params to sanitize, like the team member that is attending the event.
 * @param {Calendar.Event} event The event to import.
 */
function sanitizeEvent(event, params) {
  if(typeof params !== 'undefined') {
    
    // Modify title of the event
    if(typeof params.username !== 'undefined') {
      event.summary = '[' + params.username + '] ' + event.summary;
    }

    // Assign color
    if(typeof params.color !== 'undefined') {
      event.colorId = params.color;
    }
  }
  
  // Modify organizer.
  // Without this the event cannot be imported. Owner of calendar should be organizer or attendee
  event.organizer = {
    id: TEAM_CALENDAR_ID,
  };

  // Remove attendees.
  event.attendees = [];
  
  return event;
}

/**
 * In a given user's calendar, looks for occurrences of the given keyword
 * in events within the specified date range and returns any such events
 * found.
 * @param {Session.User} user The user email to retrieve events for.
 * @param {string} keyword The keyword to look for.
 * @param {Date} start The starting date of the range to examine.
 * @param {Date} end The ending date of the range to examine.
 * @param {Date} optSince A date indicating the last time this script was run.
 * @return {Calendar.Event[]} An array of calendar events.
 */
function findOOOEvents(user, keyword, start, end, optSince) {
  let params = {
    q: keyword,
    timeMin: formatDateAsRFC3339(start),
    timeMax: formatDateAsRFC3339(end),
    showDeleted: true,
  };
  if (optSince) {
    // This prevents the script from examining events that have not been
    // modified since the specified date (that is, the last time the
    // script was run).
    params.updatedMin = formatDateAsRFC3339(optSince);
  }
  let pageToken = null;
  let events = [];
  do {
    params.pageToken = pageToken;
    let response;
    try {
      response = Calendar.Events.list(user, params);
    } catch (e) {
      console.error('Error retriving events for %s, %s: %s; skipping',
          user, keyword, e.toString());
      continue;
    }
    events = events.concat(response.items.filter(function(item) {
      return shouldImportEvent(user, keyword, item);
    }));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return events;
}

/**
 * Determines if the given event should be imported into the shared team
 * calendar.
 * @param {Session.User} user The user email that is attending the event.
 * @param {string} keyword The keyword being searched for.
 * @param {Calendar.Event} event The event being considered.
 * @return {boolean} True if the event should be imported.
 */
function shouldImportEvent(user, keyword, event) {
  // Filters out events where the keyword did not appear in the summary
  // (that is, the keyword appeared in a different field, and are thus
  // is not likely to be relevant).
  if (event.summary.toLowerCase().indexOf(keyword) < 0) {
    return false;
  }
  if (!event.organizer || event.organizer.email == user) {
    // If the user is the creator of the event, always imports it.
    return true;
  }
  // Only imports events the user has accepted.
  if (!event.attendees) return false;
  let matching = event.attendees.filter(function(attendee) {
    return attendee.self;
  });
  return matching.length > 0 && matching[0].responseStatus == 'accepted';
}

/**
 * In the On call calendar, looks for occurrences
 * in events within the specified date range and returns any such events
 * found.
 * @param {Date} start The starting date of the range to examine.
 * @param {Date} end The ending date of the range to examine.
 * @param {Date} optSince A date indicating the last time this script was run.
 * @return {Calendar.Event[]} An array of calendar events.
 */
function findOnCallEvents(start, end, optSince) {
  let params = {
    timeMin: formatDateAsRFC3339(start),
    timeMax: formatDateAsRFC3339(end),
    showDeleted: true,
  };
  if (optSince) {
    // This prevents the script from examining events that have not been
    // modified since the specified date (that is, the last time the
    // script was run).
    params.updatedMin = formatDateAsRFC3339(optSince);
  }
  let pageToken = null;
  let events = [];
  do {
    params.pageToken = pageToken;
    let response;
    try {
      response = Calendar.Events.list(ONCALL_CALENDAR_ID, params);
    } catch (e) {
      console.error('Error retriving events for %s: %s; skipping',
          ONCALL_CALENDAR_ID, e.toString());
      continue;
    }

    events = response.items;

    pageToken = response.nextPageToken;
  } while (pageToken);
  return events;
}



/**
 * Returns an RFC3339 formated date String corresponding to the given
 * Date object.
 * @param {Date} date a Date.
 * @return {string} a formatted date string.
 */
function formatDateAsRFC3339(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ');
}

/**
* Get both direct and indirect members (and delete duplicates).
* @param {string} the e-mail address of the group.
* @return {object} direct and indirect members.
*/
function getAllMembers(groupEmail) {
  var group = GroupsApp.getGroupByEmail(groupEmail);
  var users = group.getUsers();
  var childGroups = group.getGroups();
  for (var i = 0; i < childGroups.length; i++) {
    var childGroup = childGroups[i];
    users = users.concat(getAllMembers(childGroup.getEmail()));
  }
  // Remove duplicate members
  var uniqueUsers = [];
  var userEmails = {};
  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    if (!userEmails[user.getEmail()]) {
      uniqueUsers.push(user);
      userEmails[user.getEmail()] = true;
    }
  }

  let emails = Object.keys(userEmails);
  return emails;

  //return uniqueUsers;
}