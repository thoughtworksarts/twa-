var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  projectSync: {
    handsTab: 'Hands',
    sourceTab: 'Projects',
    destinationSheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
    destinationTab: 'Current Projects',
    assignmentRow: 5,
    assignmentCol: 8
  },
  timelineSync: {
    timelineTab: 'Timeline',
    listOfUpcomingEventsPropertyKey: 'listOfUpcomingEvents',
    eventsFromDate: 'March 29, 2021',
    dateCol: 3,
    eventCol: 4,
    filterRow: 2,
    beginRow: 4
  },
  toggles: {
    performDataUpdates: true,
    showLogAlert: false
  },
  userProperties: PropertiesService.getUserProperties()
}

function customOnEdit() {
  const activeSubsheet = state.spreadsheet.getActiveSheet();
  const activeSubsheetName = activeSubsheet.getName();
  const activeColumn = activeSubsheet.getActiveRange().getColumn();
  if(activeSubsheetName === config.projectSync.sourceTab || activeSubsheetName === config.projectSync.handsTab) {
    syncProjects();
  } else if(activeSubsheetName === config.timelineSync.timelineTab && activeColumn === config.timelineSync.eventCol) {
    syncTimelineEvents();
  }
}

function customOnOpen() {
  SpreadsheetApp.getUi().createMenu('Generate')
      .addItem('List of Upcoming Events', 'alertListOfUpcomingEvents')
      .addToUi();
}

function customUpdates() {
  syncTimelineEvents();
}

function getNameSubstitution(name) {
  return name;
}

function preProcessSubsheets() {
  state.personValuesSubsheet = new PersonValuesSubsheet(state.spreadsheet, '(workings)', { start:'G3', end:'G5' });
  const calendarId = state.spreadsheet.getSheetByName('(workings)').getRange('H3').getValue();
  state.twaCalendar = CalendarApp.getCalendarById(calendarId);
  buildTodoAndySubsheet();
}

function buildTodoAndySubsheet() {
  const range = {
    offsets: {
      row: 2,
      col: 2
    },
    maxRows: 500,
    maxCols: 11
  };

  const sections = {
    todo: {
      columns: {
        label: 2,
        noun: 2,
        verb: 3,
        timing: 4,
        workDate: 5,
        startTime: 6,
        durationHours: 7
      },
      rangeColumns: {},
      hasDoneCol: false,
      hasEvents: true,
      allowFillInTheBlanksDates: true
    }
  };

  const triggerCols = [
    sections.todo.columns.noun,
    sections.todo.columns.verb,
    sections.todo.columns.timing,
    sections.todo.columns.workDate,
    sections.todo.columns.startTime,
    sections.todo.columns.durationHours
  ];

  state.eventSubsheets.push(new EventSubsheet(state.spreadsheet, 'Todo-Andy', '630855359', range, sections, triggerCols));
}

function postProcessSubsheets() {
  state.validEventCategories = ['Todo'];
}

function isSpecificValidEventData(row, section) {
  var timing = row[section.rangeColumns.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function syncProjects() {
  var sourceSheet = SpreadsheetApp.openById(config.gsheet.id);
  var sourceTab = sourceSheet.getSheetByName(config.projectSync.sourceTab);
  var destinationSheet = SpreadsheetApp.openById(config.projectSync.destinationSheetID);
  var destinationTab = destinationSheet.getSheetByName(config.projectSync.destinationTab);

  overwriteProjectsWithRichTextValues(sourceTab, destinationTab);
  copyProjectsAssignmentValues(sourceTab, destinationTab);
}

function overwriteProjectsWithRichTextValues(sourceTab, destinationTab) {
  var sourceRange = sourceTab.getRange(1, 1, sourceTab.getMaxRows(), sourceTab.getMaxColumns());
  var destinationRange = destinationTab.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());
  destinationRange.clearContent();
  destinationRange.setRichTextValues(sourceRange.getRichTextValues());
}

function copyProjectsAssignmentValues(sourceTab, destinationTab) {
  var numRows = sourceTab.getMaxRows() - config.projectSync.assignmentRow;
  var sourceRange = sourceTab.getRange(config.projectSync.assignmentRow, config.projectSync.assignmentCol, numRows, 1);
  var destinationRange = destinationTab.getRange(5, 8, numRows, 1);
  destinationRange.setValues(sourceRange.getValues());
}

function syncTimelineEvents() {
  config.userProperties.setProperty(config.timelineSync.listOfUpcomingEventsPropertyKey, '');
  var today = new Date();
  var oneWeekAgo = new Date(today.setDate(today.getDate() - 7));
  var timelineRanges = getTimelineRanges(state.spreadsheet.getSheetByName('Timeline'));
  const calendarEvents = getTWACalendarEvents();
  var listOfUpcomingEventsForAlert = '';

  for(var i = 0; i < timelineRanges.dateValues.length; i++) {
    var weekCommenceDate = timelineRanges.dateValues[i][0];
    if(weekCommenceDate instanceof Date) {
      var calendarEventsThisWeek = findCalendarEventsThisWeek(weekCommenceDate, calendarEvents, timelineRanges.eventFilters);
      timelineRanges.eventValues[i][0] = calendarEventsThisWeek.length > 0 ? formatCalendarEventsForCell(calendarEventsThisWeek) : '';
      if(oneWeekAgo < weekCommenceDate) {
        listOfUpcomingEventsForAlert += calendarEventsThisWeek.length > 0 ? formatCalendarEventsForAlert(calendarEventsThisWeek) : '';
      }
    }
  }

  timelineRanges.eventRange.setValues(timelineRanges.eventValues);
  config.userProperties.setProperty(config.timelineSync.listOfUpcomingEventsPropertyKey, listOfUpcomingEventsForAlert);
}

function getTimelineRanges(timelineSheet) {
  const numRows = timelineSheet.getMaxRows() - config.timelineSync.beginRow + 1;
  var dateRange = timelineSheet.getRange(config.timelineSync.beginRow, config.timelineSync.dateCol, numRows, 1);
  var eventRange = timelineSheet.getRange(config.timelineSync.beginRow, config.timelineSync.eventCol, numRows, 1);
  var eventFilterRange = timelineSheet.getRange(config.timelineSync.filterRow, config.timelineSync.eventCol, 1, 1);

  return {
    dateRange: dateRange,
    eventRange: eventRange,
    dateValues: dateRange.getValues(),
    eventValues: eventRange.getValues(),
    eventFilters: eventFilterRange.getValue().split('\n')
  };
}

function getTWACalendarEvents() {
  const today = new Date(config.timelineSync.eventsFromDate);
  const threeYears = new Date();
  threeYears.setFullYear(threeYears.getFullYear() + 3);
  return getCalendarEvents(state.twaCalendar, today, threeYears);
}

function findCalendarEventsThisWeek(weekCommenceDate, calendarEvents, eventFilters) {
  var result = [];
  calendarEvents.forEach(function(calendarEvent) {
    if(isValidCalendarEventForWeek(calendarEvent, weekCommenceDate, eventFilters)) {
      result.push(calendarEvent);
    }
  });
  return result;
}

function isValidCalendarEventForWeek(calendarEvent, weekCommenceDate, eventFilters) {
  const weekConcludeDate = weekCommenceDate.addDays(7);
  return calendarEvent.startDateTime >= weekCommenceDate &&
         calendarEvent.startDateTime < weekConcludeDate &&
         !eventFilters.includes(calendarEvent.title);
}

function formatCalendarEventsForCell(calendarEventsForCell) {
  var resultStr = '';
  calendarEventsForCell.forEach(function(calendarEvent) {
    resultStr += buildCalendarEventCellLine(calendarEvent)
  });
  return resultStr.trim('\n');
}

function formatCalendarEventsForAlert(calendarEventsForAlert) {
  var resultStr = '';
  calendarEventsForAlert.forEach(function(calendarEvent) {
    resultStr += buildCalendarEventAlertLine(calendarEvent)
  });
  return resultStr;
}

function buildCalendarEventCellLine(calendarEvent) {
  const dayNumberStart = calendarEvent.startDateTime.getDate();
  const dayNumberEnd = getDateMinusFewSeconds(calendarEvent.endDateTime).getDate();
  const unsureDate = calendarEvent.title.endsWith('?');
  const prefix = unsureDate ? '[?] ' : '';

  return prefix +
         calendarEvent.startDateTime.getDayShortStr() + ' ' +
         dayNumberStart +
         (dayNumberStart === dayNumberEnd ? '' : '-' + dayNumberEnd) + ': ' +
         (dayNumberStart <= 9 && dayNumberStart === dayNumberEnd && !unsureDate ? ' ' : '') +
         calendarEvent.title + '\n';
}

function buildCalendarEventAlertLine(calendarEvent) {
  const dayNumberStart = calendarEvent.startDateTime.getDate();
  const dayNumberEnd = getDateMinusFewSeconds(calendarEvent.endDateTime).getDate();
  const unsureDate = calendarEvent.title.endsWith('?');
  const prefix = unsureDate ? '[?] ' : '';

  return prefix +
         calendarEvent.startDateTime.getDayLongStr() + ', ' +
         calendarEvent.startDateTime.getMonthLongStr() + ' ' +
         dayNumberStart +
         (dayNumberStart === dayNumberEnd ? '' : '-' + dayNumberEnd) + ': ' +
         (dayNumberStart <= 9 && dayNumberStart === dayNumberEnd && !unsureDate ? ' ' : '') +
         calendarEvent.title + '\n';
}

function alertListOfUpcomingEvents() {
  var listOfUpcomingEvents = config.userProperties.getProperty(config.timelineSync.listOfUpcomingEventsPropertyKey) || '';
  alert(listOfUpcomingEvents.length == 0 ? 'Processing updates... try again in a minute.' : listOfUpcomingEvents);
}

function getDateMinusFewSeconds(givenDate) {
  return new Date(givenDate - 5000);
}