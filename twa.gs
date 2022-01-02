var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  projectSync: {
    handsSheetName: 'Hands',
    sourceSheetName: 'Projects',
    destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
    destinationSheetName: 'Current Projects',
    assignmentRow: 5,
    assignmentCol: 8
  },
  timelineSync: {
    timelineSheetName: 'Timeline',
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
  const activeSheet = state.spreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  const activeColumn = activeSheet.getActiveRange().getColumn();
  if(activeSheetName === config.projectSync.sourceSheetName || activeSheetName === config.projectSync.handsSheetName) {
    syncProjects();
  } else if(activeSheetName === config.timelineSync.timelineSheetName && activeColumn === config.timelineSync.eventCol) {
    syncTimelineEvents();
  }
}

function customOnOpen() {
  SpreadsheetApp.getUi().createMenu('Generate')
      .addItem('List of Upcoming Events', 'alertListOfUpcomingEvents')
      .addToUi();

  syncTimelineEvents();
}

function getNameSubstitution(name) {
  return name;
}

function setUpSheets() {


  const calendarId = state.spreadsheet.getSheetByName('(workings)').getRange('H3').getValue();
  state.twaCalendar = CalendarApp.getCalendarById(calendarId);

  setUpValuesSheet();
  setUpTodoAndySheet();
}

function setUpValuesSheet() {
  registerValuesSheet({
    name: '(workings)',
    range: 'G3:H5',
    columns: {
      users: 0,
      twaCalendar: 1
    }
  });
}

function setUpTodoAndySheet() {
  const sheetConfig = {
    name: 'Todo-Andy',
    id: '630855359',
    widgets: {
      todo: {
        columns: {
          label: 1,
          noun: 1,
          verb: 2,
          timing: 3,
          workDate: 4,
          startTime: 5,
          durationHours: 6
        },
        hasDoneCol: false,
        hasEvents: true,
        allowFillInTheBlanksDates: true
      }
    },

    scriptResponsiveWidgetNames: ['Todo']
  };

  const widgets = sheetConfig.widgets;
  sheetConfig.triggerCols = [
    widgets.todo.columns.noun,
    widgets.todo.columns.verb,
    widgets.todo.columns.timing,
    widgets.todo.columns.workDate,
    widgets.todo.columns.startTime,
    widgets.todo.columns.durationHours
  ];

  registerFeatureSheet(sheetConfig, state.features.updateCalendarFromSpreadsheet);
}

function customEventWidgetValidation(row, widget) {
  var timing = row[widget.columns.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function syncProjects() {
  var sourceSpreadsheet = SpreadsheetApp.openById(config.gsheet.id);
  var sourceSheet = sourceSpreadsheet.getSheetByName(config.projectSync.sourceSheetName);
  var destinationSpreadsheet = SpreadsheetApp.openById(config.projectSync.destinationSpreadsheetID);
  var destinationSheet = destinationSpreadsheet.getSheetByName(config.projectSync.destinationSheetName);

  overwriteProjectsWithRichTextValues(sourceSheet, destinationSheet);
  copyProjectsAssignmentValues(sourceSheet, destinationSheet);
}

function overwriteProjectsWithRichTextValues(sourceSheet, destinationSheet) {
  var sourceRange = sourceSheet.getRange(1, 1, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());
  var destinationRange = destinationSheet.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());
  destinationRange.clearContent();
  destinationRange.setRichTextValues(sourceRange.getRichTextValues());
}

function copyProjectsAssignmentValues(sourceSheet, destinationSheet) {
  var numRows = sourceSheet.getMaxRows() - config.projectSync.assignmentRow;
  var sourceRange = sourceSheet.getRange(config.projectSync.assignmentRow, config.projectSync.assignmentCol, numRows, 1);
  var destinationRange = destinationSheet.getRange(5, 8, numRows, 1);
  destinationRange.setValues(sourceRange.getValues());
}

function syncTimelineEvents() {
  config.userProperties.setProperty(config.timelineSync.listOfUpcomingEventsPropertyKey, '');
  var today = new Date();
  var oneWeekAgo = new Date(today.setDate(today.getDate() - 7));
  var timelineRanges = getTimelineRanges(state.spreadsheet.getSheetByName(config.timelineSync.timelineSheetName));
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
  return state.googleCalendar.getEvents(state.twaCalendar, today, threeYears);
}

function findCalendarEventsThisWeek(weekCommenceDate, calendarEvents, eventFilters) {
  var result = [];
  calendarEvents.forEach((calendarEvent) => {
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
  calendarEventsForCell.forEach((calendarEvent) => {
    resultStr += buildCalendarEventCellLine(calendarEvent)
  });
  return resultStr.trim('\n');
}

function formatCalendarEventsForAlert(calendarEventsForAlert) {
  var resultStr = '';
  calendarEventsForAlert.forEach((calendarEvent) => {
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
         calendarEvent.startDateTime.getDayStr() + ' ' +
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
         calendarEvent.startDateTime.getDayStr() + ', ' +
         calendarEvent.startDateTime.getMonthStr() + ' ' +
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