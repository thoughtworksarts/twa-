var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  toggles: {
    performDataUpdates: true,
    logAllEvents: false,
    showLogAlert: false
  }
}

function buildSheets() {
  buildValuesSheet();
  buildProjectsSheet();
  buildTimelineSheet();
  buildTodoAndySheet();
}

function buildValuesSheet() {
  registerValuesSheet({
    name: '(workings)',
    range: 'G3:H5',
    usersColumnIndex: 0,
    eventsCalendarIdRowIndex: 0,
    eventsCalendarIdColumnIndex: 1
  });
}

function buildProjectsSheet() {
  const config = {
    name: 'Projects',
    destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
    destinationSheetName: 'Current Projects',
    nonRichTextColumnOverwrite: {
      column: 'H',
      startRow: 5
    }
  };
  registerFeatureSheet(config, [state.features.replicateSheetInExternalSpreadsheet]);
}

function buildTimelineSheet() {
  const config = {
    name: 'Timeline',
    fromDate: 'March 29, 2021',
    eventsToNumYearsFromNow: 3,
    dateColumn: 'C',
    eventColumn: 'D',
    filterRow: 2,
    beginRow: 4
  };
  registerFeatureSheet(config, [state.features.updateSpreadsheetFromCalendar]);
}

function buildTodoAndySheet() {
  const config = {
    name: 'Todo-Andy',
    id: '630855359',
    widgets: {
      todo: {
        name: {
          column: 'B', rowOffset: -2
        },
        columns: {
          noun: 'B',
          verb: 'C',
          timing: 'D',
          workDate: 'E',
          startTime: 'F',
          durationHours: 'G'
        },
        hasEvents: true,
        allowFillInTheBlanksDates: true
      }
    },

    scriptResponsiveWidgetNames: ['Todo:Andy']
  };

  const widgets = config.widgets;
  config.triggerColumns = [
    widgets.todo.columns.noun,
    widgets.todo.columns.verb,
    widgets.todo.columns.timing,
    widgets.todo.columns.workDate,
    widgets.todo.columns.startTime,
    widgets.todo.columns.durationHours
  ];

  registerFeatureSheet(config, [state.features.updateCalendarFromSpreadsheet]);
}

function customEventWidgetValidation(row, widget) {
  var timing = row[widget.columns.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}