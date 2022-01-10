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
    replicateSheetInExternalSpreadsheet: {
      destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
      destinationSheetName: 'Current Projects',
      nonRichTextColumnOverwrite: {
        column: 'H',
        startRow: 5
      }
    },
    features: [ReplicateSheetInExternalSpreadsheet]
  };
  registerFeatureSheet(config);
}

function buildTimelineSheet() {
  const config = {
    name: 'Timeline',
    updateSpreadsheetFromCalendar: {
      fromDate: 'March 29, 2021',
      eventsToNumYearsFromNow: 3,
      dateColumn: 'C',
      eventColumn: 'D',
      filterRow: 2,
      beginRow: 4
    },
    features: [UpdateSpreadsheetFromCalendar]
  };
  registerFeatureSheet(config);
}

function buildTodoAndySheet() {
  const activeColumns = {
    noun: 'B',
    verb: 'C',
    timing: 'D',
    workDate: 'E',
    startTime: 'F',
    durationHours: 'G'
  };
  const config = {
    name: 'Todo-Andy',
    id: '630855359',
    updateCalendarFromSpreadsheet: {
      widgetCategories: {
        todo: {
          name: { column: 'B', rowOffset: -2 },
          columns: activeColumns,
          allowFillInTheBlanksDates: true
        }
      },
      scriptResponsiveWidgetNames: ['Todo:Andy']
    },
    triggerColumns: Object.values(activeColumns),
    features: [UpdateCalendarFromSpreadsheet]
  };
  registerFeatureSheet(config);
}

function isValidCustomSheetEventData(row, columns) {
  var timing = row[columns.zeroBasedIndices.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}