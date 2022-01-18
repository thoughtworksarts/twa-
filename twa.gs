var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  toggles: {
    performDataUpdates: true,
    verboseLogging: false,
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
    features: {
      replicateSheetInExternalSpreadsheet: {
        destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
        destinationSheetName: 'Current Projects',
        nonRichTextColumnOverwrite: { column: 'H', startRow: 5 }
      }
    }
  };
  registerFeatureSheet(config);
}

function buildTimelineSheet() {
  const config = {
    name: 'Timeline',
    features: {
      updateSpreadsheetFromCalendar: {
        fromDate: 'March 29, 2021',
        eventsToNumYearsFromNow: 3,
        dateColumn: 'C',
        eventColumn: 'D',
        filterRow: 2,
        beginRow: 4
      }
    }
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
    triggerColumns: Object.values(activeColumns),
    hiddenValueRow: 3,
    features: {
      updateCalendarFromSpreadsheet: {
        priority: 'HIGH_PRIORITY',
        widgetCategories: {
          todo: {
            name: { column: 'C', rowOffset: -2 },
            columns: activeColumns,
            allowFillInTheBlanksDates: true
          }
        },
        scriptResponsiveWidgetNames: ['Todo:Andy']
      },
      collapseDoneSection: {
        numRowsToDisplay: 5
      },
      resetSpreadsheetStyles: getDefaultSpreadsheetStyles()
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Usage Guidance',
        text: 'This is guidance on Todo sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
      },
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Timing' , 'Work Stream'],
        features: {
          orderMainSection: {
            priority: 'HIGH_PRIORITY',
            by: {
              timing: [{ column: 'D', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
              workStream: [{ column: 'B', direction: 'ascending' }, { column: 'D', direction: 'ascending' }]
            }
          },
          updateSheetHiddenValue: {
            cellToUpdate: { column: 'D' }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveMatchingRowsFromMainToDone: {
            priority: 'HIGH_PRIORITY',
            matchColumn: activeColumns.timing,
            matchText: ') DONE'
          },
          collapseDoneSection: {
            numRowsToDisplay: 5
          }
        }
      }
    }
  };
  registerFeatureSheet(config);
}

function isValidCustomSheetEventData(row, columns) {
  var timing = row[columns.zeroBasedIndices.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function getDefaultSpreadsheetStyles() {
  return {
    titles: {
      fontFamily: 'Roboto Mono',
      fontSize: 24,
      rowHeight: 55
    },
    titlesAboveBelow: {
      rowHeight: 9,
      fontFamily: 'Roboto Mono',
      fontSize: 1
    },
    hiddenValues: {
      fontFamily: 'Roboto Mono',
      fontSize: 1,
    },
    headers: {
      fontFamily: 'Roboto Mono',
      fontSize: 12,
      border: { top: true, left: false, bottom: true, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' },
      rowHeight: 50
    },
    contentSections: {
      fontFamily: 'Roboto Mono',
      fontSize: 9,
      fontColor: null,
      border: { top: null, left: false, bottom: null, right: false, vertical: false, horizontal: true, color: '#999999', style: 'SOLID' },
      rowHeight: 40
    },
    underContentSections: {
      border: { top: true, left: false, bottom: null, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' }
    }
  };
}