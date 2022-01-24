function getSpreadsheetConfig() {
  return {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  };
}

function getValuesSheetConfig() {
  return {
    name: '(workings)',
    range: 'G3:H5',
    usersColumnIndex: 0,
    eventsCalendarIdRowIndex: 0,
    eventsCalendarIdColumnIndex: 1
  };
}

function getFeatureSheetConfigs() {
  return [
    this.getProjectsSheet(),
    this.getTimelineSheet(),
    this.getTodoAndySheet(),
    this.getReservoirSheet()
  ];
}

function getProjectsSheet() {
  return {
    name: 'Projects',
    features: {
      replicateSheetInExternalSpreadsheet: {
        destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
        destinationSheetName: 'Current Projects',
        nonRichTextColumnOverwrite: { column: 'H', startRow: 5 }
      }
    }
  };
}

function getTimelineSheet() {
  return {
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
}

function getTodoAndySheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'hiddenValues',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone'
  ];
  return {
    name: 'Todo-Andy',
    id: '630855359',
    hiddenValueRow: 3,
    features: {
      updateCalendarFromSpreadsheet: {
        priority: 'HIGH_PRIORITY',
        workDateLabel: 'Work date',
        widgetCategories: {
          todo: {
            name: { column: 'C', rowOffset: -1 },
            columns: {
              noun: 'B',
              verb: 'C',
              timing: 'D',
              workDate: 'E',
              startTime: 'F',
              durationHours: 'G'
            },
            allowFillInTheBlanksDates: true
          }
        },
        scriptResponsiveWidgetNames: ['Todo']
      },
      collapseDoneSection: {
        numRowsToDisplay: 5
      },
      resetSpreadsheetStyles: this.getDefaultSheetStyles(sections)
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
            matchColumn: 'D',
            matchText: ') DONE'
          },
          collapseDoneSection: {
            numRowsToDisplay: 5
          }
        }
      }
    }
  };
}

function getReservoirSheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'mainSubRanges',
    'underMain'
  ];
  const overrides = {
    contents: {
      rowHeight: 95,
      fontSize: propertyOverrides.IGNORE
    }
  };
  const appends = {
    contentsSubRanges: [{
        beginColumnOffset: 0,
        numColumns: 2,
        fontSize: 12
      }, {
        beginColumnOffset: 2,
        numColumns: 3,
        fontSize: 9
      }
    ]
  };
  return {
    name: 'Reservoir',
    id: '531646230',
    features: {
      resetSpreadsheetStyles: this.getDefaultSheetStyles(sections, overrides, appends)
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Reservoir',
        text: 'This is guidance on Reservoir sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
      }
    }
  };
}

function isValidEventData(row, columns) {
  var timing = row[columns.zeroBasedIndices.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function getDefaultSheetStyles(sections, overrides={}, appends={}) {
  let sheetStyle = {
    sections: sections,
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
    contents: {
      fontFamily: 'Roboto Mono',
      fontSize: 9,
      fontColor: null,
      border: { top: null, left: false, bottom: null, right: false, vertical: false, horizontal: true, color: '#999999', style: 'SOLID' },
      rowHeight: 40
    },
    underContents: {
      border: { top: true, left: false, bottom: null, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' }
    }
  };
  sheetStyle = this.overrideSheetStyles(sheetStyle, overrides);
  sheetStyle = this.appendSheetStyles(sheetStyle, appends);
  return sheetStyle;
}

function overrideSheetStyles(sheetStyle, overrides) {
  for(const sectionKey in overrides) {
    const section = overrides[sectionKey];
    for(const propertyKey in section) {
      const property = section[propertyKey];
      sheetStyle[sectionKey][propertyKey] = property;
    }
  }
  return sheetStyle;
}

function appendSheetStyles(sheetStyle, appends) {
  sheetStyle = Object.assign(sheetStyle, appends);
  return sheetStyle;
}