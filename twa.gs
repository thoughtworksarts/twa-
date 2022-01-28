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
    this.getHandsSheet(),
    this.getCurrentAndySheet(),
    this.getCurrentPaigeSheet(),
    this.getMapSheet()
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

function getCurrentAndySheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'hiddenValues',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'outsides'
  ];
  const styles = this.getStyles(sections);
  return {
    name: 'Current:Andy',
    id: '630855359',
    hiddenValueRow: 3,
    features: {
      updateCalendarFromSpreadsheet: {
        priority: 'HIGH_PRIORITY',
        workDateLabel: 'Work date',
        widgetCategories: {
          current: {
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
        scriptResponsiveWidgetNames: ['Current:Andy']
      },
      resetSpreadsheetStyles: styles,
      collapseDoneSection: {
        numRowsToDisplay: 3
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Usage Guidance',
        text: 'This is guidance on Current sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
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
          }
        }
      }
    }
  };
}

function getHandsSheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'outsides'
  ];
  const styles = this.getStyles(sections);
  styles.contents[0].rowHeight = 44;
  styles.headers[0].fontSize = 10;

  return {
    name: 'Hands',
    id: '972426638',
    features: {
      resetSpreadsheetStyles: styles,
      collapseDoneSection: {
        numRowsToDisplay: 3
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Hands',
        text: 'This sheet tracks when people volunteer, or "raise their hands." People are added the moment they make an enquiry, and they progress through the statuses whether they end up on a project or not.<br><br>If the same person volunteers for a second project, they are enetered again on the sheet. Usually people do one project at a time but it is possible in theory at least they could end up on the active section of this sheet twice.<br><br>The data is entered manually from Jigsaw, which takes about 1 minute per person - but the payback is people don\'t easily "fall through the cracks."'
      },
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Status', 'Project', 'Office', 'Country'],
        features: {
          orderMainSection: {
            by: {
              status:  [{ column: 'I', direction: 'ascending' }, { column: 'M', direction: 'ascending' }, { column: 'N', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }],
              project: [{ column: 'M', direction: 'ascending' }, { column: 'I', direction: 'ascending' }, { column: 'N', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }],
              office:  [{ column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }, { column: 'I', direction: 'ascending' }, { column: 'M', direction: 'ascending' }, { column: 'N', direction: 'ascending' }],
              country: [{ column: 'F', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'I', direction: 'ascending' }, { column: 'M', direction: 'ascending' }, { column: 'N', direction: 'ascending' }]
            }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveMatchingRowsFromMainToDone: {
            matchColumn: 'I',
            matchText: [') Opportunity Completed', ') No Opportunity Found', ') No Reply', ') No Longer Available']
          }
        }
      }
    }
  };
}

function getCurrentPaigeSheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'outsides'
  ];
  const styles = this.getStyles(sections);
  return {
    name: 'Current:Paige',
    id: '1960053305',
    features: {
      resetSpreadsheetStyles: styles,
      collapseDoneSection: {
        numRowsToDisplay: 3
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Usage Guidance',
        text: 'This is guidance on Current sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
      },
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Status'],
        features: {
          orderMainSection: {
            by: {
              status: [{ column: 'C', direction: 'ascending' }, { column: 'B', direction: 'ascending' }]
            }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveMatchingRowsFromMainToDone: {
            matchColumn: 'C',
            matchText: ') DONE'
          }
        }
      }
    }
  };
}

function getMapSheet() {
  const sections = [
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'underMain',
    'outsides'
  ];
  let styles = this.getStyles(sections);
  styles.contents[0].rowHeight = 95;
  styles.contents[0].fontSize = propertyOverrides.IGNORE;
  styles.contents.push({
      beginColumnOffset: 0,
      numColumns: 1,
      fontSize: 12
    }, {
      beginColumnOffset: 1,
      numColumns: 3,
      fontSize: 9
    });
  return {
    name: 'Map',
    id: '531646230',
    features: {
      resetSpreadsheetStyles: styles
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Map',
        text: `The 'Map' tab is based loosely around <a href='https://www.mindmapping.com/mind-map'>Mind Maps</a>, in that is is designed to work in harmony the natural functioning of the human mind and get ideas down in a mental tree-like structure.<br><br>All the text fields are free type, no need to worry about whether your edits correctly reference other areas of the dashboard. Just go ahead and start typing in whatever way matches the way things are in your mind.<br><br>There is a hierarchy inherent to mind-mapping but it shouldn't be overthought - arrange things in an intuitive way. Also, there is an inherent prioritization inherent in the positions of branches and twigs, but this doesn't constrain action. The next todo item could right now be on the furthest twig. Instead, arrage things as intuitively as you can.<br><br>The benefit of this tab is it can be referenced when building and updating Todo lists, or to get a fast but comprehensive overview of the set of current concerns.`
      }
    }
  };
}

function isValidEventData(row, columns) {
  var timing = row[columns.zeroBasedIndices.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function getStyles(sections) {
  let styles = {
    sections: sections,
    titles: [{
      beginColumnOffset: 0,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 24,
      fontColor: '#0c0c0c',
      background: '#f3f3f3',
      rowHeight: 55,
      border: { top: false, left: false, bottom: false, right: false, vertical: false, horizontal: false }
    }, {
      beginColumnOffset: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      border: { top: false, left: false, bottom: false, right: false, vertical: false, horizontal: false }
    }],
    titlesAboveBelow: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      rowHeight: 9
    }],
    hiddenValues: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3'
    }],
    headers: [{
      fontFamily: 'Roboto Mono',
      fontSize: 13,
      fontColor: '#ffffff',
      background: '#999999',
      rowHeight: 56,
      border: { top: true, left: false, bottom: true, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' }
    }],
    contents: [{
      fontFamily: 'Roboto Mono',
      fontSize: 9,
      fontColor: null,
      background: null,
      rowHeight: 48,
      border: { top: null, left: false, bottom: null, right: false, vertical: false, horizontal: true, color: '#999999', style: 'SOLID' }
    }],
    underContents: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      border: { top: true, left: false, bottom: null, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' }
    }],
    outsides: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      border: { top: false, left: false, bottom: false, right: false, vertical: false, horizontal: false }
    }]
  };
  return styles;
}