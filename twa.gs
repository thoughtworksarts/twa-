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
    this.getMapSheet(),
    this.getDocsSheet(),
    this.getAimsSheet()
  ];
}

function getProjectsSheet() {
  return {
    name: 'Projects',
    features: {
      copySheetToExternalSpreadsheet: {
        events: [Event.onSpreadsheetEdit],
        destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
        destinationSheetName: 'Current Projects',
        nonRichTextColumnOverwrite: { column: 'H', startRow: 5 }
      }
    }
  };
}

function getTimelineSheet() {
  const styles = this.getTimelineStyles([
    'titles',
    'headers',
    'generic',
    'rowsOutside',
    'columnsOutside',
    'matchers'
  ]);
  return {
    name: 'Timeline',
    features: {
      copyCalendarEventsToSheet: {
        events: [Event.onCalendarEdit, Event.onSpreadsheetEdit],
        triggerColumns: ['D'],
        fromDate: 'March 29, 2021',
        eventsToNumYearsFromNow: 3,
        dateColumn: 'C',
        eventColumn: 'D',
        filterRow: 2,
        beginRow: 4
      },
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetHiddenRowsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.generic,
        startRowOffset: -1,
        visibleIfMatch: {
          column: 'D',
          text: state.today.getFullYear()
        }
      },
      copySheetValuesBySection: {
        events: [Event.onOvernightTimer],
        beginColumnOffset: 3,
        from: {
          section: SectionMarker.headers,
          copyIfMatch: {
            column: 'D',
            text: state.today.getFullYear()
          }
        },
        to: {
          section: SectionMarker.title
        }
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Timeline',
        text: 'Free type in the colored lanes, and cross reference with Google Calendar events (in the grey lane on the left).'
      },
      years: {
        type: 'buttons',
        title: 'Year to display',
        options: ['2021', '2022', '2023', '2024'],
        features: {
          setSheetHiddenRowsBySection: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            section: SectionMarker.generic,
            startRowOffset: -1,
            visibleIfMatch: {
              column: 'D',
              text: PropertyCommand.EVENT_DATA
            }
          },
          copySheetValuesBySection: {
            events: [Event.onSidebarSubmit],
            beginColumnOffset: 3,
            from: {
              section: SectionMarker.headers,
              copyIfMatch: {
                column: 'D',
                text: PropertyCommand.EVENT_DATA
              }
            },
            to: {
              section: SectionMarker.title
            }
          }
        }
      },
      color: {
        type: 'text',
        title: 'Help',
        text: '1. Use these event typing conventions:<table><tr><td><pre>&nbsp;&nbsp;words?</pre></td><td>dates not yet confirmed</td></tr><tr><td><pre>&nbsp;[words]&nbsp;&nbsp;</pre></td><td>behind-the-scenes, less time-sensitive, or internal/operational</td></tr><tr><td><pre>&nbsp;&nbsp;words*</pre></td><td>holidays, admin or overriding concerns</td></tr></table><br>2. Don\'t edit the grey lane, it is overwritten by Google Calendar events. Either create an event in Google Calendar or invite <a href="mailto:jahya.net_55gagu1o5dmvtkvfrhc9k39tls@group.calendar.google.com">this email address</a> to a Google Calendar event.<br><br>3. Type into the filter box above to hide items from the grey lane below.'
      }
    }
  };
}

function getCurrentAndySheet() {
  const styles = this.getStyles([
    'titles',
    'titlesAboveBelow',
    'hiddenValues',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'rowsOutside',
    'columnsOutside'
  ]);
  return {
    name: 'Current:Andy',
    id: '630855359',
    hiddenValueRow: 3,
    features: {
      copySheetEventsToCalendar: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer],
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
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetGroupsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.done,
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
          orderSheetMainSection: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            by: {
              timing: [{ column: 'D', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
              workStream: [{ column: 'B', direction: 'ascending' }, { column: 'D', direction: 'ascending' }]
            }
          },
          setSheetHiddenValue: {
            events: [Event.onSidebarSubmit],
            cellToUpdate: { column: 'D' }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveSheetRowsMainToDone: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            match: {
              value: ') DONE',
              column: 'D'
            }
          }
        }
      }
    }
  };
}

function getHandsSheet() {
  const styles = this.getStyles([
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'rowsOutside',
    'columnsOutside'
  ]);
  styles.contents[0].rowHeight = 44;
  styles.headers[0].fontSize = 10;
  return {
    name: 'Hands',
    id: '972426638',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetGroupsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.done,
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
          orderSheetMainSection: {
            events: [Event.onSidebarSubmit],
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
          moveSheetRowsMainToDone: {
            events: [Event.onSidebarSubmit],
            match: {
              value: [') Opportunity Completed', ') No Opportunity Found', ') No Reply', ') No Longer Available'],
              column: 'I'
            }
          }
        }
      }
    }
  };
}

function getCurrentPaigeSheet() {
  const styles = this.getStyles([
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'done',
    'underMain',
    'underDone',
    'rowsOutside',
    'columnsOutside'
  ]);
  return {
    name: 'Current:Paige',
    id: '1960053305',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetGroupsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.done,
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
          orderSheetMainSection: {
            events: [Event.onSidebarSubmit],
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
          moveSheetRowsMainToDone: {
            events: [Event.onSidebarSubmit],
            match: {
              value: ') DONE',
              column: 'C'
            }
          }
        }
      }
    }
  };
}

function getMapSheet() {
  let styles = this.getTwoColumnStyles([
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'underMain',
    'rowsOutside',
    'columnsOutside'
  ]);
  styles.contents[0].rowHeight = 95;
  return {
    name: 'Map',
    id: '531646230',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
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

function getDocsSheet() {
  let styles = this.getTwoColumnStyles([
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'underMain',
    'rowsOutside',
    'columnsOutside'
  ]);
  styles.contents[0].rowHeight = 120;
  return {
    name: 'Docs',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Docs',
        text: `Quick access to and explanation of key overarching docs and links.`
      }
    }
  };
}

function getAimsSheet() {
  let styles = this.getTwoColumnStyles([
    'titles',
    'titlesAboveBelow',
    'headers',
    'main',
    'underMain',
    'rowsOutside',
    'columnsOutside'
  ]);
  styles.contents[0].rowHeight = 160;
  return {
    name: 'Aims',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSpreadsheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Aims',
        text: `High level aims of Thoughtworks Arts.`
      }
    }
  };
}

function isValidEventData(row, columns) {
  var timing = row[columns.zeroBasedIndices.timing];
  return timing == '(1) Now' || timing == '(2) Next';
}

function getTimelineStyles(sections) {
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
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      border: { top: false, left: false, bottom: false, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 2,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 7,
      fontColor: '#999999',
      background: '#f3f3f3',
      border: { top: true, left: true, bottom: true, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 3,
      fontFamily: 'Roboto Mono',
      fontSize: 10,
      fontColor: null,
      background: null,
      border: { top: true, left: true, bottom: true, right: true, vertical: true, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }],
    headers: [{
      beginColumnOffset: 0,
      numColumns: 2,
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      rowHeight: 36,
      border: { top: true, left: false, bottom: true, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 2,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 10,
      fontColor: '#666666',
      background: '#f3f3f3',
      border: { top: true, left: false, bottom: true, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 3,
      fontFamily: 'Roboto Mono',
      fontSize: 8,
      fontColor: null,
      background: null,
      border: { top: true, left: true, bottom: true, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }],
    contents: [{
      beginColumnOffset: 0,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 14,
      fontColor: null,
      background: null,
      border: { top: null, left: null, bottom: null, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 1,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 9,
      fontColor: null,
      background: null,
      border: { top: null, left: null, bottom: null, right: true, vertical: false, horizontal: false, color: '#666666', style: 'SOLID_MEDIUM' }
    }, {
      beginColumnOffset: 2,
      numColumns: 1,
      fontFamily: 'Roboto Mono',
      fontSize: 8,
      fontColor: null,
      background: null,
      borders: [
        { top: null, left: null, bottom: null, right: null, vertical: false, horizontal: true, color: '#ffffff', style: 'SOLID' },
        { top: null, left: null, bottom: null, right: true, vertical: false, horizontal: null, color: '#b7b7b7', style: 'SOLID_MEDIUM' }
      ]
    }, {
      beginColumnOffset: 3,
      fontFamily: 'Roboto Mono',
      fontSize: 7,
      fontColor: null,
      background: null,
      rowHeight: 41,
      borders: [
        { top: null, left: null, bottom: null, right: null, vertical: false, horizontal: true, color: '#ffffff', style: 'SOLID' },
        { top: true, left: null, bottom: true, right: true, vertical: null, horizontal: null, color: '#666666', style: 'SOLID_MEDIUM' }
      ]
    }, {
      beginColumnOffset: 0,
      numColumns: 3,
      border: { top: true, left: true, bottom: true, right: null, vertical: null, horizontal: null, color: '#666666', style: 'SOLID_MEDIUM' }
    }],
    rowsOutside: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      rowHeight: 9
    }],
    columnsOutside: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      columnWidth: 12
    }],
    matchers: [{
      match: {
        value: getMondayThisWeek(),
        column: 'C'
      },
      beginColumnOffset: 2,
      border: { top: true, left: true, bottom: true, right: true, vertical: null, horizontal: null, color: '#ea4335', style: 'SOLID_THICK' }
    }]
  };
  return styles;
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
      rowHeight: 9,
      border: { top: true, left: false, bottom: null, right: false, vertical: false, horizontal: false, color: '#333333', style: 'SOLID_THICK' }
    }],
    rowsOutside: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      rowHeight: 9,
      border: { top: null, left: false, bottom: false, right: false, vertical: false, horizontal: false }
    }],
    columnsOutside: [{
      fontFamily: 'Roboto Mono',
      fontSize: 1,
      fontColor: '#f3f3f3',
      background: '#f3f3f3',
      columnWidth: 12,
      border: { top: false, left: false, bottom: false, right: false, vertical: false, horizontal: false }
    }]
  };
  return styles;
}

function getTwoColumnStyles(sections) {
  let styles = this.getStyles(sections);
  const defaultFontSize = styles.contents[0].fontSize;
  styles.contents[0].fontSize = PropertyCommand.IGNORE;
  styles.contents.push({
    beginColumnOffset: 0,
    numColumns: 1,
    fontSize: 12
  }, {
    beginColumnOffset: 1,
    fontSize: defaultFontSize
  });
  return styles;
}