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
    this.getTimelineConfig(),
    this.getMapConfig(),
    this.getCurrentAndyConfig(),
    this.getCurrentPaigeConfig(),
    this.getTriggersConfig(),
    this.getPublishingConfig(),
    this.getProjectsConfig(),
    this.getHandsConfig(),
    this.getAimsConfig(),
    this.getDocsConfig()
  ];
}

function getTimelineConfig() {
  const sections = ['titlesAbove', 'titles', 'headers', 'generic', 'rowBottomOutside', 'columnsOutside', 'matchers'];
  const styles = state.style.getTimeline(sections);

  return {
    name: 'Timeline',
    features: {
      copyCalendarEventsToSheet: {
        events: [Event.onCalendarEdit, Event.onSheetEdit],
        triggerColumns: ['D'],
        fromDate: 'March 29, 2021',
        eventsToNumYearsFromNow: 3,
        dateColumn: 'C',
        eventColumn: 'D',
        filterRow: 2,
        beginRow: 4
      },
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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
      },
      review: getReviewConfig(SectionMarker.aboveTitle, 'D')
    }
  };
}

function getMapConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'underMain', 'rowsOutside', 'columnsOutside']
  let styles = state.style.getTwoPanel(sections);
  styles.contents.left.fontSize = 12;
  styles.contents.right.fontSize = 9;
  styles.contents.all.rowHeight = 95;

  return {
    name: 'Map',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Map',
        text: `The 'Map' tab is based loosely around <a href='https://www.mindmapping.com/mind-map'>Mind Maps</a>, in that is is designed to work in harmony the natural functioning of the human mind and get ideas down in a mental tree-like structure.<br><br>All the text fields are free type, no need to worry about whether your edits correctly reference other areas of the dashboard. Just go ahead and start typing in whatever way matches the way things are in your mind.<br><br>There is a hierarchy inherent to mind-mapping but it shouldn't be overthought - arrange things in an intuitive way. Also, there is an inherent prioritization inherent in the positions of branches and twigs, but this doesn't constrain action. The next todo item could right now be on the furthest twig. Instead, arrage things as intuitively as you can.<br><br>The benefit of this tab is it can be referenced when building and updating Todo lists, or to get a fast but comprehensive overview of the set of current concerns.`
      },
      review: getReviewConfig()
    }
  };
}

function getCurrentAndyConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'hiddenValues', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getDefault(sections);

  return {
    name: 'Current:Andy',
    features: {
      copySheetEventsToCalendar: {
        events: [Event.onSheetEdit, Event.onOvernightTimer],
        username: 'Andy',
        priority: 'HIGH_PRIORITY',
        sheetIdForUrl: '630855359',
        workDateLabel: 'Work date',
        eventValidator: {
          method: (row, data, columns) => {
            const timing = row[columns.zeroBasedIndices.timing];
            return data.valid.filter(v => timing.endsWith(v)).length === 1;
          },
          data: { valid: [') Now', ') Next'] }
        },
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
            }
          }
        }
      },
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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
          orderSheetMainSections: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            by: {
              timing: [{ column: 'D', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
              workStream: [{ column: 'B', direction: 'ascending' }, { column: 'D', direction: 'ascending' }]
            }
          },
          setSheetValue: {
            events: [Event.onSidebarSubmit],
            update: {
              rowMarker: SectionMarker.hiddenValues,
              column: 'D',
              value: PropertyCommand.EVENT_DATA
            }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveSheetRowsToDone: {
            events: [Event.onSidebarSubmit],
            from: SectionMarker.main,
            priority: 'HIGH_PRIORITY',
            match: {
              value: ') DONE',
              column: 'D'
            }
          }
        }
      },
      review: getReviewConfig()
    }
  };
}

function getCurrentPaigeConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getDefault(sections);

  return {
    name: 'Current:Paige',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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
          orderSheetMainSections: {
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
          moveSheetRowsToDone: {
            events: [Event.onSidebarSubmit],
            from: SectionMarker.main,
            match: {
              value: ') DONE',
              column: 'C'
            }
          }
        }
      },
      review: getReviewConfig()
    }
  };
}

function getTriggersConfig() {
  return {
    name: 'Triggers',
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Triggers',
        text: `This will be help text.`
      },
      review: getReviewConfig()
    }
  };
}

function getPublishingConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getTwoPanel(sections, 2);
  styles.contents.left.fontSize = 9;
  styles.contents.right.fontSize = 9;

  return {
    name: 'Publishing',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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
        title: 'Publishing',
        text: 'This is guidance on Publishing sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
      },
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Channel' , 'Status'],
        features: {
          orderSheetMainSections: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            by: {
              channel: [{ column: 'C', direction: 'ascending' }, { column: 'D', direction: 'ascending' }],
              status: [{ column: 'D', direction: 'ascending' }, { column: 'C', direction: 'ascending' }]
            }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Done Items'],
        features: {
          moveSheetRowsToDone: {
            events: [Event.onSidebarSubmit],
            from: SectionMarker.main,
            priority: 'HIGH_PRIORITY',
            match: {
              value: ') PUBLISHED',
              column: 'D'
            }
          }
        }
      },
      review: getReviewConfig()
    }
  };
}

function getProjectsConfig() {
  return {
    name: 'Projects',
    features: {
      copySheetToExternalSpreadsheet: {
        events: [Event.onSheetEdit],
        destinationSpreadsheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
        destinationSheetName: 'Current Projects',
        nonRichTextColumnOverwrite: { column: 'H', startRow: 5 }
      }
    },
    sidebar: {
      guidance: {
        type: 'text',
        title: 'Projects',
        text: 'This is guidance on Projects sheet. It may be several lines of text, or even rich html? Nunc vulputate mauris imperdiet vehicula faucibus. Curabitur facilisis turpis libero, id volutpat velit aliquet a. Curabitur at euismod mi.'
      },
      review: getReviewConfig()
    }
  };
}

function getHandsConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getThreePanel(sections, 9, 2);
  styles.contents.left.fontSize = 9;
  styles.contents.middle.fontSize = 9;
  styles.contents.right.fontSize = 9;
  styles.contents.all.rowHeight = 44;
  styles.headers.all.fontSize = 10;
  styles.titles.between.endColumnOffset = 8;
  styles.titles.review.endColumnOffset = 7;
  styles.titles.after = state.style.getBlank({ endColumnOffset: 0, numColumns: 6, border: state.style.border.empty })

  return {
    name: 'Hands',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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
          orderSheetMainSections: {
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
          moveSheetRowsToDone: {
            events: [Event.onSidebarSubmit],
            from: SectionMarker.main,
            match: {
              value: [') Opportunity Completed', ') No Opportunity Found', ') No Reply', ') No Longer Available'],
              column: 'I'
            }
          }
        }
      },
      review: getReviewConfig(SectionMarker.title, 'I')
    }
  };
}

function getAimsConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'underMain', 'rowsOutside', 'columnsOutside'];
  let styles = state.style.getTwoPanel(sections);
  styles.contents.left.fontSize = 12;
  styles.contents.right.fontSize = 9;
  styles.contents.all.rowHeight = 160;

  return {
    name: 'Aims',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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

function getDocsConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'underMain', 'rowsOutside', 'columnsOutside'];
  let styles = state.style.getTwoPanel(sections);
  styles.contents.left.fontSize = 12;
  styles.contents.right.fontSize = 9;
  styles.contents.all.rowHeight = 120;

  return {
    name: 'Docs',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
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

function getReviewConfig(rowMarker=SectionMarker.title, column=false) {
  return {
    type: 'buttons',
    title: 'Last review',
    options: ['Today'],
    features: {
      setSheetValue: {
        events: [Event.onSidebarSubmit],
        update: {
          rowMarker: rowMarker,
          column: column || PropertyCommand.LAST_COLUMN,
          value: PropertyCommand.CURRENT_DATE
        }
      }
    }
  }
}