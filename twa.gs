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
    this.getDashboardConfig(),
    this.getTimelineConfig(),
    this.getMapConfig(),
    this.getCurrentAndyConfig(),
    this.getCurrentPaigeConfig(),
    this.getTriggersConfig(),
    this.getPublishingConfig(),
    this.getProjectsConfig(),
    this.getHandsConfig(),
    this.getStakeholdersConfig(),
    this.getNetworkConfig(),
    this.getSpendConfig(),
    this.getStreamsConfig(),
    this.getAimsConfig(),
    this.getDocsConfig()
  ];
}

function getDashboardConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'underMain', 'columnsOutside', 'rowsOutside'];
  const styles = state.style.getFourPanel(sections, 1, 2, 1);
  styles.headers.all.fontSize = PropertyCommand.IGNORE;
  styles.headers.left = { fontSize: 13, beginColumnOffset: 0, numColumns: 1 };
  styles.headers.middle = { fontSize: 9, beginColumnOffset: 1, numColumns: 2 };
  styles.headers.right = { fontSize: 13, beginColumnOffset: 3 };
  styles.contents.left.fontSize = 12;
  styles.contents.leftMiddle.fontSize = 9;
  styles.contents.daysCol = { beginColumnOffset: 2, numColumns: 1, border: state.style.border.thinPanelDivider };
  styles.contents.rightMiddle.fontSize = 9;
  styles.contents.right.fontSize = 9;
  Object.assign(styles.titles.review, state.style.getBlank());

  return {
    name: 'Dashboard',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
    },
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Dashboard'
      },
      guidance: {
        type: 'text',
        title: 'Guidance',
        text: 'Edit the numbers here to change when review dates turn yellow and then red. However, edit the due dates themselves using the sidebars on each of the individual sheets.'
      }
    }
  };
}

function getTimelineConfig() {
  const sections = ['titlesAbove', 'titles', 'headers', 'generic', 'rowBottomOutside', 'columnsOutside', 'rowMatchers'];
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
        visibilityMatcher: {
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
      heading: {
        type: 'heading',
        title: 'Timeline'
      },
      review: getReviewConfig(SectionMarker.aboveTitle, 'D'),
      years: {
        type: 'buttons',
        title: 'Year to display',
        options: ['2021', '2022', '2023', '2024', '2025', '2026'],
        features: {
          setSheetHiddenRowsBySection: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            section: SectionMarker.generic,
            startRowOffset: -1,
            visibilityMatcher: {
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
      guidance: {
        type: 'ul',
        title: 'Guidance',
        texts: ['Free type in the colored lanes', 'Cross reference with GCal (left)']
      },
      help: {
        type: 'text',
        title: 'More help',
        text: '1. Use these event typing conventions:<table><tr><td><pre>&nbsp;&nbsp;words?</pre></td><td>dates not yet confirmed</td></tr><tr><td><pre>&nbsp;[words]&nbsp;&nbsp;</pre></td><td>behind-the-scenes, less time-sensitive, or internal/operational</td></tr><tr><td><pre>&nbsp;&nbsp;words*</pre></td><td>holidays, admin or overriding concerns</td></tr></table><br>2. Don\'t edit the grey lane, it is overwritten by Google Calendar events. Either create an event in Google Calendar or invite <a href="mailto:jahya.net_55gagu1o5dmvtkvfrhc9k39tls@group.calendar.google.com">this email address</a> to a Google Calendar event.<br><br>3. Type into the filter box above to hide items from the grey lane below, or put [brackets] around the event title in GCal.'
      }
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
      heading: {
        type: 'heading',
        title: 'Map'
      },
      review: getReviewConfig()
    }
  };
}

function getCurrentAndyConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'hiddenValues', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getDefault(sections);
  styles.headers.all.fontSize = PropertyCommand.IGNORE;
  styles.headers.left = { fontSize: 13, beginColumnOffset: 0, numColumns: 3 };
  styles.headers.middle = { fontSize: 9, beginColumnOffset: 3, numColumns: 3 };
  styles.headers.right = { fontSize: 13, beginColumnOffset: 6 };

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
          data: { valid: [') Priority', ') Following'] }
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
      heading: {
        type: 'heading',
        title: 'Current:Andy'
      },
      review: getReviewConfig(),
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
      create: {
        type: 'buttons',
        title: 'Create',
        options: ['Priority', 'Following', 'Rolling'],
        features: {
          createSheetItem: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            getValues: (option) => {
              const options = {
                'Priority':  '(1) Priority',
                'Following': '(2) Following',
                'Rolling':   '(3) Rolling'
              };
              const timing = options[option];
              return ['', '', timing, '', '', '', '', ''];
            }
          },
          setSheetStylesBySection: {
            events: [Event.onSidebarSubmit],
            styles: styles
          },
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
      }
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
        id: 'setSheetStylesBySection.Current:Paige',
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetGroupsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.done,
        numRowsToDisplay: 3
      },
      createSheetItem: {
        id: 'createSheetItem.Current:Paige'
      }
    },
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Current:Paige'
      },
      review: getReviewConfig(),
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
      }
    }
  };
}

function getTriggersConfig() {
  return {
    name: 'Triggers',
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Triggers'
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

  const get = {
    featureId: {
      createInPublishing: 'createSheetItem.Publishing',
      createInCurrentPaige: 'createSheetItem.Current:Paige',
      setSheetStylesBySectionPaige: 'setSheetStylesBySection.Current:Paige'
    },
    defaultStatus: {
      publishing: '(0) Not Started',
      currentPaige: '(1) Todo'
    },
    channel: {
      twaNewsletter: 'TWA.io: Newsletter',
      mediumBlog: 'Medium: Blog',
      social: 'Social (FB, Twitter, Inst)',
      internalStakeholder: 'Internal: Stakeholder Update'
    },
    publishingSectionIndex: {
      "Paige": 1,
      "Andy": 2
    },
    message: {
      createTwaNewsletter: (triggerText) =>          { return 'Publish TWA.io Newsletter for\n' + triggerText; },
      createMedium:        (triggerText) =>          { return 'Copy to Medium for\n' + triggerText; },
      createSocial:        (triggerText) =>          { return 'Social media for\n' + triggerText; },
      contactAssociates:   (channel, triggerText) => { return 'Contact associates for ' + channel + '\n' + triggerText; },
      contactInternal:     (channel, triggerText) => { return 'Share on GChat groups and with relevant stakeholders for ' + channel + '\n' + triggerText; },
      generatedBy:         (channel) =>              { return 'Generated by publication of\n' + channel; },
      contactHow: 'Share social AND blog links with\n1) authors 2) partners and 3) any mentioned or relevants people',
      contactWhenDone: 'Set this to DONE'
    }
  };

  const executeFeature = (id, values=false, section=false) => {
    const feature = state.builder.findFeature(id) || false;
    if(values) feature.config.values = values;
    if(section) feature.config.insertionSectionIndex = get.publishingSectionIndex[section];
    feature.execute();
  }

  const publicationResponses = {
    "TinyLetter: Newsletter": {
      reminderTexts: ['TWA.io'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createTwaNewsletter(triggerText), get.channel.twaNewsletter, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Andy');
      }
    },
    "TWA.io: Newsletter": {
      reminderTexts: ['Social'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createSocial(triggerText), get.channel.social, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
      }
    },
    "TWA.io: Blog": {
      reminderTexts: ['Medium', 'Social', 'Associates'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createMedium(triggerText), get.channel.mediumBlog, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
        executeFeature(get.featureId.createInPublishing,   [get.message.createSocial(triggerText), get.channel.social,     get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
        executeFeature(get.featureId.createInCurrentPaige, [get.message.contactAssociates(channel, triggerText), get.defaultStatus.currentPaige, get.message.contactHow, get.message.contactWhenDone, '', get.message.generatedBy(channel)]);
        executeFeature(get.featureId.setSheetStylesBySectionPaige);
      }
    },
    "TWA.io: Project": {
      reminderTexts: ['Social', 'Associates'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createSocial(triggerText), get.channel.social, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
        executeFeature(get.featureId.createInCurrentPaige, [get.message.contactAssociates(channel, triggerText), get.defaultStatus.currentPaige, get.message.contactHow, get.message.contactWhenDone, '', get.message.generatedBy(channel)]);
        executeFeature(get.featureId.setSheetStylesBySectionPaige);
      }
    },
    "AAH.io: Blog": {
      reminderTexts: ['Shares', 'Associates'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createSocial(triggerText), get.channel.social, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
        executeFeature(get.featureId.createInCurrentPaige, [get.message.contactAssociates(channel, triggerText), get.defaultStatus.currentPaige, get.message.contactHow, get.message.contactWhenDone, '', get.message.generatedBy(channel)]);
        executeFeature(get.featureId.setSheetStylesBySectionPaige);
      }
    },
    "AAH.io: Bio": {
      reminderTexts: ['Associates'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInCurrentPaige, [get.message.contactAssociates(channel, triggerText), get.defaultStatus.currentPaige, get.message.contactHow, get.message.contactWhenDone, '', get.message.generatedBy(channel)]);
        executeFeature(get.featureId.setSheetStylesBySectionPaige);
      }
    },
    "AAH.io: Project(s)": {
      reminderTexts: ['Shares', 'Associates'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.createSocial(triggerText), get.channel.social, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Paige');
        executeFeature(get.featureId.createInCurrentPaige, [get.message.contactAssociates(channel, triggerText), get.defaultStatus.currentPaige, get.message.contactHow, get.message.contactWhenDone, '', get.message.generatedBy(channel)]);
        executeFeature(get.featureId.setSheetStylesBySectionPaige);
      }
    },
    "Internal: Newsletter": {
      reminderTexts: ['Stakeholders'],
      executeFeatures: (channel, triggerText) => {
        executeFeature(get.featureId.createInPublishing,   [get.message.contactInternal(channel, triggerText), get.channel.internalStakeholder, get.defaultStatus.publishing, '', get.message.generatedBy(channel)], 'Andy');
      }
    }
  }

  const getReadableReminders = (reminders, newLine='\n', bullet='• ') => {
    return bullet + reminders.join(newLine + bullet);
  };

  const getAllReadableRemindersAsHtml = () => {
    let html = '<table>';
    for(const channel in publicationResponses) {
      html += '<tr><td><pre>' + channel.replace(': ', '<br>') + '</pre></td><td style="padding-left: 2.8em; padding-right: 1.8em;"><pre>' + getReadableReminders(publicationResponses[channel].reminderTexts, '<br>') + '</pre></td><tr>';
    }
    return html + '</table>';
  };

  return {
    name: 'Publishing',
    features: {
      cacheSheetRowsBySection: {
        events: [Event.onSheetEdit],
        section: SectionMarker.main,
        cacheKey: 'Publishing.main'
      },
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetGroupsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.done,
        numRowsToDisplay: 3
      },
      alertSheetOnEdit: {
        events: [Event.onSheetEdit],
        priority: 'HIGH_PRIORITY',
        cacheKey: 'Publishing.main',
        triggerValue: '(4) PUBLISHED',
        buttonSet: 'YES_NO',
        getMessage: (row) => {
          const channel = row['C'];
          const response = publicationResponses[channel];
          if(isObject(response)) {
            const reminders = getReadableReminders(publicationResponses[channel].reminderTexts);
            return { title: 'Publication Triggers', text: 'You just published a ' + channel + ', which should trigger the following:\n\n' + reminders + '\n\nWould you like to create these entries now?' };
          }
          return false;
        },
        respondToPrompt(button, row) {
          const channel = row['C'];
          let responses = publicationResponses[channel].executeFeatures(channel, row['B']);
        }
      },
      createSheetItem: {
        id: 'createSheetItem.Publishing'
      }
    },
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Publishing'
      },
      review: getReviewConfig(),
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
      create: {
        type: 'buttons',
        title: 'Create new',
        options: ['Paige', 'Andy'],
        features: {
          createSheetItem: {
            events: [Event.onSidebarSubmit],
            priority: 'HIGH_PRIORITY',
            getInsertionSectionIndex: (sectionName) => {
              return get.publishingSectionIndex[sectionName];
            },
            getValues: () => {
              return ['', '', '(0) Not Started', '', ''];
            }
          },
          setSheetStylesBySection: {
            events: [Event.onSidebarSubmit],
            styles: styles
          },
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Archive Published Items'],
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
      triggers: {
        type: 'text',
        title: 'Publication Triggers',
        text: getAllReadableRemindersAsHtml()
      },
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
      heading: {
        type: 'heading',
        title: 'Projects'
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
  styles.titles.after = state.style.getBlank({ endColumnOffset: 0, numColumns: 6, border: state.style.border.empty });

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
      heading: {
        type: 'heading',
        title: 'Hands'
      },
      review: getReviewConfig(SectionMarker.title, 'I'),
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
        options: ['Archive Lowered Hands'],
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
      }
    }
  };
}

function getStakeholdersConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'done', 'underMain', 'underDone', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getTwoPanel(sections, 8);
  styles.contents.all.rowHeight = 44;
  styles.titles.between.endColumnOffset = 3;
  styles.titles.review.endColumnOffset = 2;
  styles.titles.after = state.style.getBlank({ endColumnOffset: 0, numColumns: 2, border: state.style.border.empty });

  return {
    name: 'Stakeholders',
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
      heading: {
        type: 'heading',
        title: 'Stakeholders'
      },
      review: getReviewConfig(SectionMarker.title, 'I'),
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Type', 'Office', 'Country'],
        features: {
          orderSheetMainSections: {
            events: [Event.onSidebarSubmit],
            by: {
              type:  [{ column: 'G', direction: 'ascending' }, { column: 'F', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
              office:  [{ column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }, { column: 'G', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
              country: [{ column: 'F', direction: 'ascending' }, { column: 'G', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'B', direction: 'ascending' }]
            }
          }
        }
      },
      archive: {
        type: 'buttons',
        title: 'Tidy',
        options: ['Move Inactive to Archive'],
        features: {
          moveSheetRowsToDone: {
            events: [Event.onSidebarSubmit],
            from: SectionMarker.main,
            match: {
              value: ') Exited / Inactive',
              column: 'H'
            }
          }
        }
      }
    }
  };
}

function getNetworkConfig() {
  return {
    name: 'Network',
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Network'
      },
      review: getReviewConfig(SectionMarker.title, 'C')
    }
  };
}

function getSpendConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'generic', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getTwoPanel(sections);
  styles.titles.between.endColumnOffset = 2;
  styles.titles.review.endColumnOffset = 1;
  styles.headers.all.verticalAlignment = PropertyCommand.IGNORE;
  styles.headers.all.border = PropertyCommand.IGNORE;
  styles.headers.left = { beginColumnOffset: 0, numColumns: 1, verticalAlignment: state.style.alignment.vertical.middle };
  styles.headers.right = { beginColumnOffset: 1, verticalAlignment: state.style.alignment.vertical.bottom };
  styles.contents.all.rowHeight = 33;
  styles.contents.all.fontSize = PropertyCommand.IGNORE;
  styles.contents.all.border = PropertyCommand.IGNORE;
  styles.contents.left.fontSize = 12;
  styles.contents.right.border = PropertyCommand.IGNORE;
  styles.rowsOutside.all.border = PropertyCommand.IGNORE;
  styles.columnsOutside.all.border = PropertyCommand.IGNORE;

  return {
    name: 'Spend',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      },
      setSheetHiddenRowsBySection: {
        events: [Event.onOvernightTimer],
        section: SectionMarker.generic,
        startRowOffset: -1,
        visibilityMatcher: {
          column: 'B',
          text: state.today.getFullYear()
        }
      }
    },
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Spend'
      },
      review: getReviewConfig(SectionMarker.title, 'G'),
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
            visibilityMatcher: {
              column: 'B',
              text: PropertyCommand.EVENT_DATA
            }
          }
        }
      }
    }
  };
}

function getStreamsConfig() {
  const sections = ['titles', 'titlesAboveBelow', 'headers', 'main', 'underMain', 'rowsOutside', 'columnsOutside'];
  const styles = state.style.getFourPanel(sections, 1, 1, 3);
  styles.contents.all.rowHeight = 68;
  styles.contents.left.fontSize = 11;
  styles.contents.right.fontSize = 9;
  styles.contents.middleAim = { beginColumnOffset: 3, numColumns: 1, border: { top: null, left: true, bottom: null, right: true, vertical: null, horizontal: null, color: '#ffffff', style: 'SOLID' }};
  styles.titles.between.endColumnOffset = 2;
  styles.titles.review.endColumnOffset = 1;
  styles.titles.after = state.style.getBlank({ endColumnOffset: 0, numColumns: 1, border: state.style.border.empty });

  return {
    name: 'Streams',
    features: {
      setSheetStylesBySection: {
        events: [Event.onSheetEdit, Event.onOvernightTimer, Event.onHourTimer],
        styles: styles
      }
    },
    sidebar: {
      heading: {
        type: 'heading',
        title: 'Streams'
      },
      review: getReviewConfig(SectionMarker.title, 'G'),
      arrange: {
        type: 'buttons',
        title: 'Arrange by',
        options: ['Stream', 'Aims'],
        features: {
          orderSheetMainSections: {
            events: [Event.onSidebarSubmit],
            by: {
              stream: [{ column: 'B', direction: 'ascending' }, { column: 'D', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }],
              aims: [{ column: 'D', direction: 'ascending' }, { column: 'E', direction: 'ascending' }, { column: 'F', direction: 'ascending' }, { column: 'B', direction: 'ascending' }],
            }
          }
        }
      }
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
      heading: {
        type: 'heading',
        title: 'Aims'
      },
      review: getReviewConfig()
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
      heading: {
        type: 'heading',
        title: 'Docs'
      },
      review: getReviewConfig()
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