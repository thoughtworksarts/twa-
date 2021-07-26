var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  projectSync: {
    sourceTab: 'Projects',
    destinationSheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
    destinationTab: 'Current Projects'
  },
  toggles: {
    performDataUpdates: true,
    showLogAlert: false
  }
}

function getNameSubstitution(name) {
  return name;
}

function preProcessSubsheets() {
  state.personValuesSubsheet = new PersonValuesSubsheet(state.spreadsheet, '(workings)', { start:'G3', end:'G5' });
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

  state.eventSubsheets.push(new EventSubsheet(state.spreadsheet, 'Todo:Andy', '630855359', range, sections, triggerCols));
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
  var sourceData = sourceTab.getDataRange();
  var destinationSheet = SpreadsheetApp.openById(config.projectSync.destinationSheetID);
  var destinationTab = destinationSheet.getSheetByName(config.projectSync.destinationTab);
  var destinationRange = destinationTab.getRange(1, 1, sourceData.getNumRows(), sourceData.getNumColumns())
  destinationRange.setValues(sourceData.getValues());
}
