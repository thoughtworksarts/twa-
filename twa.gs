var config = {
  gsheet: {
    name: 'twa—',
    id: '1jvrit8ybRObHLCh3hNBdQnANssKyLy6UR0SFTTCIusY'
  },
  projectSync: {
    handsTab: 'Hands',
    sourceTab: 'Projects',
    destinationSheetID: '1UJMpl988DHsl3FSgZU4VoXysaKolK-IrzNz_xxbSguM',
    destinationTab: 'Current Projects',
    assignmentRow: 5,
    assignmentCol: 8
  },
  toggles: {
    performDataUpdates: true,
    showLogAlert: false
  }
}

function customOnEdit() {
  if(state.spreadsheet.getActiveSheet().getName() === config.projectSync.sourceTab || state.spreadsheet.getActiveSheet().getName() === config.projectSync.handsTab) {
    syncProjects();
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

  state.eventSubsheets.push(new EventSubsheet(state.spreadsheet, 'Director:Todo', '630855359', range, sections, triggerCols));
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
  var destinationSheet = SpreadsheetApp.openById(config.projectSync.destinationSheetID);
  var destinationTab = destinationSheet.getSheetByName(config.projectSync.destinationTab);

  overwriteProjectsWithRichTextValues(sourceTab, destinationTab);
  copyProjectsAssignmentValues(sourceTab, destinationTab);
}

function overwriteProjectsWithRichTextValues(sourceTab, destinationTab) {
  var sourceRange = sourceTab.getRange(1, 1, sourceTab.getMaxRows(), sourceTab.getMaxColumns());
  var destinationRange = destinationTab.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());
  destinationRange.clearContent();
  destinationRange.setRichTextValues(sourceRange.getRichTextValues());
}

function copyProjectsAssignmentValues(sourceTab, destinationTab) {
  var numRows = sourceTab.getMaxRows() - config.projectSync.assignmentRow;
  var sourceRange = sourceTab.getRange(config.projectSync.assignmentRow, config.projectSync.assignmentCol, numRows, 1);
  var destinationRange = destinationTab.getRange(5, 8, numRows, 1);
  destinationRange.setValues(sourceRange.getValues());
}