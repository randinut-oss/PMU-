// ================= CONFIG =================

var FIRST_DATA_ROW_COMPILED = 4;        // Compiled_List: start at row 4
var FIRST_DATA_COL_COMPILED = 2;         // Compiled_List: start at column B (index 2)

var FIRST_DATA_ROW_LMS = 3;              // LMS: start at row 3 (A3)
var FIRST_DATA_COL_LMS = 1;               // LMS: start at column A
var LMS_NUM_COLS = 12;                    // LMS: 12 columns (A to L)

var SOURCE_SHEET_NAMES = ['FOBM', 'FOCIT', 'FOS', 'FOT', 'FOE', 'DOM'];


// ================= AUTO UPDATE =================

function onEdit(e) {
  if (!e) return;

  var sheetName = e.range.getSheet().getName();
  if (SOURCE_SHEET_NAMES.indexOf(sheetName) === -1) return;

  Utilities.sleep(300);
  runFullSystem();
}


// ================= MENU =================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Module Tools')
    .addItem('🔄 Full Compile & LMS Update', 'runFullSystem')
    .addToUi();
}


// ================= MAIN CONTROLLER =================

function runFullSystem() {
  try {
    var result = compileModules();
    if (!result) {
      SpreadsheetApp.getActive().toast('No data found in source sheets', '⚠️ Warning', 3);
      return;
    }

    updateCompiledList(result.allOutput, result.mergeGroups);
    updateLMS(result.allOutput, result.mergeGroups);

    SpreadsheetApp.getActive().toast('✅ System updated successfully!', 'Success', 3);
  } catch (error) {
    SpreadsheetApp.getActive().toast('Error: ' + error.message, '❌ Error', 5);
    console.error(error);
  }
}


// ================= READ + MERGE LOGIC =================

function compileModules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mergeGroups = {};
  var standaloneRows = [];

  SOURCE_SHEET_NAMES.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet not found: ' + sheetName);
      return;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 3) {
      console.log('Sheet ' + sheetName + ' has no data');
      return;
    }

    // Start from row index 2 (third row) – data starts at row 3 in source sheets
    for (var r = 2; r < data.length; r++) {
      var row = data[r];

      var courseCode = (row[7] || '').toString().trim();
      if (!courseCode) continue;

      var mergedIdRaw = (row[1] || '').toString().trim();
      if (!mergedIdRaw || mergedIdRaw === '') mergedIdRaw = 'NO';

      // Clean data
      for (var c = 0; c < row.length; c++) {
        if (row[c] instanceof Date) {
          row[c] = Utilities.formatDate(row[c], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof row[c] === 'string') {
          row[c] = row[c].trim();
        }
      }

      var cleanRow = {
        mergedId: mergedIdRaw,
        faculty: row[2] || '',
        program: row[3] || '',
        degree: row[4] || '',
        batch: row[5] || '',
        semester: row[6] || '',
        courseCode: row[7] || '',
        courseName: row[8] || '',
        credits: row[9] || '',
        coreElective: row[10] || '',
        owner: row[11] || '',
        tas: row[12] || ''
      };

      if (mergedIdRaw.toUpperCase() === 'NO') {
        standaloneRows.push(cleanRow);
      } else {
        if (!mergeGroups[mergedIdRaw]) mergeGroups[mergedIdRaw] = [];
        mergeGroups[mergedIdRaw].push(cleanRow);
      }
    }
  });

  // Build merged rows (for Compiled_List only)
  var mergedRows = [];
  for (var id in mergeGroups) {
    var group = mergeGroups[id];

    function unique(field) {
      var set = {};
      group.forEach(function(r) {
        if (r[field]) set[r[field]] = true;
      });
      return Object.keys(set);
    }

    function combine(field) {
      var u = unique(field);
      if (u.length === 0) return '';
      if (u.length === 1) return u[0];
      return u.join(' / ');
    }

    mergedRows.push({
      mergedId: id,
      faculty: combine('faculty'),
      program: combine('program'),
      degree: combine('degree'),
      batch: combine('batch'),
      semester: combine('semester'),
      courseCode: combine('courseCode'),
      courseName: combine('courseName'),
      credits: combine('credits'),
      coreElective: combine('coreElective'),
      owner: combine('owner'),
      tas: combine('tas')
    });
  }

  var standaloneOutput = standaloneRows.map(function(r) {
    var newRow = Object.assign({}, r);
    newRow.mergedId = 'NO';
    return newRow;
  });

  var allOutput = standaloneOutput.concat(mergedRows);

  // Sort by course name
  allOutput.sort(function(a, b) {
    return (a.courseName || '').localeCompare(b.courseName || '');
  });

  console.log('Total rows compiled: ' + allOutput.length);
  console.log('Merge groups: ' + Object.keys(mergeGroups).length);

  return { allOutput: allOutput, mergeGroups: mergeGroups };
}


// ================= UPDATE COMPILED =================

function updateCompiledList(allOutput, mergeGroups) {
  if (!allOutput || !Array.isArray(allOutput) || allOutput.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var compiledSheet = ss.getSheetByName('Compiled_List');
  if (!compiledSheet) {
    compiledSheet = ss.insertSheet('Compiled_List');
    // Add headers
    var headers = [['Merged ID', 'Faculty', 'Program', 'Degree Name', 'Batch Number',
                    'Semester', 'Course Code', 'Course Name', 'No of Credits',
                    'Core/Elective', 'Name of the Course Owner']];
    compiledSheet.getRange(1, 2, 1, 11).setValues(headers);
    compiledSheet.getRange(1, 2, 1, 11).setFontWeight('bold');
    compiledSheet.getRange(1, 2, 1, 11).setHorizontalAlignment('center');
  }

  var numCols = 11; // B to L
  var startRow = FIRST_DATA_ROW_COMPILED; // 4
  var startCol = FIRST_DATA_COL_COMPILED; // 2 (B)

  // Clear from B4 downward
  var lastRow = compiledSheet.getLastRow();
  if (lastRow >= startRow) {
    var clearRange = compiledSheet.getRange(
      startRow,
      startCol,
      lastRow - startRow + 1,
      numCols
    );
    clearRange.clearContent();
    clearRange.clearFormat();
  }

  // Build output
  var output = allOutput.map(function(row) {
    var out = new Array(numCols).fill('');
    out[0] = row.mergedId || '';              // B: Merged ID
    out[1] = row.faculty || '';                // C: Faculty
    out[2] = row.program || '';                // D: Program
    out[3] = row.degree || '';                 // E: Degree Name
    out[4] = row.batch || '';                  // F: Batch Number
    out[5] = row.semester || '';               // G: Semester
    out[6] = row.courseCode || '';             // H: Course Code
    out[7] = row.courseName || '';             // I: Course Name
    out[8] = row.credits || '';                // J: No of Credits
    out[9] = row.coreElective || '';           // K: Core/Elective
    out[10] = row.owner || '';                 // L: Name of the Course Owner
    return out;
  });

  // Write data
  var targetRange = compiledSheet.getRange(
    startRow,
    startCol,
    output.length,
    numCols
  );
  targetRange.setValues(output);

  // Apply formatting
  targetRange.setHorizontalAlignment('center');
  targetRange.setWrap(true);
  targetRange.setVerticalAlignment('middle');

  // Highlight merged rows
  var bg = allOutput.map(function(r) {
    var isMerged = r.mergedId !== 'NO' &&
                   mergeGroups && mergeGroups.hasOwnProperty(r.mergedId) &&
                   mergeGroups[r.mergedId].length > 1;
    return new Array(numCols).fill(isMerged ? '#ADD8E6' : null);
  });
  targetRange.setBackgrounds(bg);

  console.log('Compiled_List updated with ' + output.length + ' rows');
}


// ================= UPDATE LMS =================
// Build LMS directly from faculty-level rows
// Merge ONLY: A (Merged ID), B (Merged Or No), I (LMS Title), J (LMS Link)

function updateLMS(dataRows, mergeGroups) {
  // Safety: ensure mergeGroups is an object
  if (!mergeGroups || typeof mergeGroups !== 'object') mergeGroups = {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lmsSheet = ss.getSheetByName('LMS');

  if (!lmsSheet) {
    lmsSheet = ss.insertSheet('LMS');

    var headers = [[
      'Merged ID', 'Merged Or No', 'Semester', 'Degree Name',
      'Batch Number', 'C/E', 'Credit Value', 'Program',
      'LMS Title', 'LMS Link', 'Code', 'Module'
    ]];

    lmsSheet.getRange(1, 1, 1, 12).setValues(headers)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    lmsSheet.setFrozenRows(1);
  }

  var startRow = FIRST_DATA_ROW_LMS; // 3

  // Clear old data + remove merges
  var maxRows = lmsSheet.getMaxRows();
  if (maxRows >= startRow) {
    var clearRange = lmsSheet.getRange(startRow, 1, maxRows - startRow + 1, LMS_NUM_COLS);
    clearRange.breakApart();
    clearRange.clearContent();
    clearRange.clearFormat();
  }

  // 🔹 STEP 1: Rebuild faculty-level rows from mergeGroups
  var facultyRows = [];

  // Standalone rows (mergedId = NO) from dataRows
  if (dataRows && Array.isArray(dataRows)) {
    dataRows.forEach(function(row) {
      if (row.mergedId === 'NO') {
        facultyRows.push(row);
      }
    });
  }

  // Add original grouped rows (not combined) from mergeGroups
  for (var id in mergeGroups) {
    if (mergeGroups.hasOwnProperty(id) && Array.isArray(mergeGroups[id])) {
      mergeGroups[id].forEach(function(r) {
        facultyRows.push(r);
      });
    }
  }

  if (facultyRows.length === 0) return;

  // 🔹 STEP 2: Sort by merged ID to keep groups together
  facultyRows.sort(function(a, b) {
    if (a.mergedId === b.mergedId) {
      return (a.courseName || '').localeCompare(b.courseName || '');
    }
    if (a.mergedId === 'NO') return 1;
    if (b.mergedId === 'NO') return -1;
    return a.mergedId.localeCompare(b.mergedId);
  });

  var output = [];
  var mergeRanges = [];
  var currentRow = startRow;

  for (var i = 0; i < facultyRows.length; i++) {
    var row = facultyRows[i];
    var groupStartRow = currentRow;
    var groupSize = 1;

    var isRealMerged = row.mergedId !== 'NO';

    // Count how many same merged ID rows in sequence
    if (isRealMerged) {
      for (var j = i + 1; j < facultyRows.length; j++) {
        if (facultyRows[j].mergedId === row.mergedId) {
          groupSize++;
        } else {
          break;
        }
      }
    }

    for (var g = 0; g < groupSize; g++) {
      var r = facultyRows[i + g];
      var out = new Array(12).fill('');

      if (g === 0) {
        out[0] = isRealMerged ? r.mergedId : '';
        out[1] = isRealMerged ? 'Yes' : 'No';
        out[8] = ''; // LMS Title
        out[9] = ''; // LMS Link

        if (isRealMerged && groupSize > 1) {
          mergeRanges.push({row: groupStartRow, col: 1, size: groupSize});
          mergeRanges.push({row: groupStartRow, col: 2, size: groupSize});
          mergeRanges.push({row: groupStartRow, col: 9, size: groupSize});
          mergeRanges.push({row: groupStartRow, col: 10, size: groupSize});
        }
      }

      // Always show module data (no merge)
      out[2] = r.semester || '';
      out[3] = r.degree || '';
      out[4] = r.batch || '';
      out[5] = r.coreElective || '';
      out[6] = r.credits || '';
      out[7] = r.program || '';
      out[10] = r.courseCode || '';
      out[11] = r.courseName || '';

      output.push(out);
      currentRow++;
    }

    i += (groupSize - 1);
  }

  if (output.length > 0) {
    var target = lmsSheet.getRange(startRow, 1, output.length, LMS_NUM_COLS);
    target.setValues(output)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true);

    // Apply merges
    mergeRanges.forEach(function(m) {
      lmsSheet.getRange(m.row, m.col, m.size, 1).merge();
    });

    SpreadsheetApp.getActive().toast('LMS sheet updated (' + output.length + ' rows)', '✅ LMS', 2);
  }
}
