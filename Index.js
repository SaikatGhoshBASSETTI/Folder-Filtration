
//Helper Functions


//Retrieves the configuration object by merging default presets with settings from the 'Main_Info' sheet. */

function getConfig() {

  const config = {

    headers: [
      'Module Name',
      'Module Number',
      'RefDB',
      'TC',
      'Script',
      'Type',
      'Version',
      'Link',
      'ModuleID'
    ],

    moduleName: 'E3',
    refDb: 'E8',
    tc: 'C2',
    scripts: 'C7',
    type: 'C7',
    version: 'C7',
    moduleId: 'C7'
  };

  const infoSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Main_Info');

  const data = infoSheet
    .getDataRange()
    .getValues();

  for (let i = 1; i < data.length; i++) {

    const key = String(data[i][0]).trim();
    const value = String(data[i][1]).trim();

    if (key) {
      config[key] = value;
    }
  }

  return config;
}

//Safely retrieves a trimmed display value from a specified cell, returning an empty string if an error occurs. */

function getCellValue(sheet, cellReference) {
  try {
    return sheet.getRange(cellReference).getDisplayValue().trim();
  } catch (error) {
    Logger.log(`Error reading ${cellReference}: ${error}`);
    return '';
  }
}

//Extracts and trims the prefix before the first period from a given string identifier. */

function getModuleNumber(name) {
  
  return name.split('.')[0].trim();

}

// Generates a Google Sheets `=HYPERLINK` formula using the provided URL and display text. */

function createHyperlink(url, text = 'Link') {
  return `=HYPERLINK("${url}","${text}")`;
}

//Extracts the version number (e.g., '7.1') from a folder name string based on predefined patterns. */

function getVersionFromFolderName(folderName) {

  if (!folderName) {
    return '';
  }
  if (/trunk/i.test(folderName)) {
    return '7.1';
  }
  const match = folderName.match(/\d+\.\d+/);

  return match ? String(match[0]) : '';
}

function processSheet(file, folder, output, config) {

  try {

    const spreadsheet = SpreadsheetApp.openById(
      file.getId()
    );

    const sheet = spreadsheet.getSheets()[0];

    const moduleName = getCellValue(
      sheet,
      config.moduleName
    );

    const moduleNumber = getModuleNumber(
      file.getName()
    );

    const refDb = getCellValue(
      sheet,
      config.refDb
    );

    const totalTC = getCellValue(
      sheet,
      config.tc
    );

    const scripts = getCellValue(
      sheet,
      config.scripts
    );

    const type = folder.getName();

    const parentfolder = folder.getParents().next();
    const version = getVersionFromFolderName(parentfolder.getName());


    const moduleId = getCellValue(
      sheet,
      config.moduleId
    );

    output.push([
      moduleName,
      moduleNumber,
      refDb,
      totalTC,
      scripts,
      type,
      version,
      file.getUrl(),       
      moduleId
    ]);

    Logger.log(
      `Processed: ${file.getName()}`
    );

  } catch (error) {

    Logger.log(
      `Error processing ${file.getName()}: ${error}`
    );
  }
}


function traverseFolders(folder, output, config) {
  const folderName = folder.getName().toLowerCase();

  if (['archive', 'draft'].some(keyword => folderName.includes(keyword))) {
    Logger.log(`Skipped Folder: ${folder.getName()}`);
    return;
  }

  Logger.log(`Entering Folder: ${folder.getName()}`);

  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  while (files.hasNext()) {
    processSheet(files.next(), folder, output, config);
  }

  const subFolders = folder.getFolders();

  while (subFolders.hasNext()) {
    traverseFolders(subFolders.next(), output, config);
  }

  Logger.log(`Leaving Folder: ${folder.getName()}`);
}


//compare sheets
function compareWithReferenceSheet(referenceSpreadsheetId) {

  const config = getConfig();

  try {

    Logger.log('========== COMPARISON STARTED ==========');

    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const sourceSheet = activeSpreadsheet.getSheetByName(
      config.CONSOLIDATED_SHEET_NAME
    );

    const outputSheet = activeSpreadsheet.getSheetByName(
      config.OUTPUT_SHEET_NAME
    );

    const referenceSheet = SpreadsheetApp
      .openById(referenceSpreadsheetId)
      .getSheets()[0];

    outputSheet.clearContents();

    Logger.log(
      `Source Sheet: ${sourceSheet.getName()}`
    );

    Logger.log(
      `Reference Sheet: ${referenceSheet.getName()}`
    );

    const sourceData =
      sourceSheet.getDataRange().getValues();

    const referenceData =
      referenceSheet.getDataRange().getValues();

    Logger.log(
      `Source Rows: ${sourceData.length - 1}`
    );

    Logger.log(
      `Reference Rows: ${referenceData.length - 1}`
    );

    const output = [sourceData[0]];
    const highlightCells = [];

    // Build lookup map from reference sheet
    const referenceMap = new Map();

    for (let i = 1; i < referenceData.length; i++) {

      const moduleName =
        String(referenceData[i][0]).trim();

      const version =
        String(referenceData[i][6]).trim();

      const key =
        `${moduleName}|${version}`;

      referenceMap.set(
        key,
        referenceData[i]
      );
    }

    Logger.log(
      `Reference Lookup Created : ${referenceMap.size}`
    );

    // Compare source rows
    for (let i = 1; i < sourceData.length; i++) {

      const sourceRow = sourceData[i];

      const moduleName =
        String(sourceRow[0]).trim();

      const version =
        String(sourceRow[6]).trim();

      const key =
        `${moduleName}|${version}`;

      Logger.log(
        `Checking : ${key}`
      );

      // Record not found
      if (!referenceMap.has(key)) {

        Logger.log(
          `NOT FOUND : ${key}`
        );

        continue;
      }

      const referenceRow =
        referenceMap.get(key);

      const moduleMatch =
        String(sourceRow[0]).trim() ===
        String(referenceRow[0]).trim();

      const refDbMatch =
        String(sourceRow[2]).trim() ===
        String(referenceRow[2]).trim();

      const linkMatch =
        String(sourceRow[7]).trim() ===
        String(referenceRow[7]).trim();

      if (
        moduleMatch &&
        refDbMatch &&
        linkMatch
      ) {

        Logger.log(
          `MATCHED : ${key}`
        );

        continue;
      }

      Logger.log(
        `MISMATCH FOUND : ${key}`
      );

      const outputRow =
        [...referenceRow];

      output.push(outputRow);

      const outputRowNumber =
        output.length;

      const mismatchDetails = [];

      // Module Name mismatch
      if (!moduleMatch) {

        highlightCells.push({
          row: outputRowNumber,
          col: 1
        });

        mismatchDetails.push(
          config.A01_COLUMN_MATCHED
        );
      }

      // RefDB mismatch
      if (!refDbMatch) {

        highlightCells.push({
          row: outputRowNumber,
          col: 3
        });

        mismatchDetails.push(
          config.A02_COLUMN_MATCHED
        );
      }

      // Link mismatch
      if (!linkMatch) {

        highlightCells.push({
          row: outputRowNumber,
          col: 8
        });

        mismatchDetails.push(
          config.A03_COLUMN_MATCHED
        );
      }

      Logger.log(
        `Mismatch Columns : ${mismatchDetails.join(', ')}`
      );
    }

    // Write output
    if (output.length > 1) {

      outputSheet
        .getRange(
          1,
          1,
          output.length,
          output[0].length
        )
        .setValues(output);

      Logger.log(
        `Rows Written To Sheet2 : ${output.length - 1}`
      );

      // Highlight mismatched cells
      highlightCells.forEach(cell => {

        outputSheet
          .getRange(
            cell.row,
            cell.col
          )
          .setBackground(config.MISMATCHED_BG_COLOR);

      });

      Logger.log(
        `Highlighted Cells : ${highlightCells.length}`
      );

    } else {

      Logger.log(
        'No Mismatches Found'
      );
    }

    Logger.log(
      '========== COMPARISON COMPLETED =========='
    );

  } catch (error) {

    Logger.log(
      `ERROR : ${error}`
    );

    throw error;
  }
} 

//Fetch All Data Function

function allDataFetch() {
  const configs = getConfig();

  try {

    const ROOT_FOLDER_ID = configs.ROOT_FOLDER_ID;

    const config = getConfig();

    const output = [config.headers];

    const rootFolder = DriveApp.getFolderById(
      ROOT_FOLDER_ID
    );

    Logger.log(
      `Starting traversal from: ${rootFolder.getName()}`
    );

    traverseFolders(
      rootFolder,
      output,
      config
    );

    const masterSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(configs.CONSOLIDATED_SHEET_NAME);

    masterSheet.clearContents();

    masterSheet
      .getRange(
        1,
        1,
        output.length,
        output[0].length
      )
      .setValues(output);

    Logger.log(
      `Successfully processed ${output.length - 1} records`
    );

   

  } catch (error) {

    Logger.log(
      `allDataFetch Error: ${error}`
    );

    throw error;
  }
}

//////////////////////////////////////////
function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('TP Automation')

    .addItem(
      'Fetch All Data',
      'menuAllDataFetch'
    )

    .addItem(
      'Compare Sheets',
      'menuCompareSheets'
    )

    .addSeparator()

    .addItem(
      'Exit',
      'menuExit'
    )

    .addToUi();
}

function menuAllDataFetch() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Fetch Data',
    'Do you want to start data extraction?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  allDataFetch();

  ui.alert(
    'Success',
    'Data extraction completed.',
    ui.ButtonSet.OK
  );
}

function menuCompareSheets() {
  
  const config_1 = getConfig();

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Compare Sheets',
    'Enter Reference Spreadsheet ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const spreadsheetId =
    response.getResponseText().trim();

  compareWithReferenceSheet(
    spreadsheetId
  );

  ui.alert(
    'Success',
    'Comparison completed.',
    ui.ButtonSet.OK
  );
}
