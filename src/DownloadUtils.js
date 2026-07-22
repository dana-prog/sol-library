/**
 * Creates a copy of the active spreadsheet with values only (no formulas),
 * removes private sheets (names starting with "_"), and triggers download as XLSX.
 * The temporary copy is deleted automatically after a short delay.
 *
 * NOTE: this function installs a trigger that deletes the temporary file/the property and the trigger itself after 1 minute.
 * deleteTmpExportResources deletes those temporary resources.
 * however, when the trigger executes it will look for the function (by name) in the calling project and not in the library
 * Therefore a callback function with the name passed to the trigger should be defined in the calling project and can (and should)
 * delegate the implementation to deleteTmpExportResources defined here in the library.
 * So the code in the calling project should look like:
 *
 * const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;
 * SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
 *
 * @returns {void}
 */
function exportValuesXSLX(deleteTmpResourcesCallbackFnName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const copy = spreadsheet.copy(spreadsheet.getName() + "_values_only_" + timestamp);

  // delete private sheets
  copy.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith('_')) {
      copy.deleteSheet(sheet);
    }
  });

  // replace formulas with values
  copy.getSheets().forEach(s => {
    const range = s.getDataRange();
    range.copyTo(range, {contentsOnly: true});
  });

  downloadFile(copy, deleteTmpResourcesCallbackFnName);
}

/**
 * Creates a JSON file containing all cell formulas in the active spreadsheet.
 * @param deleteTmpResourcesCallbackFnName
 */
function exportFormulasJSON(deleteTmpResourcesCallbackFnName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

  // Extract formulas from all non-private sheets
  const allFormulas = {};
  spreadsheet.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith('_')) {
      const range = sheet.getDataRange();
      const formulas = range.getFormulas();
      const values = range.getValues();

      const sheetFormulas = [];
      for (let row = 0; row < formulas.length; row++) {
        for (let col = 0; col < formulas[row].length; col++) {
          if (formulas[row][col]) {
            sheetFormulas.push({
              cell: String.fromCharCode(65 + col) + (row + 1),
              formula: formulas[row][col],
              value: values[row][col]
            });
          }
        }
      }
      allFormulas[sheet.getName()] = sheetFormulas;
    }
  });

  // Create temp JSON file
  const jsonBlob = Utilities.newBlob(JSON.stringify(allFormulas, null, 2), 'application/json', `formulas_${timestamp}.json`);
  const file = DriveApp.createFile(jsonBlob);

  downloadFile(file, deleteTmpResourcesCallbackFnName);
}

/**
 * Creates a JSON file containing all named functions in the calling project.
 * @param deleteTmpResourcesCallbackFnName
 */
function exportNamedFunctionsJSON(deleteTmpResourcesCallbackFnName) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

  const functions = _dumpNamedFunctions(spreadsheetId);

  if (functions.error) {
    Logger.log('Error: ' + functions.error);
    return;
  }

  const jsonBlob = Utilities.newBlob(
    JSON.stringify(functions, null, 2),
    'application/json',
    `named_functions_${timestamp}.json`
  );
  const file = DriveApp.createFile(jsonBlob);

  downloadFile(file, deleteTmpResourcesCallbackFnName);
}

/**
 * Named Functions are not exposed by SpreadsheetApp or the Sheets REST API.
 * The xlsx export endpoint preserves them as <definedName> elements inside
 * xl/workbook.xml (LAMBDA-serialized). We fetch that export and parse it.
 * Source: https://gist.github.com/tanaikech/9a9e571ed662e35eec0aa747bb4e025a
 */
function _dumpNamedFunctions(spreadsheetId) {
  try {
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=${spreadsheetId}`;
    const response = UrlFetchApp.fetch(url, {
      headers: { authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      return { error: `export http ${response.getResponseCode()}` };
    }

    const blobs = Utilities.unzip(response.getBlob().setContentType(MimeType.ZIP));
    const workbookBlob = blobs.find((blob) => blob.getName() === 'xl/workbook.xml');
    if (!workbookBlob) {
      return { error: 'xl/workbook.xml not found in export' };
    }

    const root = XmlService.parse(workbookBlob.getDataAsString()).getRootElement();
    const definedNamesElement = root.getChild('definedNames', root.getNamespace());
    if (!definedNamesElement) {
      return [];
    }

    return definedNamesElement.getChildren().map((element) => ({
      name: element.getAttribute('name').getValue(),
      definition: element.getValue(),
    }));
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Deletes temporary export resources created by exportValuesXSLX:
 * - Trashes the copied file
 * - Removes the associated script property
 * - Deletes the time-based trigger
 *
 * @param {GoogleAppsScript.Events.TimeDriven} event
 * @returns {void}
 */
function deleteTmpExportResources(event) {
  const triggerId = event.triggerUid;
  const propertyName = `trigger_${triggerId}`;
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty(propertyName);

  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      log('DownloadUtils', 'deleteTmpExportResources', 'Deleted file: ' + fileId);
    } catch (e) {
      // TODO: resolve exception thrown from: DriveApp.getFileById(fileId)
      log('DownloadUtils', 'deleteTmpExportResources', `Error deleting file:\n${e}`, LOG_LEVEL.ERROR);
    }

    props.deleteProperty(propertyName);
    log('DownloadUtils', 'deleteTmpExportResources', 'Deleted property: ' + propertyName);
  }

  // delete the trigger itself (time-based triggers are persistent and eventually will hit quota limits)
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function downloadFile(file, deleteTmpResourcesCallbackFnName) {
  // url for download
  const url = "https://docs.google.com/spreadsheets/d/" + file.getId() + "/export?format=xlsx";

  // html trigger browser download
  const html = HtmlService
    .createHtmlOutput(`
    <script>
      window.open("${url}", "_blank");
      google.script.host.close();
    </script>
  `)
    .setWidth(10)
    .setHeight(10);

  SpreadsheetApp
    .getUi()
    .showModalDialog(html, "Downloading...");

  const trigger =
    ScriptApp
      .newTrigger(deleteTmpResourcesCallbackFnName)
      .timeBased()
      .after(60 * 1000)
      .create();

  PropertiesService
    .getScriptProperties()
    .setProperty(`trigger_${trigger.getUniqueId()}`, file.getId());
}