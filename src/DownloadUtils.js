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

function exportFormulasAsJson(deleteTmpResourcesCallbackFnName) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();
  const values = range.getValues();

  const json = [];
  for (let row = 0; row < formulas.length; row++) {
    for (let col = 0; col < formulas[row].length; col++) {
      if (formulas[row][col]) {
        json.push({
          cell: String.fromCharCode(65 + col) + (row + 1),
          formula: formulas[row][col],
          value: values[row][col]
        });
      }
    }
  }

  const jsonString = JSON.stringify(json, null, 2);
  const blob = Blob.createTextBlob(jsonString, 'application/json');
  const file = DriveApp.createFile(blob).setName('formulas.json');
  downloadFile(file, deleteTmpResourcesCallbackFnName);
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
    .setProperty(`trigger_${trigger.getUniqueId()}`, copy.getId());
}