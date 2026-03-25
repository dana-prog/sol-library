declare const SOLLibrary: {
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
 exportValuesXSLX(deleteTmpResourcesCallbackFnName: any): void;

/**
 * Deletes temporary export resources created by exportValuesXSLX:
 * - Trashes the copied file
 * - Removes the associated script property
 * - Deletes the time-based trigger
 *
 * @param {GoogleAppsScript.Events.TimeDriven} event
 * @returns {void}
 */
 deleteTmpExportResources(event: GoogleAppsScript.Events.TimeDriven): void;

/**
 * Logs a message with file and function context.
 *
 * @param {string} fileName Source file name.
 * @param {string} functionName Function name.
 * @param {*} message Message to log.
 * @param {string} [level] Log level.
 */
 log(fileName: string, functionName: string, message: any, level?: string): void;

/**
 * Logs arguments (object) with optional message.
 *
 * @param {string} fileName Source file name.
 * @param {string} functionName Function name.
 * @param {Object} args Arguments object.
 * @param {string|null} [message=null] Optional message prefix.
 * @param {string} [level] Log level.
 */
 logArgs(fileName: string, functionName: string, args: any, message?: string | null, level?: string): void;

/**
 * Builds a formatted log message with file and function context.
 *
 * @param {string} fileName
 * @param {string} functionName
 * @param {string} message
 * @returns {string}
 */
 buildLogMessage(fileName: string, functionName: string, message: string): string;


/**
 * Toggles the user setting for showing alert logs.
 * When true, logs are displayed in an alert dialog.
 */
 toggleAlertLogs(): void;

/**
 * Returns true if alert logs are enabled.
 *
 * @returns {boolean}
 */
 getLogAlerts(): boolean;

  LOG_ALERTS_PROPERTY_NAME: "logAlerts";

  LOG_LEVEL: "LOG";

  INFO_LEVEL: "INFO";

  WARN_LEVEL: "WARN";

  ERROR_LEVEL: "ERROR";

/**
 * Returns a sheet by name, optionally creating it if it does not exist.
 *
 * @param {string} sheetName Sheet name.
 * @param {boolean} [create=false] Whether to create the sheet if missing.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet or null.
 */
 getSheet(sheetName: string, create?: boolean): GoogleAppsScript.Spreadsheet.Sheet | null;

/**
 * Returns the column number for the given header name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|GoogleAppsScript.Spreadsheet.Range} rangeOrSheet Sheet or range.
 * @param {string} colHeader Column header name.
 * @returns {number} Column number (1-based).
 */
 getColNumByHeader(rangeOrSheet: GoogleAppsScript.Spreadsheet.Sheet | GoogleAppsScript.Spreadsheet.Range, colHeader: string): number;

/**
 * Appends multiple rows to the sheet.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<Array<*>>} rows Rows to append.
 */
 appendRowsToSheet(sheetName: string, rows: Array<Array<any>>): void;

/**
 * Sorts the sheet by a column.
 *
 * @param {string} sheetName Sheet name.
 * @param {string} colName Column header name.
 * @param {boolean} [ascending=true] Sort order.
 */
 sortSheet(sheetName: string, colName: string, ascending?: boolean): void;

/**
 * Returns all values from a column.
 *
 * @param {string|GoogleAppsScript.Spreadsheet.Sheet|GoogleAppsScript.Spreadsheet.Range} sheetOrRange Either sheet or sheet name or a range.
 * @param {string|number} colHeader Column header name or index.
 * @param {boolean} [includeHeader=true] Whether to include header row.
 * @returns {Array<*>} Column values.
 */
 getColumnValues(sheetOrRange: string | GoogleAppsScript.Spreadsheet.Sheet | GoogleAppsScript.Spreadsheet.Range, colHeader: string | number, includeHeader?: boolean): Array<any>;

/**
 * Returns an object mapping headers to values for a row.
 *
 * @param {string} sheetName Sheet name.
 * @param {number} rowNum Row number (1-based).
 * @returns {Object} Row values keyed by header.
 */
 getRowValues(sheetName: string, rowNum: number): any;

/**
 * Returns the currently selected row number.
 * @param {string} sheetName Sheet name.
 *
 * @returns {number} Row number (1-based).
 */
 getSelectedRow(sheetName: string): number;

/**
 * Returns a map of header to column index (0-based).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|GoogleAppsScript.Spreadsheet.Range} rangeOrSheet Sheet or range.
 * @returns {Object<string, number>|-1} Header map or -1 if empty.
 */
 getHeaderMap(rangeOrSheet: GoogleAppsScript.Spreadsheet.Sheet | GoogleAppsScript.Spreadsheet.Range): {
    [x: string]: number;

} | -1;

/**
 * Converts objects to sheet rows based on headers.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<Object>} objects Objects to convert (prop names must match column headers).
 * @returns {Array<Array<*>>} Rows aligned with sheet headers.
 */
 generateSheetRows(sheetName: string, objects: Array<any>): Array<Array<any>>;

/**
 * Selects a row and scrolls it into view.
 *
 * @param {string} sheetName Sheet name.
 * @param {number} rowNum Row number (1-based).
 */
 selectRow(sheetName: string, rowNum: number): void;

/**
 * Appends a row, optionally runs a callback, and selects it.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<*>} row Row values.
 * @param {Function|null} [callback=null] Callback with (sheet, newRowNum).
 */
 addRow(sheetName: string, row: Array<any>, callback?: Function | null): void;

/**
 * Updates an existing row in the given sheet with the provided values.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<*>} row - Row values.
 * @param {number} rowNum - The row number to update.
 */
 updateRow(sheetName: string, row: Array<any>, rowNum: number): void;

/**
 * Returns all row numbers where the specified column equals the given value (excluding header).
 *
 * @param {string} sheetName Sheet name.
 * @param {string} colHeader Column header name.
 * @param {*} value Value to match.
 * @returns {number[]} Matching row numbers (1-based).
 */
 getRowNumbers(sheetName: string, colHeader: string, value: any): number[];

/**
 * Returns the range of a column by header or index.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|string} sheetNameOrObj Either sheet or sheet name.
 * @param {string|number} colHeader Column header name or index.
 * @param {boolean} [includeHeader=true] Whether to include header row.
 * @returns {GoogleAppsScript.Spreadsheet.Range} Column range.
 */
 getColumnRange(sheetNameOrObj: GoogleAppsScript.Spreadsheet.Sheet | string, colHeader: string | number, includeHeader?: boolean): GoogleAppsScript.Spreadsheet.Range;

/**
 * Returns a pretty-printed JSON string.
 *
 * @param {*} obj Object to stringify.
 * @returns {string} JSON string.
 */
 jsonStringify(obj: any): string;

/**
 * Sends a POST request with JSON payload and returns parsed response.
 *
 * @param {string} url Request URL.
 * @param {*} payload Request body.
 * @returns {*} Parsed JSON response.
 */
 post(url: string, payload: any): any;

/**
 * Displays an alert dialog, or logs if UI is unavailable.
 *
 * @param {string} title Alert title.
 * @param {string} message Alert message.
 */
 alert(title: string, message: string): void;

/**
 * Converts an array to an object using a property-to-index map.
 *
 * @param {Array<*>} arr Source array.
 * @param {Object<string, number>} propNameToIndexMap Map of property names to indexes.
 * @returns {Object} Result object.
 */
 arrayToObj(arr: Array<any>, propNameToIndexMap: {
    [x: string]: number;

}): any;

/**
 * Converts a string to camelCase.
 *
 * @param {string} str Input string.
 * @returns {string} camelCase string.
 */
 toCamelCase(str: string): string;

/**
 * Capitalizes a string (first letter or all words).
 *
 * @param {string} str Input string.
 * @param {boolean} [allWords=true] Whether to capitalize all words.
 * @returns {string} Capitalized string.
 */
 capitalize(str: string, allWords?: boolean): string;


};
