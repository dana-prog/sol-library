const WRITE_TO_LOG_FILE_PROPERTY_NAME = 'writeToLogFile';
const LOG_SPREADSHEET_ID = '1dauoW72DKPJaMIhwTuCEa7tUqikiW5o4Rkcu1NHd_z0';
const LOG_SHEET_ID = 0;

const LOG_LEVELS = {
  LOG: 'LOG',
  INFO: 'INFO',
  WARN: 'WARN',
  ERROR: 'ERROR',
};

let _logSheet;

/**
 * Logs a message with file and function context.
 *
 * @param {string} fileName Source file name.
 * @param {string} functionName Function name.
 * @param {*} message Message to log.
 * @param {string} [level] Log level.
 */
function log(fileName, functionName, message, level = LOG_LEVELS.LOG) {
  const fullMessage = buildLogMessage(fileName, functionName, message);

  switch (level) {
    case LOG_LEVELS.ERROR:
      console.error(fullMessage);
      break;
    case LOG_LEVELS.WARN:
      console.warn(fullMessage);
      break;
    case LOG_LEVELS.INFO:
      console.info(fullMessage);
      break;
    case LOG_LEVELS.LOG:
      console.log(fullMessage);
      break;
    default:
      console.warn('Unknown log level: ' + level + '. Message: ' + fullMessage);
  }

  if (getWriteToLogFileEnabled()) {
    _writeToLogSheet(fileName, functionName, message, level)
  }
}

/**
 * Logs arguments (object) with optional message.
 *
 * @param {string} fileName Source file name.
 * @param {string} functionName Function name.
 * @param {Object} args Arguments object.
 * @param {string|null} [message=null] Optional message prefix.
 * @param {string} [level] Log level.
 */
function logArgs(fileName, functionName, args, message = null, level = LOG_LEVELS.LOG) {
  log(fileName, functionName, (message != null ? message + '\n' : '') + jsonStringify(args), level);
}

/**
 * Builds a formatted log message with file and function context.
 *
 * @param {string} fileName
 * @param {string} functionName
 * @param {string} message
 * @returns {string}
 */
function buildLogMessage(fileName, functionName, message) {
  return `[${fileName}::${functionName}]\n${message}`;
}

/**
 * Toggles the user setting for showing alert logs.
 * When true, logs are displayed in an alert dialog.
 */
function toggleWriteToLogFile() {
  if (getWriteToLogFileEnabled()) {
    PropertiesService.getDocumentProperties().deleteProperty(WRITE_TO_LOG_FILE_PROPERTY_NAME);
  } else {
    PropertiesService.getDocumentProperties().setProperty(WRITE_TO_LOG_FILE_PROPERTY_NAME, 'true');
  }
}

/**
 * Returns true if alert logs are enabled.
 *
 * @returns {boolean}
 */
function getWriteToLogFileEnabled() {
  return PropertiesService.getDocumentProperties().getProperty(WRITE_TO_LOG_FILE_PROPERTY_NAME) === 'true';
}

function _getLogSheet() {
  if (!_logSheet) {
    _logSheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID).getSheetById(LOG_SHEET_ID);
  }

  return _logSheet;
}

function _writeToLogSheet(fileName, functionName, message, level) {
  const logSheet = _getLogSheet();
  const row = logSheet.getLastRow() + 1;

  // HARD-CODED COLUMNS: 1=TIMESTAMP, 2=LEVEL, 3=FILE, 4=FUNCTION, 5=MESSAGE
  logSheet.getRange(row, 1, 1, 5).setValues([[new Date().toISOString(), level, fileName, functionName, message]]);
}