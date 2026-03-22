const LOG_ALERTS_PROPERTY_NAME = 'logAlerts';

const LOG_LEVEL = 'LOG';
const INFO_LEVEL = 'INFO';
const WARN_LEVEL = 'WARN';
const ERROR_LEVEL = 'ERROR';

/**
 * Logs a message with file and function context.
 *
 * @param {string} fileName Source file name.
 * @param {string} functionName Function name.
 * @param {*} message Message to log.
 * @param {string} [level] Log level.
 */
function log(fileName, functionName, message, level = LOG_LEVEL) {
  const fullMessage = buildLogMessage(fileName, functionName, message);
  _log(fullMessage, level);
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
function logArgs(fileName, functionName, args, message = null, level = LOG_LEVEL) {
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
 * @private
 */
function _log(message, level = LOG_LEVEL) {
  switch (level) {
    case ERROR_LEVEL:
      console.error(message);
      break;
    case WARN_LEVEL:
      console.warn(message);
      break;
    case INFO_LEVEL:
      console.info(message);
      break;
    case LOG_LEVEL:
      console.log(message);
      break;
    default:
      console.warn('Unknown log level: ' + level + '. Message: ' + message);
  }

  if (getLogAlerts()) {
    alert('Log', message);
  }
}

/**
 * Toggles the user setting for showing alert logs.
 * When true, logs are displayed in an alert dialog.
 */
function toggleAlertLogs() {
  if (getLogAlerts()) {
    PropertiesService.getUserProperties().deleteProperty(LOG_ALERTS_PROPERTY_NAME);
  } else {
    PropertiesService.getUserProperties().setProperty(LOG_ALERTS_PROPERTY_NAME, 'true');
  }
}

/**
 * Returns true if alert logs are enabled.
 *
 * @returns {boolean}
 */
function getLogAlerts() {
  return PropertiesService.getUserProperties().getProperty(LOG_ALERTS_PROPERTY_NAME) === 'true';
}