/**
 * Returns a pretty-printed JSON string.
 *
 * @param {*} obj Object to stringify.
 * @returns {string} JSON string.
 */
function jsonStringify(obj) {
  return JSON.stringify(obj, null, 2);
}

/**
 * Sends a POST request with JSON payload and returns parsed response.
 *
 * @param {string} url Request URL.
 * @param {*} payload Request body.
 * @returns {*} Parsed JSON response.
 */
function post(url, payload) {
  logArgs('Utils', 'post', {payloadStr: JSON.stringify(payload)}, url);
  const res = UrlFetchApp.fetch(
    url,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );

  return JSON.parse(res.getContentText());
}

/**
 * Displays an alert dialog, or logs if UI is unavailable.
 *
 * @param {string} title Alert title.
 * @param {string} message Alert message.
 */
function alert(title, message) {
  try {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    log('alert', 'alert', message);
  }
}

/**
 * Converts an array to an object using a property-to-index map.
 *
 * @param {Array<*>} arr Source array.
 * @param {Object<string, number>} propNameToIndexMap Map of property names to indexes.
 * @returns {Object} Result object.
 */
function arrayToObj(arr, propNameToIndexMap) {
  const obj = {};

  Object.entries(propNameToIndexMap).forEach(([propName, index]) => {
    if (index >= arr.length) return;
    obj[toCamelCase(propName)] = arr[index];
  });

  return obj;
}

/**
 * Converts a string to camelCase.
 *
 * @param {string} str Input string.
 * @returns {string} camelCase string.
 */
function toCamelCase(str) {
  return str
    .match(/[A-Za-z0-9]+/g)
    .map((w, i) => i === 0
      ? w.toLowerCase()
      : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
    )
    .join('');
}

/**
 * Capitalizes a string (first letter or all words).
 *
 * @param {string} str Input string.
 * @param {boolean} [allWords=true] Whether to capitalize all words.
 * @returns {string} Capitalized string.
 */
function capitalize(str, allWords = true) {
  if (!str) return '';

  if (!allWords) {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  return str
    .split(/\s+/)
    .map(w => w ? w.charAt(0).toUpperCase() + w.slice(1) : '')
    .join(' ');
}