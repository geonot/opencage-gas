/**
 * OpenCage Geocoder for Google Apps Script
 *
 * A library for interacting with the OpenCage Geocoding API, designed
 * for easy use within Google Sheets and other Google Workspace products.
 *
 * @version 1.0.0
 * @license MIT
 * @see https://github.com/geonot/opencage-gas
 */

// =================================================================
// SECTION 1: CORE LIBRARY
// =================================================================

/**
 * Main library object to encapsulate geocoding logic.
 */
const OpenCageGeocoder = {
  
  _BASE_URL: 'https://api.opencagedata.com/geocode/v1/json',
  _LIBRARY_VERSION: '1.0.0',

  /**
   * The core internal function for making requests to the OpenCage API.
   * @private
   * @param {Object} params - The query parameters for the API call.
   * @returns {Object} The parsed JSON response from the API.
   * @throws {Error} If API key is missing or if the API returns an error.
   */
  _request: function(params) {
    // 1. Get API key from Script Properties. This is the secure way.
    const apiKey = PropertiesService.getUserProperties().getProperty('OPENCAGE_API_KEY');
    
    // 2. Do not make a request without a valid key.
    if (!apiKey) {
      throw new Error('API key not found. Please set OPENCAGE_API_KEY in "Project Settings > Script Properties".');
    }
    params.key = apiKey;

    // 3. Build the URL with query parameters.
    const queryString = Object.keys(params)
      .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
      .join('&');
    const url = this._BASE_URL + '?' + queryString;

    // 4. Set request options, including the custom User-Agent.
    const options = {
      'method': 'get',
      'headers': {
        'User-Agent': 'OpenCage-GoogleAppsScript/' + this._LIBRARY_VERSION
      },
      'muteHttpExceptions': true // IMPORTANT: This allows us to handle non-200 responses.
    };

    let response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (e) {
      throw new Error('Network error calling OpenCage API: ' + e.message);
    }

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const jsonResponse = JSON.parse(responseBody);
    
    // 5. Handle response codes correctly.
    if (responseCode === 200) {
      // Handle case where request is valid but no results are returned.
      // The `results` array will simply be empty, so we just return the full response.
      return jsonResponse;
    } else {
      // For any non-200 code, get the error message from the API response and throw it.
      const errorMessage = (jsonResponse.status && jsonResponse.status.message) 
        ? jsonResponse.status.message 
        : 'Unknown API Error (HTTP ' + responseCode + ')';
      
      // REQUIREMENT: Stop on 402/403. Throwing an error accomplishes this.
      throw new Error(errorMessage);
    }
  },

  /**
   * Performs forward geocoding (address to coordinates).
   * @param {string} query The address or place name to geocode.
   * @param {Object} [options] Optional parameters like 'language', 'countrycode', etc.
   * @returns {Object} The full API response object.
   */
  forward: function(query, options) {
    if (!query) throw new Error('Query cannot be empty for forward geocoding.');
    
    const params = options || {};
    params.q = query;
    return this._request(params);
  },

  /**
   * Performs reverse geocoding (coordinates to address).
   * @param {number} latitude The latitude.
   * @param {number} longitude The longitude.
   * @param {Object} [options] Optional parameters like 'language'.
   * @returns {Object} The full API response object.
   */
  reverse: function(latitude, longitude, options) {
    if (typeof latitude !== 'number' || typeof longitude !== 'number') {
      throw new Error('Latitude and Longitude must be numbers for reverse geocoding.');
    }
    
    const params = options || {};
    params.q = latitude + ',' + longitude;
    return this._request(params);
  }
};


// =================================================================
// SECTION 2: GOOGLE SHEETS INTEGRATION
// =================================================================

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 * @private
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📍 OpenCage Geocoder')
    .addItem('Geocode Selected Range', 'geocodeSelectedRange')
    .addToUi();
}

/**
 * A menu-triggered function to geocode a selected column of addresses.
 * It reads from the selected column and writes results to the adjacent columns.
 */
function geocodeSelectedRange() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const addresses = range.getValues();
  const ui = SpreadsheetApp.getUi();

  const results = [];
  
  ui.alert('Geocoding Started', `Processing ${addresses.length} addresses. This may take a moment...`, ui.ButtonSet.OK);

  for (let i = 0; i < addresses.length; i++) {
    const address = addresses[i][0];
    let output = ['','','']; // [Formatted Address, Lat, Lng]
    
    if (address) { // Only process non-empty cells
      try {
        const response = OpenCageGeocoder.forward(address);
        if (response.results.length > 0) {
          const topResult = response.results[0];
          output[0] = topResult.formatted;
          output[1] = topResult.geometry.lat;
          output[2] = topResult.geometry.lng;
        } else {
          output[0] = 'No results found';
        }
      } catch (e) {
        output[0] = 'ERROR: ' + e.message;
      }
    }
    results.push(output);
  }
  
  // Write the results to the three columns to the right of the selection
  sheet.getRange(range.getRow(), range.getColumn() + 1, results.length, 3).setValues(results);
  
  ui.alert('Geocoding Complete', 'Finished processing all selected addresses.', ui.ButtonSet.OK);
}


// =================================================================
// SECTION 3: TESTING
// Test your library from the Script Editor by selecting `runTests` and clicking "Run".
// View results in "Execution log".
// =================================================================

/**
 * A wrapper function to run all tests.
 * IMPORTANT: You must set a VALID API key in Script Properties for some tests to pass.
 */
function runAllTests() {
  Logger.log('--- STARTING OPENCAGE TEST SUITE ---');
  testForwardGeocoding_Success();
  testReverseGeocoding_Success();
  test_NoResults();
  testErrorHandling_401_InvalidKey();
  testErrorHandling_402_PaymentRequired();
  Logger.log('--- TEST SUITE COMPLETE ---');
}

function testForwardGeocoding_Success() {
  try {
    const response = OpenCageGeocoder.forward('Brandenburg Gate, Berlin');
    if (response.results.length > 0 && response.results[0].components.country_code === 'de') {
      Logger.log('PASS: Forward Geocoding');
    } else {
      Logger.log('FAIL: Forward Geocoding - Unexpected response');
    }
  } catch(e) {
    Logger.log('FAIL: Forward Geocoding - ' + e.message);
  }
}

function testReverseGeocoding_Success() {
  try {
    const response = OpenCageGeocoder.reverse(41.4036, 2.1744);
    if (response.results.length > 0 && response.results[0].components.road.includes('Sagrada')) {
      Logger.log('PASS: Reverse Geocoding');
    } else {
      Logger.log('FAIL: Reverse Geocoding - Unexpected response');
    }
  } catch(e) {
    Logger.log('FAIL: Reverse Geocoding - ' + e.message);
  }
}

function test_NoResults() {
  try {
    // REQUIREMENT: test query that returns 0 results
    const response = OpenCageGeocoder.forward('NOWHERE-INTERESTING'); 
    if (response.results.length === 0) {
      Logger.log('PASS: No Results Test');
    } else {
      Logger.log('FAIL: No Results Test - Expected 0 results, got ' + response.results.length);
    }
  } catch(e) {
    Logger.log('FAIL: No Results Test - ' + e.message);
  }
}

function testErrorHandling_401_InvalidKey() {
  // REQUIREMENT: Use test keys
  PropertiesService.getUserProperties().setProperty('OPENCAGE_API_KEY', '62b364a687554b32b50939634e9d6d5a'); // Invalid key
  try {
    OpenCageGeocoder.forward('London');
    Logger.log('FAIL: 401 Error Test - Did not throw an error.');
  } catch(e) {
    if (e.message.includes('invalid API key')) {
      Logger.log('PASS: 401 Error Test');
    } else {
      Logger.log('FAIL: 401 Error Test - Wrong error message: ' + e.message);
    }
  }
}

function testErrorHandling_402_PaymentRequired() {
  // REQUIREMENT: Use test keys
  PropertiesService.getUserProperties().setProperty('OPENCAGE_API_KEY', '4372eff77b8343cebfc843eb4da4ddc4'); // Key that triggers 402
  try {
    OpenCageGeocoder.forward('London');
    Logger.log('FAIL: 402 Error Test - Did not throw an error.');
  } catch(e) {
    if (e.message.includes('quota exceeded')) {
      Logger.log('PASS: 402 Error Test');
    } else {
      Logger.log('FAIL: 402 Error Test - Wrong error message: ' + e.message);
    }
  }
}
