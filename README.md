# OpenCage Geocoder for Google Apps Script

A Google Apps Script library for interacting with the [OpenCage Geocoding API](https://opencagedata.com). This library makes it easy to forward and reverse geocode directly within your Google Sheets and other Google Workspace projects.

[![Build Status](https://github.com/geonot/opencage-gas/actions/workflows/main.yml/badge.svg)](https://github.com/geonot/opencage-gas/actions)

## Features

-   ✅ Simple forward and reverse geocoding functions.
-   ✅ Securely stores your API key using `PropertiesService` (not in code).
-   ✅ Helper functions to geocode an entire column of addresses from a custom menu in Google Sheets.
-   ✅ Follows all OpenCage best practices, including error handling and custom user-agent.
-   ✅ No external dependencies.

## Prerequisites

-   A Google Account (for using Google Sheets).
-   An [OpenCage API key](https://opencagedata.com/api-key). The free trial is perfect to get started.

## Setup Instructions

1.  **Open your Google Sheet.**
2.  Go to `Extensions > Apps Script`.
3.  Delete any code in the `Code.gs` file and **paste the entire contents** of the `OpenCage.gs` file from this repository.
4.  Save the project (click the floppy disk icon). Give it a name like "OpenCage".
5.  **Set your API Key:**
    -   In the Apps Script editor, click the **Project Settings** icon (a cog wheel ⚙️ on the left).
    -   Scroll down to **Script Properties** and click **Add script property**.
    -   For **Property**, enter `OPENCAGE_API_KEY`.
    -   For **Value**, paste your actual OpenCage API key.
    -   Click **Save script properties**.

    

6.  Reload your Google Sheet. A new "📍 OpenCage Geocoder" menu should appear.

## Usage

### Method 1: Geocode a Column of Addresses (Recommended)

This is the easiest and most powerful way to use the library.

1.  Enter your addresses in a single column (e.g., Column A).
2.  Select the cells you want to geocode.
3.  Go to the custom menu `📍 OpenCage Geocoder > Geocode Selected Range`.
4.  The script will process the selected addresses and write the formatted address, latitude, and longitude in the columns to the right.

### Method 2: Use as a Library in your own Scripts

You can also call the geocoder from your own Apps Script functions.

#### Forward Geocoding (Address to Coordinates)

```javascript
function myFunction() {
  try {
    // Basic request
    const response = OpenCageGeocoder.forward('Eiffel Tower, Paris');
    
    // Request with optional parameters
    const options = {
      language: 'fr',
      countrycode: 'fr'
    };
    const responseWithOptions = OpenCageGeocoder.forward('Arc de Triomphe', options);

    if (response.results.length > 0) {
      const topResult = response.results[0];
      Logger.log(topResult.formatted); // "Tour Eiffel, 5 Avenue Anatole France, 75007 Paris, France"
      Logger.log(topResult.geometry.lat); // 48.85824
      Logger.log(topResult.geometry.lng); // 2.2945
    }

  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
```

#### Reverse Geocoding (Coordinates to Address)

```javascript
function myReverseFunction() {
  try {
    const response = OpenCageGeocoder.reverse(40.71427, -74.00597);

    // What happens when there are no results
    const noResultsResponse = OpenCageGeocoder.forward('NOWHERE-INTERESTING');
    Logger.log('Number of results: ' + noResultsResponse.results.length); // Prints 0

  } catch (e) {
    // What happens when the API returns an error (e.g., invalid key)
    Logger.log('Error: ' + e.message);
  }
}
```

### Best Practices

For best results, please review OpenCage's advice on [how to format forward geocoding queries](https://opencagedata.com/api#forward-best-results). A more specific query will always yield better results.

### License

This project is licensed under the MIT License. See the `LICENSE` file for details.
