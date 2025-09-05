# Pappers SIRET Lookup for Google Sheets

A Google Apps Script that automatically fetches French company information using SIRET numbers via the Pappers API.

## What it does

This script takes SIRET numbers from your Google Sheets and enriches them with company data from the Pappers API, including:
- Full 14-digit SIRET validation
- Company name
- Company details (if available)

## Setup

### 1. Get a Pappers API Key
1. Visit [Pappers API](https://www.pappers.fr/api)
2. Create a free account
3. Get your API key from your dashboard

### 2. Install the Script
1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code
4. Copy and paste the provided script
5. Replace `'YOURAPIKEY'` with your actual Pappers API key:
   ```javascript
   const PAPPERS_API_KEY = 'your_actual_api_key_here';
   ```
6. Save the project (Ctrl+S or Cmd+S)
7. Give it a meaningful name like "SIRET-Pappers-Lookup"

### 3. Authorize the Script
1. Run any function (like `testRow2`) once
2. Google will ask for permissions - accept them
3. The script can now access your spreadsheet and make API calls

## Sheet Structure

The script expects this column layout:

| Column A | Column B | Column C | Column D | Column E |
|----------|----------|----------|----------|----------|
| company_id | siret | user | siret_extracted | company_name |
| 1 | 32212091600208 | user@email.com |32212091600208 | APPLE FRANCE |

- **Column B**: Your input SIRET numbers (can be 11 or 14 digits)
- **Column D**: Full SIRET returned by the API
- **Column E**: Company name returned by the API

## Usage

### Test Single Rows
```javascript
testRow2()    // Test row 2
testRow3()    // Test row 3
```

### Process Ranges
```javascript
processRows3to4()      // Process rows 3-4
processFirst10Rows()   // Process rows 2-11
processRowRange(5, 20) // Process rows 5-20
```

### Process All Data
```javascript
processAllSIRETs()     // Process all rows with data
```

### Process Selected Rows
1. Select the rows you want to process in your sheet
2. Run `processSelectedRows()`

### Process Marked Rows
1. Put "SIRET" in column F for rows you want to process
2. Run `processMarkedRows()`
3. Processed rows will be marked with "✔"

### Auto-processing (Optional)
The script can automatically process rows when you edit column B. This is enabled by default via the `onEdit` trigger.

## Functions Reference

### Core Functions
- `processSingleRow(row)` - Process a specific row number
- `processRowRange(startRow, endRow)` - Process a range of rows
- `processAllSIRETs()` - Process all rows with SIRET data

### Utility Functions
- `testRow2()`, `testRow3()` - Quick tests for specific rows
- `processFirst10Rows()` - Process rows 2-11
- `processSelectedRows()` - Process currently selected rows
- `processMarkedRows()` - Process rows marked with "SIRET"
- `setupHeaders()` - Set up column headers

## Error Handling

The script handles various error scenarios:
- **"SIRET invalide"** - SIRET is too short or contains non-digits
- **"Entreprise non trouvée"** - API couldn't find the company
- **"Erreur API: [code]"** - API returned an error code
- **"Erreur: [message]"** - Network or parsing error

## Rate Limiting

The script includes automatic delays (300ms) between API calls to respect the Pappers API rate limits and avoid being blocked.

## Troubleshooting

### Script not running automatically
- Check that you've authorized the script properly
- The `onEdit` trigger only works when you manually edit cells in column B

### API errors
- Verify your API key is correct
- Check your Pappers API quota/limits
- Ensure your internet connection is stable

### No results found
- Verify the SIRET numbers are correct
- Some companies may not be in the Pappers database
- Try with known valid SIRETs first

## Logging

Check execution logs for debugging:
1. Go to **Extensions > Apps Script**
2. Click **Executions** in the left sidebar
3. Click on any execution to see detailed logs

## Contributing

Feel free to submit issues and pull requests to improve this script.

## License

This project is open source and available under the MIT License.

## Disclaimer

This script uses the Pappers API. Make sure to comply with their terms of service and rate limits. The script is provided as-is without any warranty.
