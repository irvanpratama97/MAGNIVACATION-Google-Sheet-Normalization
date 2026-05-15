# MAGNIVACATION-Google-Sheet-Normalization
Automated data normalization in Google Sheets using Apps Script to standardize formats and improve data quality for analysis.

# 📌 Project Overview
This project demonstrates how I **Normalized Data** in Google Sheets using **Google Apps Script**.  
The script ensures that datasets are **clean, consistent, and ready for analysis** by:
- Standardizing date formats (e.g., `YYYY-MM-DD`)
- Unmerging the merged cells
- Removing duplicate rows
- Trimming extra spaces
- etc.

---

## 🛠 Tools & Technologies
- **Google Apps Script** (JavaScript-based automation)
- **Google Sheets**

## 🚀 How It Works
1. **User uploads or pastes raw data** into a Google Sheet.
2. **Run the script** from the Extensions → Apps Script menu.
3. The script:
   - Cleans and formats the data
   - Removes duplicates
   - Outputs a normalized dataset in the same sheet

---

## 📂 Code
```javascript
/**
 * Function to normalize and merge data from monthly sheets into a single Master Data sheet.
 * This script handles merged cells logic and skips "TOTAL" rows.
 */
function normalizeAndMergeData() {
  // 1. Setup Source and Target Spreadsheet
  // Replace 'ID_FILE_SOURCE_DISINI' with the actual ID of your (SOURCE) MAGNIVACATION file
  const sourceSsId = '[Spreadsheet ID here]'; 
  const sourceSs = SpreadsheetApp.openById(sourceSsId);
  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSs.getSheetByName('MASTER DATA');
  
  // List of monthly tabs to process in the source file
  const monthTabs = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  
  let masterData = [];
  let globalNo = 1;

  // 2. Loop through each monthly tab
  monthTabs.forEach(monthName => {
    const sheet = sourceSs.getSheetByName(monthName);
    if (!sheet) return; // Skip if the sheet name is not found

    const values = sheet.getDataRange().getValues();
    
    // Variables to store current values for "filling down" merged cell data
    let currentRegDate = "";
    let currentPackage = "";
    let currentMonth = "";
    let currentDay = "";
    let currentYear = 2025; // Defaulting to 2025 based on your screenshot

    // Start loop from row 4 (assuming row 1-3 are headers)
    for (let i = 3; i < values.length; i++) {
      let row = values[i];
      let applicantId = row[2]; // Column C: ID Applicant

      // A. Check if the row is a "TOTAL" row or empty - if so, skip it
      if (!applicantId || applicantId.toString().toUpperCase().includes("TOTAL")) {
        continue;
      }

      // B. Handle "Fill Down" logic for merged/grouped columns
      // If Register Date (Col B) is not empty, update current date values
      if (row[1] !== "") {
        currentRegDate = row[1]; // Example: "1 January"
        let parts = currentRegDate.toString().split(" ");
        currentDay = parts[0];
        currentMonth = parts[1];
      }

      // If Package (Col D) is not empty, update current package
      if (row[3] !== "") {
        currentPackage = row[3];
      }

      // C. Construct the normalized row for target:
      // [No, Date, Month, Year, Notes, Applicant ID, Package, Number of Participants, Phone Number, Pickup Address]
      let normalizedRow = [
        globalNo++,                                      // No
        new Date(currentYear, getMonthIndex(currentMonth), currentDay), // Date (Object)
        currentMonth,                                    // Month
        currentYear,                                     // Year
        "",                                              // Notes (Empty)
        applicantId,                                     // Applicant ID
        currentPackage,                                  // Package
        row[4],                                          // Number of Participants (Col E)
        row[5],                                          // Phone Number (Col F)
        row[6]                                           // Pickup Address (Col G)
      ];

      masterData.push(normalizedRow);
    }
  });

  // 3. Write data to the Target Sheet
  if (masterData.length > 0) {
    // Clear existing data from row 2 downwards
    if (targetSheet.getLastRow() > 1) {
      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
    }
    
    // Set the values to the target sheet
    targetSheet.getRange(2, 1, masterData.length, masterData[0].length).setValues(masterData);
    
    // Optional: Set date format for Column B
    targetSheet.getRange(2, 2, masterData.length, 1).setNumberFormat('d-mmm-yyyy');
    
    SpreadsheetApp.getUi().alert('Data successfully normalized and merged!');
  }
}

/**
 * Helper function to convert month name to index (0-11)
 */
function getMonthIndex(monthName) {
  const months = {
    'January': 0, 'February': 1, 'March': 2, 'April': 3, 'May': 4, 'June': 5,
    'July': 6, 'August': 7, 'September': 8, 'October': 9, 'November': 10, 'December': 11,
    'Januari': 0, 'Februari': 1, 'Maret': 2, 'Mei': 4, 'Juni': 5, 'Juli': 6, 'Agustus': 7, 'Oktober': 9, 'Desember': 11
  };
  return months[monthName] || 0;
}
