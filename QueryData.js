function recalculateValues() {
  populateGroupingsTab();
  generateTabularMatrixReport();
}

/**
 * Populates the "Groupings" tab dynamically by extracting Principal, Category, and Type
 * information from the "Monthly Budget" tab.
 *
 * Assumes "Monthly Budget" structure:
 * - Column A: Contains Type headers (e.g., "Income", "Expenses") or "Total " finishers.
 * - Column B: Contains Principal headers (e.g., "Person 1") or "Total " finishers.
 * - Column C: Contains the actual Category names (e.g., "Salary", "Rent", "Groceries").
 * - Blank cells in A, B, or C are used to delineate sections.
 * - The script should stop parsing when it encounters "Total Expenses & Debt Repayment & Savings" in Column A.
 *
 * "Groupings" tab columns: Principal | Categories | Type
 */
function populateGroupingsTab() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const monthlyBudgetSheet = spreadsheet.getSheetByName('Monthly Budget');
  if (!monthlyBudgetSheet) {
    ui.alert('Error', 'The "Monthly Budget" sheet was not found. Please ensure it exists and is named "Monthly Budget".', ui.ButtonSet.OK);
    return;
  }

  let groupingsSheet = spreadsheet.getSheetByName('Groupings');
  if (!groupingsSheet) {
    groupingsSheet = spreadsheet.insertSheet('Groupings');
    Logger.log('Created new "Groupings" sheet.');
  }

  groupingsSheet.clearContents();
  groupingsSheet.clearFormats();

  const START_ROW = 7; // Define the new starting row for data extraction
  const FINAL_TYPE_FINISHER = "total expenses & debt repayment & savings"; // The specific line to stop parsing at

  const lastRow = monthlyBudgetSheet.getLastRow();
  if (lastRow < START_ROW) {
    ui.alert('No Data', `The "Monthly Budget" sheet has no data starting from row ${START_ROW}. No groupings to extract.`, ui.ButtonSet.OK);
    groupingsSheet.getRange(1, 1, 1, 3).setValues([['Principal', 'Categories', 'Type']]);
    return;
  }

  const numRowsToGet = lastRow - START_ROW + 1;
  const budgetData = monthlyBudgetSheet.getRange(START_ROW, 1, numRowsToGet, 3).getValues(); // Reads A, B, C

  Logger.log('--- Starting Groupings Extraction Debug ---');
  Logger.log(`Total rows to process (from sheet row ${START_ROW}): ${budgetData.length}`);

  let currentType = null;
  let currentPrincipal = null;
  let lastSeenType = null; // Track the last valid type seen
  let lastSeenPrincipal = null; // Track the last valid principal seen

  const uniqueGroupings = new Set();
  const seenTypes = new Set(); // Keep track of all types we've seen
  const seenPrincipals = new Set(); // Keep track of all principals we've seen

  for (let i = 0; i < budgetData.length; i++) {
    const sheetRowNum = i + START_ROW; // Actual row number in the sheet for logging
    const row = budgetData[i];
    const cellA = String(row[0]).trim().toLowerCase(); // Column A (Type or section finisher) - normalized
    const cellB = String(row[1]).trim().toLowerCase(); // Column B (Principal or section finisher) - normalized
    const cellC = String(row[2]).trim(); // Column C (Category) - preserve original case
    
    // Keep original case versions for logging and storage
    const originalCellA = String(row[0]).trim();
    const originalCellB = String(row[1]).trim();
    const originalCellC = String(row[2]).trim();

    Logger.log(`Processing Sheet Row ${sheetRowNum}: A='${originalCellA}', B='${originalCellB}', C='${originalCellC}'`);
    Logger.log(`  Current State: Type='${currentType}', Principal='${currentPrincipal}'`);

    // --- Rule 0: Hard Stop at Final Type Finisher ---
    if (cellA === FINAL_TYPE_FINISHER) {
      Logger.log(`  -> Detected final Type finisher: '${originalCellA}'. Stopping parsing.`);
      break;
    }

    // Rule 1: Detect General Section Finishers (e.g., "Total Income", "Total Person 1")
    if (cellA.startsWith("total ") || cellB.startsWith("total ")) {
      Logger.log(`  -> Detected general 'Total' line. Resetting current state but preserving last seen values.`);
      currentType = null;
      currentPrincipal = null;
      continue;
    }

    // Rule 2: Detect Type Header (Column A has value, Column B & C are empty)
    if (originalCellA && !originalCellB && !originalCellC) {
      Logger.log(`  -> Detected new Type: '${originalCellA}'`);
      currentType = originalCellA;
      lastSeenType = originalCellA;
      seenTypes.add(originalCellA);
      currentPrincipal = null; // Reset principal when a new type starts
      continue;
    }

    // Rule 3: Detect Principal Header (Column B has value, Column A & C are empty)
    if (originalCellB && !originalCellA && !originalCellC) {
      Logger.log(`  -> Detected new Principal: '${originalCellB}'`);
      currentPrincipal = originalCellB;
      lastSeenPrincipal = originalCellB;
      seenPrincipals.add(originalCellB);
      continue; // Move to next row after setting principal
    }

    // Rule 4: Detect Category (Column C has value)
    if (originalCellC) {
      // Use current values if available, otherwise fall back to last seen values
      const effectiveType = currentType || lastSeenType;
      const effectivePrincipal = currentPrincipal || lastSeenPrincipal;
      
      Logger.log(`  -> Potential category detected: '${originalCellC}'`);
      Logger.log(`  -> Effective Type: '${effectiveType}', Effective Principal: '${effectivePrincipal}'`);
      
      // Enhanced validation: category should not be a known type or principal
      const isKnownType = seenTypes.has(originalCellC) || originalCellC.toLowerCase().includes('income') || originalCellC.toLowerCase().includes('expense');
      const isKnownPrincipal = seenPrincipals.has(originalCellC) || originalCellC.toLowerCase().includes('person') || originalCellC.toLowerCase().includes('shelter');
      
      if (effectiveType && effectivePrincipal && !isKnownType && !isKnownPrincipal) {
        const groupingKey = `${effectivePrincipal}|${originalCellC}|${effectiveType}`;
        if (!uniqueGroupings.has(groupingKey)) {
          uniqueGroupings.add(groupingKey);
          Logger.log(`  -> Added grouping: ${groupingKey}`);
        } else {
          Logger.log(`  -> Duplicate grouping skipped: ${groupingKey}`);
        }
      } else {
        Logger.log(`  -> Category '${originalCellC}' skipped. Reasons: effectiveType=${!!effectiveType}, effectivePrincipal=${!!effectivePrincipal}, isKnownType=${isKnownType}, isKnownPrincipal=${isKnownPrincipal}`);
      }
    }

    // Rule 5: Handle completely blank rows (reset current state but preserve last seen)
    if (!originalCellA && !originalCellB && !originalCellC) {
      Logger.log(`  -> Blank row detected. Resetting current state.`);
      currentType = null;
      currentPrincipal = null;
      // Don't reset lastSeen values - they persist across blank rows
    }
  }

  Logger.log('--- Finished Groupings Extraction Debug ---');
  Logger.log(`Final unique groupings count: ${uniqueGroupings.size}`);
  Logger.log(`Seen types: ${Array.from(seenTypes).join(', ')}`);
  Logger.log(`Seen principals: ${Array.from(seenPrincipals).join(', ')}`);

  const groupingsData = [['Principal', 'Categories', 'Type']]; // Header row

  uniqueGroupings.forEach(entry => {
    const [principal, category, type] = entry.split('|');
    groupingsData.push([principal, category, type]);
  });

  if (groupingsData.length > 1) { // If there's more than just the header
    groupingsSheet.getRange(1, 1, groupingsData.length, groupingsData[0].length).setValues(groupingsData);
    groupingsSheet.autoResizeColumns(1, groupingsData[0].length);
    ui.alert('Success', `Groupings tab populated with ${groupingsData.length - 1} unique entries.`, ui.ButtonSet.OK);
  } else {
    // Only header or no data found
    groupingsSheet.getRange(1, 1, 1, 3).setValues([['Principal', 'Categories', 'Type']]);
    ui.alert('Info', 'No unique groupings found to populate the "Groupings" tab.', ui.ButtonSet.OK);
  }

  spreadsheet.setActiveSheet(groupingsSheet);
}
/**
 * Generates a flat, tabular matrix report with data grouped by Type, Principal, Category, Year, and Month.
 * Report is ordered descending by Year then Month.
 */
function generateTabularMatrixReport() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Load Groupings Data ---
  const groupingsSheet = spreadsheet.getSheetByName('Groupings');
  if (!groupingsSheet) {
    ui.alert('Error', 'The "Groupings" sheet was not found. Please ensure it exists and is named "Groupings".', ui.ButtonSet.OK);
    return;
  }

  const groupingsLastRow = groupingsSheet.getLastRow();
  // Get columns A, B, C: Principal, Category, Type (assuming data starts from row 2)
  const groupingsValues = groupingsSheet.getRange(2, 1, groupingsLastRow - 1, 3).getValues();

  // Column indices for "Groupings" sheet (0-based)
  const GRP_PRINCIPAL_COL_IDX = 0; // Column A
  const GRP_CATEGORY_COL_IDX = 1;  // Column B
  const GRP_TYPE_COL_IDX = 2;      // Column C

  // Create a map for quick lookup: { 'category_from_Groupings': { principal: 'Principal Name', type: 'Type Name' } }
  const categoryLookup = {};
  for (let i = 0; i < groupingsValues.length; i++) {
    const row = groupingsValues[i];
    const principal = String(row[GRP_PRINCIPAL_COL_IDX]).trim();
    const category = String(row[GRP_CATEGORY_COL_IDX]).trim();
    const type = String(row[GRP_TYPE_COL_IDX]).trim();

    if (principal && category && type) {
      categoryLookup[category.toLowerCase()] = {
        principal: principal,
        type: type
      };
    }
  }

  // --- 2. Load Transactions Data ---
  const transactionsSheet = spreadsheet.getSheetByName('Transactions');
  if (!transactionsSheet) {
    ui.alert('Error', 'The "Transactions" sheet was not found. Please ensure it exists and is named "Transactions".', ui.ButtonSet.OK);
    return;
  }

  const transactionsLastRow = transactionsSheet.getLastRow();
  if (transactionsLastRow < 2) { // Only header row or empty sheet
    ui.alert('No Data', 'No transactions found in the "Transactions" sheet.', ui.ButtonSet.OK);
    return;
  }

  // Get data from relevant columns for "Transactions" sheet:
  // Start reading from Column B (index 1) up to Column F (index 5) to get Date, Category, and Amount.
  // The range will cover: [Date, Description, Group, Category, Amount]
  const transactionsValues = transactionsSheet.getRange(2, 2, transactionsLastRow - 1, 5).getValues();

  // Transaction tab column indices (0-based relative to the *retrieved range* starting at Column B):
  const TRANS_DATE_COL_IDX_IN_RANGE = 0;     // Column B is Date
  const TRANS_CATEGORY_COL_IDX_IN_RANGE = 3; // Column E is Category
  const TRANS_AMOUNT_COL_IDX_IN_RANGE = 4;   // Column F is Amount

  // Data structure to store aggregated results:
  // { 'Principal Name': { 'Type Name': { 'Category Name': { 'YYYY-MM': totalAmount, ... }, ... }, ... } }
  const aggregatedData = {};

  for (let i = 0; i < transactionsValues.length; i++) {
    const row = transactionsValues[i];
    const transactionDate = row[TRANS_DATE_COL_IDX_IN_RANGE];
    const transCategory = String(row[TRANS_CATEGORY_COL_IDX_IN_RANGE]).trim();
    const transAmount = parseFloat(row[TRANS_AMOUNT_COL_IDX_IN_RANGE]);

    // Validate data types and existence
    if (!(transactionDate instanceof Date) || isNaN(transactionDate.getTime()) || isNaN(transAmount) || !transCategory) {
      continue;
    }

    const lookupKey = transCategory.toLowerCase(); // Match normalized category
    const categoryInfo = categoryLookup[lookupKey];

    if (categoryInfo) { // If this category is defined in our Groupings
      const { principal, type } = categoryInfo;

      // Extract Year and Month for grouping
      const year = transactionDate.getFullYear();
      const month = transactionDate.getMonth() + 1; // getMonth() is 0-indexed

      // Ensure 'YYYY-MM' format for consistent keying
      const yearMonthKey = `${year}-${month.toString().padStart(2, '0')}`;

      // Initialize nested objects if they don't exist
      if (!aggregatedData[type]) { // Group by Type first
        aggregatedData[type] = {};
      }
      if (!aggregatedData[type][principal]) { // Then by Principal
        aggregatedData[type][principal] = {};
      }
      if (!aggregatedData[type][principal][transCategory]) { // Then by Category
        aggregatedData[type][principal][transCategory] = {};
      }
      if (!aggregatedData[type][principal][transCategory][yearMonthKey]) { // Then by Year-Month
        aggregatedData[type][principal][transCategory][yearMonthKey] = 0;
      }

      // Add amount to the correct bucket
      aggregatedData[type][principal][transCategory][yearMonthKey] += transAmount;
    }
  }

  // --- 3. Format and Sort Report Data into a Tabular Matrix ---
  const reportOutput = [];
  reportOutput.push(['Type', 'Principal', 'Category', 'Year', 'Month', 'Total Amount']); // Report Header

  const dataRows = []; // Array to hold individual rows before sorting

  for (const type in aggregatedData) {
    for (const principal in aggregatedData[type]) {
      for (const category in aggregatedData[type][principal]) {
        for (const yearMonthKey in aggregatedData[type][principal][category]) {
          const totalAmount = aggregatedData[type][principal][category][yearMonthKey];
          const [year, month] = yearMonthKey.split('-'); // Split YYYY-MM key

          dataRows.push({
            type: type,
            principal: principal,
            category: category,
            year: parseInt(year), // Store as numbers for proper sorting
            month: parseInt(month), // Store as numbers for proper sorting
            amount: totalAmount
          });
        }
      }
    }
  }

  // Sort dataRows: descending by Year, then descending by Month
  dataRows.sort((a, b) => {
    if (b.year !== a.year) {
      return b.year - a.year; // Descending year
    }
    return b.month - a.month; // Descending month
  });

  // Populate reportOutput array from sorted dataRows
  if (dataRows.length === 0) {
    reportOutput.push(['No matching transactions found for the defined groupings in this period.', '', '', '', '', '']);
    // Ensure the "No data" row also has 6 columns to match the header
  } else {
    dataRows.forEach(row => {
      reportOutput.push([
        row.type.charAt(0).toUpperCase() + row.type.slice(1), // Capitalize Type
        row.principal,
        row.category,
        row.year,
        row.month.toString().padStart(2, '0'), // Ensure month is two digits for display
        row.amount.toFixed(2)
      ]);
    });
  }

  // --- 4. Output the Report to a New Sheet ---
  outputReportToNewSheet(reportOutput, spreadsheet, 'Tabular Financial Report');
}

/**
 * Helper function to output a 2D array to a new sheet.
 * @param {Array<Array<any>>} data The 2D array of data to write.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} baseSheetName The base name for the new report sheet.
 */
function outputReportToNewSheet(data, ss, baseSheetName) {
  const sheetName = `${baseSheetName}`; // Use a fixed name for consistency
  let reportSheet = ss.getSheetByName(sheetName);

  if (reportSheet) {
    reportSheet.clearContents(); // Clear previous report if sheet exists
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  // Ensure data has at least one row, and the first row has the expected number of columns.
  if (data.length === 0 || (data.length > 0 && data[0].length === 0)) {
    reportSheet.getRange(1, 1).setValue('No data to report or an internal error occurred.');
    Logger.log("outputReportToNewSheet received empty or malformed data.");
    return;
  }

  // Write the data to the sheet
  reportSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  reportSheet.autoResizeColumns(1, data[0].length);

  ss.setActiveSheet(reportSheet); // Make the new report sheet active
  SpreadsheetApp.getUi().alert('Report Generated', `Report generated in sheet: "${sheetName}"`, SpreadsheetApp.getUi().ButtonSet.OK);
}