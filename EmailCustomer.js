function sendVarianceReportEmailDraft() {
  sendVarianceReportEmail(false);
}

function sendVarianceReportEmailDefinitive() {
  sendVarianceReportEmail(true);
}
/**
 * Sends an email with the summarized report from "GenerateReport" tab,
 * including detailed comments for significant variances.
 */
function sendVarianceReportEmail(sendDefinitive) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const generateReportSheet = spreadsheet.getSheetByName('GenerateReport');
  if (!generateReportSheet) {
    ui.alert('Error', 'The "GenerateReport" sheet was not found. Please ensure it exists and is named "GenerateReport".', ui.ButtonSet.OK);
    return;
  }

  const tabularReportSheet = spreadsheet.getSheetByName('Tabular Financial Report');
  if (!tabularReportSheet) {
    ui.alert('Error', 'The "Tabular Financial Report" sheet was not found. Please ensure it exists and is named "Tabular Financial Report".', ui.ButtonSet.OK);
    return;
  }

  // --- 1. Get Email Recipient, Year, and Month ---
  const recipientEmail = generateReportSheet.getRange('B5').getValue();
  const reportYear = generateReportSheet.getRange('B2').getValue();
  const reportMonthText = generateReportSheet.getRange('B3').getValue(); // e.g., "Jan"

  // Convert month text (e.g., "Jan") to a two-digit number string (e.g., "01") for matching
  let reportMonthNumberString;
  try {
    const dateValue = new Date(reportMonthText + " 1, " + reportYear);
    reportMonthNumberString = (dateValue.getMonth() + 1).toString();
  } catch (e) {
    ui.alert('Error', 'Could not parse month from B3 or year from B2. Please ensure B3 is a valid 3-letter month (e.g., "Jan") and B2 is a valid year (e.g., 2024). Error: ' + e.message, ui.ButtonSet.OK);
    return;
  }

  if (!recipientEmail || !reportYear || !reportMonthText) {
    ui.alert('Missing Info', 'Please ensure B2 (Year), B3 (Month), and B5 (Recipient Email) are filled in the "GenerateReport" tab.', ui.ButtonSet.OK);
    return;
  }

  // --- 2. Get Main Report Table Data from "GenerateReport" ---
  // Assuming table starts at A8 and goes downwards
  // Columns: A (Type), B (Principal), C (Current Amount), D (Budget), E (Variance)
  const reportDataRange = generateReportSheet.getRange('A9:E' + generateReportSheet.getLastRow());
  const reportValues = reportDataRange.getValues();

  // --- 3. Build HTML Table for Email ---
  let htmlTable = '<table style="width:100%; border-collapse: collapse;">';
  // Table Headers
  htmlTable += '<tr style="background-color:#f2f2f2;">';
  htmlTable += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Type</th>';
  htmlTable += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Principal</th>';
  htmlTable += '<th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Current Amount</th>';
  htmlTable += '<th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Budget</th>';
  htmlTable += '<th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Variance</th>';
  htmlTable += '</tr>';

// Table Rows - Apply formatting here
  const VARIANCE_THRESHOLD = 0.20; // 20%
  reportValues.forEach(row => {
    const type = String(row[0]).trim();
    const principal = String(row[1]).trim();
    const currentAmount = parseFloat(row[2]);
    const budget = parseFloat(row[3]);
    // Variance (row[4]) is the absolute difference from the sheet, we will re-calculate percentage.

    htmlTable += '<tr>';
    htmlTable += `<td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${type}</td>`;
    htmlTable += `<td style="border: 1px solid #ddd; padding: 8px; text-align: left;">${principal}</td>`;

    // Format Current Amount (USD, 2 decimals)
    htmlTable += `<td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${isNaN(currentAmount) ? '' : 'USD ' + currentAmount.toFixed(2)}</td>`;

    // Format Budget (USD, 2 decimals)
    htmlTable += `<td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${isNaN(budget) ? '' : 'USD ' + budget.toFixed(2)}</td>`;

    // Calculate and Format Variance as Percentage with Conditional Coloring
    let variancePercentageDisplay = '';
    let varianceCellColor = '';

    if (!isNaN(currentAmount) && !isNaN(budget) && budget !== 0) {
      const variancePercentage = (currentAmount - budget) / budget;
      variancePercentageDisplay = (variancePercentage * 100).toFixed(2) + '%';

      // Determine color based on variance rules
      const lowerCaseType = type.toLowerCase();
      if (lowerCaseType === 'income' || lowerCaseType === 'savings') {
        if (variancePercentage < -VARIANCE_THRESHOLD) {
          varianceCellColor = 'red'; // Significantly under budget for income/savings
        } else if (variancePercentage > VARIANCE_THRESHOLD) {
          varianceCellColor = 'green'; // Significantly over budget for income/savings
        }
      } else if (lowerCaseType === 'expense') {
        if (variancePercentage > VARIANCE_THRESHOLD) {
          varianceCellColor = 'red'; // Significantly over budget for expenses
        } else if (variancePercentage < -VARIANCE_THRESHOLD) {
          varianceCellColor = 'green'; // Significantly under budget for expenses
        }
      }
    }

    htmlTable += `<td style="border: 1px solid #ddd; padding: 8px; text-align: right; ${varianceCellColor ? 'color:' + varianceCellColor + ';' : ''}">`;
    htmlTable += variancePercentageDisplay;
    htmlTable += `</td>`;

    htmlTable += '</tr>';
  });
  htmlTable += '</table>';





  // --- 4. Get All Data from "Tabular Financial Report" for detailed lookup ---
  // This is efficient as we read the whole sheet once.
  const tabularDataRange = tabularReportSheet.getDataRange();
  const tabularValues = tabularDataRange.getValues(); // Includes headers

  // Find column indices dynamically for "Tabular Financial Report"
  const tabularHeaders = tabularValues[0]; // Assuming first row is headers
  const TAB_TYPE_COL_IDX = tabularHeaders.indexOf('Type');
  const TAB_PRINCIPAL_COL_IDX = tabularHeaders.indexOf('Principal');
  const TAB_CATEGORY_COL_IDX = tabularHeaders.indexOf('Category');
  const TAB_YEAR_COL_IDX = tabularHeaders.indexOf('Year');
  const TAB_MONTH_COL_IDX = tabularHeaders.indexOf('Month');
  const TAB_AMOUNT_COL_IDX = tabularHeaders.indexOf('Total Amount');

  // Validate all necessary columns are found
  if ([TAB_TYPE_COL_IDX, TAB_PRINCIPAL_COL_IDX, TAB_CATEGORY_COL_IDX, TAB_YEAR_COL_IDX, TAB_MONTH_COL_IDX, TAB_AMOUNT_COL_IDX].some(idx => idx === -1)) {
    ui.alert('Error', 'One or more required columns (Type, Principal, Category, Year, Month, Total Amount) not found in "Tabular Financial Report". Please check headers.', ui.ButtonSet.OK);
    return;
  }

  const tabularData = tabularValues.slice(1); // Actual data, without headers

  // --- 5. Identify Variances and Generate Comments ---
  let varianceComments = '<h2>Variance Analysis:</h2>';
  let hasSignificantVariance = false;
  //const VARIANCE_THRESHOLD = 0.20; // 20%

  reportValues.forEach(row => {
    const type = String(row[0]).trim(); // Column A: Type
    const principal = String(row[1]).trim(); // Column B: Principal
    const currentAmount = parseFloat(row[2]); // Column C: Current Amount
    const budget = parseFloat(row[3]); // Column D: Budget
    const varianceAmount = parseFloat(row[4]); // Column E: Variance (C-D)

    if (isNaN(currentAmount) || isNaN(budget) || budget === 0) {
      // Skip if amounts are invalid or budget is zero to avoid DIV/0
      return;
    }

    const variancePercentage = (currentAmount - budget) / budget;

    let needsComment = false;
    let commentType = '';

    // Condition 1: Negative variance > 20% for Income or Savings
    if ((type.toLowerCase() === 'income' || type.toLowerCase() === 'savings') && variancePercentage < -VARIANCE_THRESHOLD) {
      needsComment = true;
      commentType = 'under budget';
    }
    // Condition 2: Positive variance > 20% for Expense
    else if (type.toLowerCase() === 'expense' && variancePercentage > VARIANCE_THRESHOLD) {
      needsComment = true;
      commentType = 'over budget';
    }

    if (needsComment) {
      hasSignificantVariance = true;
      varianceComments += `<p><strong>${principal} (${type}) is ${commentType} by ${Math.abs(variancePercentage * 100).toFixed(0)}%.</strong></p>`;
      varianceComments += `<ul>`;

      // Find top contributors from "Tabular Financial Report"
      const relevantTransactions = tabularData.filter(tabRow =>
        String(tabRow[TAB_PRINCIPAL_COL_IDX]).trim() === principal &&
        String(tabRow[TAB_TYPE_COL_IDX]).trim() === type &&
        String(tabRow[TAB_YEAR_COL_IDX]).trim() === String(reportYear) &&
        String(tabRow[TAB_MONTH_COL_IDX]).trim() === reportMonthNumberString
      );

      // Sort contributors based on variance direction
      //if (commentType === 'under budget') { // Income/Savings: look for lowest (most negative) amounts
      //  relevantTransactions.sort((a, b) => parseFloat(a[TAB_AMOUNT_COL_IDX]) - parseFloat(b[TAB_AMOUNT_COL_IDX]));
      //} else { // Expense: look for highest (most positive) amounts
      //  relevantTransactions.sort((a, b) => parseFloat(b[TAB_AMOUNT_COL_IDX]) - parseFloat(a[TAB_AMOUNT_COL_IDX]));
      //}

      relevantTransactions.sort((a, b) => Math.abs(parseFloat(b[TAB_AMOUNT_COL_IDX])) - Math.abs(parseFloat(a[TAB_AMOUNT_COL_IDX])));


      // Add top 3 contributors (or fewer if not enough)
      const topContributors = relevantTransactions.slice(0, 3);
      if (topContributors.length > 0) {
        topContributors.forEach(contributor => {
          const category = String(contributor[TAB_CATEGORY_COL_IDX]).trim();
          const amount = parseFloat(contributor[TAB_AMOUNT_COL_IDX]);
          varianceComments += `<li>${category}: USD ${Math.abs(amount).toFixed(2)}</li>`;
        });
      } else {
        varianceComments += `<li>No detailed transactions found for this period.</li>`;
      }
      varianceComments += `</ul>`;
    }
  });

  if (!hasSignificantVariance) {
    varianceComments += '<p>No significant variances found for this report period.</p>';
  }

  // --- 6. Send Email Draft ---
  const emailSubject = `PROSPR Financial Report: ${reportMonthText} ${reportYear} Variance Analysis`;
  const clientName = generateReportSheet.getRange('B4').getValue();
  const senderAlias = "";
  const defaultClientPlaceholder = "Client"; // The text to replace if clientName is empty or you want a fallback
  const emailBody = `
    <p>Dear ${clientName || defaultClientPlaceholder},</p>
    <p>Here is your financial report summary for ${reportMonthText} ${reportYear}:</p>
    ${htmlTable}
    <br>
    ${varianceComments}
    <br>
    <p>Please review and let us know if you have any questions.</p>
    <p>Best regards,<br>PROSPR Team</p>
  `;

  let attachments = [];
  try {
    const pdfBlob = spreadsheet.getAs(MimeType.PDF)
                               .setName(`${emailSubject}.pdf`);
    attachments.push(pdfBlob);
  } catch (e) {
    ui.alert('Attachment Warning', 'Could not create PDF attachment. Email will be processed without attachment. Error: ' + e.toString(), ui.ButtonSet.OK);
    // Continue without attachment if it fails
  }

if (sendDefinitive) {
    // Option A: Send Definitive Email
    try {
      GmailApp.sendEmail(
        recipientEmail,
        emailSubject,
        emailBody, // Plain text body
        {
          htmlBody: emailBody, // HTML body
          name: senderAlias,
          attachments: attachments
        }
      );
      ui.alert('Email Sent', `Financial report email definitively sent to ${recipientEmail} for ${clientName || defaultClientPlaceholder}.`, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Email Send Error', 'Failed to send email directly. Please check recipient, permissions, and network. Error: ' + e.message, ui.ButtonSet.OK);
    }
  } else {
    // Option B: Create Draft Email
    try {
      // GmailApp.createDraft does not directly accept attachments in the same way as sendEmail.
      // You'd typically need to create the draft, then programmatically attach later if truly needed via Gmail API.
      // For simplicity, createDraft below omits attachments, but you could send with attachments if "sendDefinitive" is true.
      GmailApp.createDraft(
        recipientEmail,
        emailSubject,
        '', // Plain text body, often left empty when htmlBody is used
        {
          htmlBody: emailBody
          // Attachments are typically added via Gmail API or when sending definitively.
        }
      );
      ui.alert('Draft Created', 'A draft email has been created in your Gmail for review.', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Draft Error', 'Failed to create Gmail draft. Please ensure GmailApp permissions are granted and the recipient email is valid. Error: ' + e.message, ui.ButtonSet.OK);
    }
  }
  //try {
    //GmailApp.createDraft(recipientEmail, emailSubject, '', { htmlBody: emailBody });
   // ui.alert('Email Sent', 'A draft email has been created in your Gmail for review.', ui.ButtonSet.OK);
  //} catch (e) {
  //  ui.alert('Email Error', 'Failed to create Gmail draft. Please ensure GmailApp permissions are granted and the recipient email is valid. /// + e.message, ui.ButtonSet.OK);
  //}
}