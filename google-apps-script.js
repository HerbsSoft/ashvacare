/**
 * SugarDefend Pro™ - Google Apps Script for Lead Capture
 * 
 * SETUP:
 * 1. Create Google Sheet > Name it "SugarDefend Pro Leads"
 * 2. Extensions > Apps Script > Paste this code
 * 3. Run setupSheet() once to create headers
 * 4. Deploy > New Deployment > Web App > Anyone access
 * 5. Copy URL and paste in index.html SHEET_URL variable
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Leads') || ss.getActiveSheet();
    
    // Get plan name
    const planNames = {
      'starter': 'Starter (₹799 - 15 दिन)',
      'monthly': 'Monthly (₹1,499 - 30 दिन)',
      'combo': 'Combo (₹3,999 - 90 दिन)'
    };
    
    sheet.appendRow([
      new Date(),                           // Timestamp
      data.name || 'N/A',                   // Name
      data.phone || 'N/A',                  // Phone
      planNames[data.plan] || data.plan,    // Plan
      data.source || 'landing-page',        // Source
      data.page || '',                      // Page URL
      'New',                                // Status
      '',                                   // Notes
      ''                                    // Follow-up
    ]);
    
    // Optional: Send notification
    // sendNotification(data.name, data.phone, data.plan);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'SugarDefend Pro API Running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');
  
  const headers = ['Timestamp', 'Name', 'Phone', 'Plan', 'Source', 'Page URL', 'Status', 'Notes', 'Follow-up Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FF6B00')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Column widths
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 200);
  sheet.setColumnWidth(9, 120);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Called 1', 'Called 2', 'WhatsApp Sent', 'Order Confirmed', 'Delivered', 'Not Interested'])
    .build();
  sheet.getRange('G2:G1000').setDataValidation(statusRule);
  
  Logger.log('✅ Sheet setup complete!');
  SpreadsheetApp.getUi().alert('Sheet setup complete! Headers created successfully.');
}

/**
 * Optional: Get daily stats
 */
function getDailyStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leads') || ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date().toDateString();
  
  let total = data.length - 1;
  let todayCount = 0;
  let planCounts = { starter: 0, monthly: 0, combo: 0 };
  
  for (let i = 1; i < data.length; i++) {
    if (new Date(data[i][0]).toDateString() === today) todayCount++;
    const plan = (data[i][3] || '').toLowerCase();
    if (plan.includes('starter')) planCounts.starter++;
    else if (plan.includes('monthly')) planCounts.monthly++;
    else if (plan.includes('combo')) planCounts.combo++;
  }
  
  Logger.log(`Total Leads: ${total}`);
  Logger.log(`Today's Leads: ${todayCount}`);
  Logger.log(`Starter: ${planCounts.starter}, Monthly: ${planCounts.monthly}, Combo: ${planCounts.combo}`);
  
  return { total, today: todayCount, plans: planCounts };
}
