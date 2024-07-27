var ranges = {};
var totalDemandForRestockingPeriod = 30;
var leadTime = 7;
var stockMovementSpreadsheetId = '1f2BxO8DvGvpAVQ82oa9nX1Wj6SlGxyoPmB5jWPLxuqU'; // Add the StockMovement spreadsheet ID here
var project = 'atomic-venture-430309-a2';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Extra Function')
    .addItem('Low Stock Detection', 'showSideBar')
    .addItem('Restock', 'showEmailBar')
    .addToUi();
}

function showSideBar() {
  var html = HtmlService.createHtmlOutputFromFile('SideBar')
      .setTitle('Enter Ranges')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetNames() {
  var stockMovementSpreadsheet = SpreadsheetApp.openById(stockMovementSpreadsheetId);
  var sheets = stockMovementSpreadsheet.getSheets();
  return sheets.map(function(sheet) {
    return sheet.getName();
  });
}

function getStockBalanceSheetNames() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(function(sheet) {
    return sheet.getName();
  });
}

function processRanges(stockProductCodeStartRange, stockProductCodeEndRange, stockProductDescStartRange, stockProductDescEndRange, stockBalanceStartRange, stockBalanceEndRange, productCodeSheetName, productCodeStartRange, productCodeEndRange, salesStartRange, salesEndRange) {
  var ui = SpreadsheetApp.getUi();
  var stockMovementSpreadsheet = SpreadsheetApp.openById(stockMovementSpreadsheetId);
  var stockBalanceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Store ranges globally
    ranges.stockProductCodeStartRange = stockProductCodeStartRange;
    ranges.stockProductCodeEndRange = stockProductCodeEndRange;
    ranges.stockProductDescStartRange = stockProductDescStartRange;
    ranges.stockProductDescEndRange = stockProductDescEndRange;
    ranges.stockBalanceStartRange = stockBalanceStartRange;
    ranges.stockBalanceEndRange = stockBalanceEndRange;
    ranges.productCodeSheetName = productCodeSheetName;
    ranges.productCodeStartRange = productCodeStartRange;
    ranges.productCodeEndRange = productCodeEndRange;
    ranges.salesStartRange = salesStartRange;
    ranges.salesEndRange = salesEndRange;
    
    // Log available sheet names for debugging
    var availableSheets = stockMovementSpreadsheet.getSheets().map(sheet => sheet.getName());
    Logger.log('Available sheets in StockMovement spreadsheet: ' + availableSheets.join(', '));

    // Get the product code sheet and ranges from the StockMovement spreadsheet
    var productCodeSheet = stockMovementSpreadsheet.getSheetByName(productCodeSheetName);
    if (!productCodeSheet) {
      throw new Error('Invalid product code sheet name: ' + productCodeSheetName);
    }
    
    var productCodeRange = productCodeSheet.getRange(productCodeStartRange + ':' + productCodeEndRange);
    var productCodes = productCodeRange.getValues();
    var productSalesData = {};

    // Initialize productSalesData with product codes
    for (var i = 0; i < productCodes.length; i++) {
      var productCode = productCodes[i][0];
      productSalesData[productCode] = [];
    }

    // Process the current sheet
    aggregateSales(stockMovementSpreadsheet, productCodes, productSalesData);

    // Process subsequent sheets in the StockMovement spreadsheet
    var sheets = stockMovementSpreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      if (sheet.getName() === productCodeSheetName) continue;
      aggregateSales(sheet, productCodes, productSalesData);
    }

    // Calculate restocking levels for each product
    var ROPs = calculateROP(productSalesData);

    // Prepare the result
    var combinedValues = [];
    for (var productCode in ROPs) {
      combinedValues.push([productCode, ROPs[productCode]]);
    }

    // Compare ROP with StockBalance
    compareWithStockBalance(ROPs, stockBalanceSpreadsheet);

  } catch (error) {
    Logger.log(error.message);
    ui.alert('Error: ' + error.message);
  }
}

function aggregateSales(sheet, productCodes, productSalesData) {
  var salesRange = sheet.getRange(ranges.salesStartRange + ':' + ranges.salesEndRange);
  var sales = salesRange.getValues();
  
  // Ensure the number of rows match
  if (productCodes.length !== sales.length) {
    throw new Error('Mismatch in row count between product codes and sales in sheet: ' + sheet.getName());
  }
  
  // Combine product codes with sales
  for (var j = 0; j < productCodes.length; j++) {
    var productCode = productCodes[j][0];
    var sale = Math.abs(sales[j][0]); // Ensure the sales value is positive
    if (productSalesData[productCode] !== undefined) {
      productSalesData[productCode].push(sale);
    }
  }
}

function calculateROP(productSalesData) {
  var ROPs = {};

  // Calculate restocking level for each product
  for (var productCode in productSalesData) {
    var sales = productSalesData[productCode];
    
    if (sales.length === 0) {
      continue; // Skip products with no sales data
    }

    // Calculate average daily sales
    var totalSales = sales.reduce((sum, sale) => sum + sale, 0);
    
    if (totalSales != 0) {
      var averageDailySales = totalSales / (sales.length * 30);

      // Calculate lead time demand
      var leadTimeDemand = averageDailySales * leadTime;
      
      // Calculate safety stock
      var ROP = leadTimeDemand + totalDemandForRestockingPeriod;
      
      // Calculate restocking level and round down
      ROP = Math.round(ROP);
      
    } else {
      ROP = 0;
    }

    ROPs[productCode] = ROP;
  }

  return ROPs;
}

function compareWithStockBalance(ROPs, stockBalanceSpreadsheet) {
  var ui = SpreadsheetApp.getUi();
  var stockBalanceSheet = stockBalanceSpreadsheet.getActiveSheet();

  if (!stockBalanceSheet) {
    throw new Error('Invalid stock balance sheet name: ' + ranges.stockBalanceSheetName);
  }

  // Get the product codes and balances
  var stockProductCodes = stockBalanceSheet.getRange(ranges.stockProductCodeStartRange + ':' + ranges.stockProductCodeEndRange).getValues().flat();
  var stockProductDescs = stockBalanceSheet.getRange(ranges.stockProductDescStartRange + ':' + ranges.stockProductDescEndRange).getValues().flat();
  var stockBalances = stockBalanceSheet.getRange(ranges.stockBalanceStartRange + ':' + ranges.stockBalanceEndRange).getValues().flat();

  var lowStockProducts = [];

  // Compare ROPs with stock balances
  for (var i = 0; i < stockProductCodes.length; i++) {
    var productCode = stockProductCodes[i];
    var productDesc = stockProductDescs[i];
    var stockBalance = stockBalances[i];

    if (ROPs[productCode] !== undefined && stockBalance < ROPs[productCode]) {
      lowStockProducts.push([productCode, productDesc, stockBalance, ROPs[productCode]]);
    }
  }

  // If there are products with stock balances lower than their ROP, create a new sheet within the spreadsheet
  if (lowStockProducts.length > 0) {
    var newSheet = stockBalanceSpreadsheet.getSheetByName('Low Stock Products') || stockBalanceSpreadsheet.insertSheet('Low Stock Products');
    newSheet.clear(); // Clear the sheet if it already exists

    // Add headers
    newSheet.appendRow(['Product Code', 'Product Description', 'Stock Balance', 'ROP']);

    // Add data
    newSheet.getRange(2, 1, lowStockProducts.length, 4).setValues(lowStockProducts);

    ui.alert('A new sheet "Low Stock Products" has been created within the spreadsheet.');
  } else {
    ui.alert('All products have sufficient stock.');
  }
}

function auth() {
  var cache = CacheService.getUserCache();
  var token = ScriptApp.getOAuthToken();  
  cache.put("token", token);
}

function showEmailBar() {
  auth();
  var html = HtmlService.createHtmlOutputFromFile('EmailBar')
      .setTitle('Restock Request')
      .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getProducts() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var productSheet = spreadsheet.getSheetByName("Low Stock Products");
  var lastrow = productSheet.getLastRow();
  var productDataRange = productSheet.getRange("B2:B"+lastrow); 
  var productData = productDataRange.getValues();

  return productData.map(function(row) {
    return { name: row[0], quantity: row[1] };
  }).filter(product => product.name); // Filter out empty product names
}

function sendRestockRequest(supplierEmail, selectedProducts) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var productSheet = spreadsheet.getSheetByName("Low Stock Products");

    if (!productSheet) {
      Logger.log("Sheet 'Low Stock Products' not found!");
      return "Error: Sheet 'Low Stock Products' not found!";
    }

    // Build the content prompt for Gemini
    var productList = Object.keys(selectedProducts).map(function(product) {
      var restockQty = selectedProducts[product];
      return `Product: ${product}, Quantity: ${restockQty}`;
    }).join("\n");

    var prompt = `Write an email to ${supplierEmail} requesting a quotation for the following products with their respective restock quantities:\n\n${productList}\n\nGiven my company name is Kabac Cable and my name is Bernard. Generate the email in plain text without any font styles. Do not include email subject. Use salutation of Sir or Madam.`;

    // Get the email body from Gemini
    var emailBody = askGemini(prompt);

    if (emailBody.length > 10000) {
      Logger.log("Body too long!");
      return "Email body is too long!";
    }

    if (!emailBody.startsWith("ERROR")) {
      GmailApp.sendEmail(supplierEmail, "Request for Restock Quotation", emailBody);
      Logger.log(`Email sent successfully to ${supplierEmail}!`);
      return `Email sent successfully to ${supplierEmail}!`;
    } else {
      Logger.log(`Failed to generate email: ${emailBody}`);
      return `Failed to generate email: ${emailBody}`;
    }
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    return `Error: ${e.message}`;
  }
}

function askGemini(prompt) {
  var cache = CacheService.getUserCache();
  var token = cache.get("token");
  if (!token) return "ERROR: Token not available.";
  
  var url = `https://us-central1-aiplatform.googleapis.com/v1/projects/${project}/locations/us-central1/publishers/google/models/gemini-1.0-pro:generateContent`;  

  var data = {
    contents: {
      role: "USER",
      parts: [{ "text": prompt }]
    },  
    generation_config: {
      temperature: 0.3,
      topP: 1,
      maxOutputTokens: 256        
    }
  };
  
  var options = {
    method: "post",
    contentType: 'application/json',   
    headers: {
     Authorization: `Bearer ${token}`,
    },
    payload: JSON.stringify(data)
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() == 200) {
      var json = JSON.parse(response.getContentText());
      var answer = json.candidates && json.candidates[0].content && json.candidates[0].content.parts[0].text;
      Logger.log(`Generated Content Length: ${answer.length}`);
      return answer || "ERROR: No content returned.";
    } else {
      return `ERROR: Response code ${response.getResponseCode()}`;
    }
  } catch (e) {
    return `ERROR: ${e.message}`;
  }
}
