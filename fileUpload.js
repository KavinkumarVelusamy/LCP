const xlsx = require('xlsx');
const path = require('path');

// Function to read data from Excel file
function readExcel(filePath, sheetName) {
  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);
  return data;
}

// Function to normalize the data for comparison
function normalizeData(data) {
  // Normalize date from serial number to readable format (dd-mmm-yyyy)
  const normalizedData = { ...data };
  if (data.invoiceDate) {
    const date = new Date((data.invoiceDate - 25569) * 86400 * 1000);
    normalizedData.invoiceDate = date.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }).replace(/ /g, '-');
  }
  // Normalize netAmount and totalAmount to include commas
  if (data.netAmount) {
    normalizedData.netAmount = parseFloat(data.netAmount).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }
  if (data.totalAmount) {
    normalizedData.totalAmount = parseFloat(data.totalAmount).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }
  // Normalize lineItem for comparison
  if (data.lineItem) {
    normalizedData.lineItem = data.lineItem.replace(/,/g, '').replace(/\s+/g, ' ').trim();
  }
  return normalizedData;
}

module.exports = {
  'Test1': function (browser) {
    // Path to your Excel file and the sheet name
    const filePath = path.resolve(__dirname, 'C:/Users/User/Desktop/jacana.xlsx');
    const sheetName = 'Sheet3';

    // Read data from Excel
    const excelDataArray = readExcel(filePath, sheetName);

    // Iterate over each row of data in the Excel sheet
    excelDataArray.forEach((excelData, index) => {
      const normalizedData = normalizeData(excelData); // Normalize data for comparison
      console.log(`Excel Data (Row ${index + 1}):`, normalizedData); // Log the data read from Excel

      // Assuming the Excel file has columns named 'username' and 'password'
      const username = normalizedData.username;
      const password = normalizedData.password;
      const imagePath = normalizedData.imagePath;

      browser.url("http://demo.lcp.neartekpod.io/")
        .maximizeWindow()
        .pause(3000)
        .click('body > header > nav > ul:nth-child(3) > li > a > button')
        .pause(3000)
        .setValue('#username', username)
        .setValue('#password', password)
        .pause(1000)
        .click('body > div > main > section > div > div > div > form > div.cc199ae96 > button')
        .pause(2000)
        .click('body > header > nav > ul:nth-child(2) > li:nth-child(2) > a')
        .pause(5000);
      const image = path.resolve(__dirname, imagePath); // Update this with the correct file path
      browser.uploadFile('#upload', image)
        .pause(5000)

      // Get values from the webpage
      const webpageValues = {};
      browser.getValue('xpath', "//input[@id='invoice_date']", function (result) {
        webpageValues.invoiceDate = result.value;
        console.log("invoiceDate:", webpageValues.invoiceDate)
      }).getValue('xpath', "//input[@id='supplier_name']", function (result) {
        webpageValues.supplierName = result.value;
        console.log("supplierName:", webpageValues.supplierName)
      }).getValue('xpath', "//input[@id='total_tax_amount']", function (result) {
        webpageValues.totalTaxAmount = result.value;
        console.log("totalTaxAmount:", webpageValues.totalTaxAmount)
      }).getValue('xpath', "//input[@id='net_amount']", function (result) {
        webpageValues.netAmount = result.value;
        console.log("netAmount:", webpageValues.netAmount)
      }).getValue('xpath', "//input[@id='total_amount']", function (result) {
        webpageValues.totalAmount = result.value;
        console.log("totalAmount:", webpageValues.totalAmount)
      }).getValue('xpath', "//input[@id='receiver_name']", function (result) {
        webpageValues.receiverName = result.value;
        console.log("receiverName:", webpageValues.receiverName)
      }).getValue('xpath', "//input[@id='invoice_type']", function (result) {
        webpageValues.invoiceType = result.value;
        console.log("invoiceType:", webpageValues.invoiceType)
      }).getValue('xpath', "//input[@id='remit_to_name']", function (result) {
        webpageValues.remitToName = result.value;
        console.log("remitToName:", webpageValues.remitToName)
      }).getValue('xpath', "//input[@id='supplier_website']", function (result) {
        webpageValues.supplierWebsite = result.value;
        console.log("supplierWebsite:", webpageValues.supplierWebsite)
      }).getValue('xpath', "//input[@id='supplier_address']", function (result) {
        webpageValues.supplierAddress = result.value;
        console.log("supplierAddress:", webpageValues.supplierAddress)
      }).getValue('xpath', "//input[@id='line_item']", function (result) {
        webpageValues.lineItem = result.value.replace(/,/g, '').replace(/\s+/g, ' ').trim();
        console.log("lineItem:", webpageValues.lineItem)
      }).perform(() => {
        // Compare values
        const keysToCompare = ['invoiceDate', 'supplierName', 'totalTaxAmount', 'netAmount', 'totalAmount', 'receiverName', 'invoiceType', 'remitToName', 'supplierWebsite', 'supplierAddress', 'lineItem'];
        let allMatch = true;
        console.log("Results")
        keysToCompare.forEach(key => {
          if (normalizedData[key] && webpageValues[key]) {
            if (normalizedData[key] !== webpageValues[key]) {
              console.log(`Mismatch on ${key}: Expected ${normalizedData[key]}, but got ${webpageValues[key]}`);
              allMatch = false;
            } else {
             
              console.log(`Match on ${key}: ${webpageValues[key]}`);
            }
          }
        });
        console.log("Final Result :")
        if (allMatch) {
          console.log("Success: All values match!");
        } else {
          console.log("There were mismatches in the data.");
        }
      }).click('xpath', "//button[text()='Submit']")
        .pause(3000)
        .click('xpath', "//button[text()='Logout']")
        .pause(3000);
    });

    browser.end();
  }
}
