const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');

async function fetchDataAndCreateExcel() {
  const apiUrl = 'https://reqres.in/api/users?';

  try {
    // Make HTTP request
    const response = await axios.get(apiUrl);

    // Parse JSON response
    const responseData = response.data;

    // Create Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Users');

    // Add headers to the worksheet
    const headers = Object.keys(responseData.data[0]);
    worksheet.addRow(headers);

    // Add data rows to the worksheet
    responseData.data.forEach((user) => {
      const row = headers.map((header) => user[header]);
      worksheet.addRow(row);
    });

    // Save the workbook to a file
    await workbook.xlsx.writeFile('users.xlsx');

    console.log('Excel file created successfully.');
  } catch (error) {
    console.error('Error fetching data:', error.message);
  }
}



async function readExcelFile(filePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const data = [];
      workbook.eachSheet(function (worksheet, sheetId) {
      const worksheetGet = workbook.getWorksheet(sheetId);
      
      const columnHeaders = worksheet.getRow(1).values;

      worksheetGet.eachRow((row, rowNumber) => {
        if (rowNumber !== 1) {
          const rowData = {};
          row.eachCell((cell, colNumber) => {
            rowData[columnHeaders[colNumber - 1]] = cell.value;
          });
          data.push(rowData);
        }
          })
    });
    return data;

  } catch (error) {
    console.error('Error reading Excel file:', error.message);
    throw error;
  }
}


fetchDataAndCreateExcel();

const filePath = path.join(__dirname, 'users.xlsx');

readExcelFile(filePath)
  .then(data => {
    console.log('Excel Data Extracted:', data);
  })
  .catch(error => {
    // Handle error
  });




