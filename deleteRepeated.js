/*
code that deletes the repeats in each row, keeps the non-repeats and creates a new excel
*/
const xlsx = require("xlsx");

// Read excel
const workbook = xlsx.readFile("test.xlsx");

// Choose first sheet
const worksheet = workbook.Sheets[workbook.SheetNames[1]];

// Read rows and transfer data to an array
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Create an array to hold non-repeating lines
const uniqueRows = [];
const newRow = [];

const totalArray = [];

// Check data each raw
for (let i = 0; i < rows.length; i++) {
  const row = rows[i];
  const myArray = [];

  // Create a new array for delete repeated data
  const newRow = [];

  // Check every data and when they not repeat add to array
  for (let j = 0; j < row.length; j++) {
    let cellValue = row[j];
    cellValue = cellValue + "";
    if (cellValue && cellValue.includes(";")) {
      const arrayValue = cellValue.split(";");
      const uniq = [...new Set(arrayValue)];
      myArray.push(uniq.join(";"));
      if (myArray.length == 2) totalArray.push(myArray);
    }
  }
}

// Create a workbook and give to name
const newWorkbook = xlsx.utils.book_new();
const newWorksheet = xlsx.utils.aoa_to_sheet(totalArray);
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Unique Rows");

// Save the workbook
xlsx.writeFile(newWorkbook, "unique_rows.xlsx");
