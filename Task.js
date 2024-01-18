const XLSX = require('xlsx');
const fs = require('fs');

const filePath = 'Assignment_Timecard.xlsx';

// Read the Excel file
const workbook = XLSX.readFile(filePath);

// Open the file stream outside the loop
const outputFile = fs.createWriteStream('output.txt', { flags: 'a' }); // Use append mode to not overwrite the file

// Assuming you have columns named 'ConsecutiveDaysWorked', 'TimeBetweenShifts', 'SingleShiftDuration'
const sheet = workbook.Sheets[workbook.SheetNames[0]];
for (let row = 2; row <= XLSX.utils.decode_range(sheet['!ref']).e.r; ++row) {
    const consecutiveDays = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v || 0;
    const timeBetweenShifts = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })]?.v || 0;
    const singleShiftDuration = sheet[XLSX.utils.encode_cell({ r: row, c: 2 })]?.v || 0;

    // Check criteria and print results
    if (consecutiveDays === 7) {
        outputFile.write(`Employee at row ${row} worked for 7 consecutive days.\n`);
    }

    if (timeBetweenShifts > 1 && timeBetweenShifts < 10) {
        outputFile.write(`Employee at row ${row} has less than 10 hours between shifts but greater than 1 hour.\n`);
    }

    if (singleShiftDuration > 14) {
        outputFile.write(`Employee at row ${row} worked for more than 14 hours in a single shift.\n`);
    }
}

// Close the file after writing results for all rows
outputFile.end();

console.log('Results written to output.txt');
