// Importing express module
const express = require('express');
const XLSX = require('xlsx');
const app = express();

app.use(express.json());

app.get('/', (req, res) => {
res.sendFile(__dirname + '/index.html');
});

app.post('/', (req, res) => {
const { username, password } = req.body;
const { authorization } = req.headers;

const workbook = XLSX.readFile("excel.xlsx");
const firstSheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets["UCZESTNICY"];

const cellRef = XLSX.utils.encode_cell({c: 2, r: 4});
const cell = sheet[cellRef];
if (cell) {
    // update existing cell
    cell.v = username;
} else {
    // add new cell
    XLSX.utils.sheet_add_aoa(sheet, [[username]], {origin: cellRef});
}
XLSX.writeFile(workbook, "./excel.xlsx");

res.send({
	username,
	password,
	authorization,
});
});

app.listen(3000, () => {
console.log('Our express server is up on port 3000');
});
