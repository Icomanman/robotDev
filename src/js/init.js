
// main app to process values from csv file

const fs = require("fs");
const csv = require("csv-parser");
const fcsv = require("fast-csv");
const ws = fs.createWriteStream("out.csv");

// fastcsv
//   .write(data, { headers: true })
//   .pipe(ws);
	