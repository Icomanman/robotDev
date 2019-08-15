
// main app to process values from csv file
const alias = require("./../../alias");

const fs = require("fs");
const csv = require("csv-parser");

// const ws = fs.createWriteStream(alias.location("out.json"));
let data = alias.location("dat.txt");
let json_data = [];

// fs.writeFile(alias.location("out.json"), json_data, 'utf8', function (err) {
//     if (err) {
//         console.log("An error occured while writing JSON Object to File.");
//         return console.log(err);
//     }

//     console.log("JSON file has been saved.");
// });

let rs = fs.createReadStream(data);

let pipo = rs.pipe(csv())

pipo.on("data", dat => json_data.push(dat));
pipo.on("end");
console.log(json_data);