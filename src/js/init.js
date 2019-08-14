
// main app to process values from csv file
const alias = require("./../../alias");

const fs = require("fs");
const to_json = require("txt-file-to-json");
// const fcsv = require("fast-csv");
// const ws = fs.createWriteStream(alias.location("out.json"));

let json_data = JSON.stringify(to_json(alias.location("dat.txt")));

let input_obj = 1;

// console.log(input_obj);
console.log(json_data);

fs.writeFile(alias.location("out.json"), json_data, 'utf8', function (err) {
    if (err) {
        console.log("An error occured while writing JSON Object to File.");
        return console.log(err);
    }
 
    console.log("JSON file has been saved.");
});
	