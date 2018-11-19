/* eslint-disable no-console */
/* eslint-disable indent */
const XLSX = require("xlsx");
const fs = require("fs");

const new_ws_name = "Sheet1";
const inputDir = process.argv[2];
const outputDir = process.argv[3];
let  results = [];

function readDirectory(directory) {
  return new Promise((resolve, reject) => {
    fs.readdir(directory, (err, files) => {
      files.forEach(file => {
        results.push(file);
      });
      if (err) reject(err);
      else {
        console.log(results);
        resolve(results);
      }
    });
  });
}

function readFilesAndConvertToXlsx(inputDir) {
  readDirectory(inputDir).then(function (results) {
    convertResults(results);
  }).catch();
}

function convertResults(results) {
  results.forEach((file) => {
    let filePath = inputDir + file;
    let jsonObject = require(filePath);
    let wb = XLSX.utils.book_new();
    let ws_data = [];
    console.log(Object.entries(jsonObject));
    Object.entries(jsonObject).forEach(([key, value]) => {
      let item = [[key, value]];
      ws_data = ws_data.concat(item);
    });
    let ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, new_ws_name);
    XLSX.writeFile(wb, outputDir + file.replace(".json", ".xlsx"));
  });
}

readFilesAndConvertToXlsx(inputDir);