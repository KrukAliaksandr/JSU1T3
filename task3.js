/* eslint-disable no-console */
/* eslint-disable indent */
const XLSX = require("xlsx");
const fs = require("fs");

const whiteSheet_name = "Sheet1";
const inputDir = process.argv[2];
const outputDir = process.argv[3];
let results = [];

function readDirectory(directory) {
  return new Promise((resolve, reject) => {
    fs.readdir(directory, (err, files) => {
      files.forEach(file => {
        results.push(file);
      });
      if (err)
      {
        reject(err);
      }
      else {
        console.log(results);
        resolve(results);
      }
      });
    });
  }


  function readFilesAndConvertToXlsx(inputDir) {
    readDirectory(inputDir).then((results)=> {
      convertResults(results);
    }).catch();
  }

  function convertResults(results) {
    results.forEach((file) => {
      let filePath = inputDir + file;
      let jsonObject = require(filePath);
      let wb = XLSX.utils.book_new();
      let whiteSheet = XLSX.utils.aoa_to_sheet(writeResults(jsonObject));
      XLSX.utils.book_append_sheet(wb, whiteSheet, whiteSheet_name);
      XLSX.writeFile(wb, outputDir + file.replace(".json", ".xlsx"));
    });
  }

  function writeResults(jsonObject) {
    let item;
    let sheetData = [];
    Object.entries(jsonObject).forEach(([key, value]) => {
      if (Array.isArray(value)) {
        let stringForArrayData = "[";
        value.forEach(arrayItem => {
          stringForArrayData = stringForArrayData.concat(",", arrayItem);
        });
        stringForArrayData = stringForArrayData.concat("]");
        value = stringForArrayData;
        item = [key, value];
      }
      else {
        item = [key, value];
      }
      sheetData = sheetData.concat([item]);
    });
    return sheetData;
  }
  readFilesAndConvertToXlsx(inputDir);