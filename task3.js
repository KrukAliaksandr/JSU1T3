let XLSX = require('xlsx');
let js  = {"name":"Alex"}
let new_ws_name = "SheetJS";
let wb = XLSX.utils.book_new();
let inputDir = "D:/work/Javascript/JSU1T3/input/";
let outputDir = "D:/work/Javascript/JSU1T3/output/";
let fs = require('fs');
let path = require('path');
let file,fileName;
let walk = function(dir, done) {
  fs.readdir(dir, function(err, list) {
    if (err) return done(err);
    let pending = list.length;
    if (!pending) return done(null);
    list.forEach(function(file) {
      fileName = file.replace(".json","");
      file = path.resolve(dir, file);
      fs.stat(file, function(err, stat) {
        if (stat && stat.isDirectory()) {
          walk(file, function(err, res) {
            if (!--pending) done(null);
          });
        } else {
          convert(file,fileName);
          if (!--pending) done(null);
        }
      });
    });
  });
};
function convert(file,fileName){
  let jsonObject = require(file);
  let wb = XLSX.utils.book_new();
  let ws_data = [[]];
  console.log(Object.entries(jsonObject));
  Object.entries(jsonObject).forEach(([key, value]) => {
    let item = [[key,value]];
    if(Array.isArray(item[1])){
      item[1] = item[1].toString;
      // let stringForArrayItems = "[";
      // forEach(valueItem in value)
      // console.log(valueItem);
      // {
      //   stringForArrayItems = stringForArrayItems.concat(",",value);
      // }
      // stringForArrayItems = stringForArrayItems.concat("","]");
      // item[1] = stringForArrayItems;
    } 
  ws_data = ws_data.concat(item);})
  let ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, new_ws_name);
  if(typeof require !== 'undefined') XLSX = require('xlsx');
  XLSX.writeFile(wb,outputDir + fileName + ".xlsx");
}
walk(inputDir, function(err) {
  if (err) throw err;
});