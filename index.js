/**
 * This module convert file with ext .xlsx to file with etx .json, convert Excel data to json format
 * to run the transformation created in the package.json command
 * the launch which should be made as follows: npm run xlsx2json
 */
'use strict';
const fs = require('fs');
const path = require('path');
const xlsxj = require("xlsx-to-json");

class Converter {
  static jsonKeysToLowerCase(jsonData) {
    return jsonData.replace(/"([^"]+)":/g, ($0, $1) => ('"' + $1.toLowerCase() + '":'));
  }

  static xlsx2json(filePath) {
    const input = filePath;
    const ext = path.extname(input);
    const fileName = path.basename(input, ext);
    const arrPath = input.split(path.sep);
    arrPath[arrPath.length - 1] = `${fileName}.json`;
    const output = arrPath.join(path.sep);
    xlsxj({ input, output }, (err, result) => {
      if (err) console.error(err);
      else {
        const jsonData = JSON.stringify(result);
        const jsonDataLower = Converter.jsonKeysToLowerCase(jsonData);

        fs.writeFile(output, jsonDataLower, (err) => {
          if (err) console.log(err);
          else console.log(`The file ${input} was converted to file ${output}`);
        });
      }
    });
  }
}

Converter.xlsx2json(path.join(__dirname, 'mcc.xlsx'));