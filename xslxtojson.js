const argv = require('minimist')(process.argv.slice(2));
const sourceFile = argv._[0];
const targetFile = argv._[1];
//const XLSX = require("xlsx");
const XLSX = require('node-xlsx').default;
const fs = require('fs');
const colors = require('colors');

let count = 0;
let JSONData = {};

colors.enable();
console.clear();

try {

    if(sourceFile && targetFile) {

        const sourceData = XLSX.parse(`${__dirname}/${sourceFile}`);

        sourceData.forEach((row) => {
            row.data.forEach((rowItem) => {
                count++;
                JSONData[rowItem[5]] = rowItem[2];
            })
        })
 
        fs.promises.writeFile(targetFile, JSON.stringify(JSONData)).then((resolve) => {
            console.log(`${count} Einträge zu JSON konvertiert und in Datei ${targetFile} gespeichert`.bgGreen.white);
        })

    }else {
        console.log('Source- und Targetfile sind nicht korrekt Usage: node xlsxtojson.js <eingabefile.xslx> <ausgabefile.json>'.bgRed.white);
    }

}

catch(error) {
    console.log(`Sourcefile mit dem Namen ${sourceFile} nicht gefunden`.bgRed.white);
}