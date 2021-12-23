'use strict';

const _ = require('lodash');
const fs = require('fs');
const xlsx = require('xlsx');

const workbook = xlsx.readFile('./excel/sample.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

const columnsDef = [
    {
        column: "A",
        content: "",
        fileName: "1bpo.json"
    },
    {
        column: "B",
        content: "",
        fileName: "2blinkage.json"
    },
    {
        column: "C",
        content: "",
        fileName: "3faric.json"
    },
    {
        column: "D",
        content: "",
        fileName: "4flinkage.json"
    },
    {
        column: "E",
        content: "",
        fileName: "5yarn.json"
    },
    {
        column: "F",
        content: "",
        fileName: "6ylinkage.json"
    },
    {
        column: "G",
        content: "",
        fileName: "7cotton.json"
    },
];
for (let columnDef of columnsDef) {
    for (let z in worksheet) {
        if (z.toString()[0] === columnDef.column) {
            columnDef.content += worksheet[z].v;
        }
    }
    const jsonContent = JSON.stringify(JSON.parse(columnDef.content), null, 2);
    createJSON(jsonContent, columnDef.fileName);
}


function removeFileIfExist(filePath) {
    if (fs.existsSync(filePath)) {
        try {
            fs.unlinkSync(filePath)
        } catch (err) {
            console.error(err)
        }
    }
}

function createJSON(content, fileName) {
    const filePath = './output/' + fileName;
    try {
        removeFileIfExist(filePath);
        fs.writeFileSync(filePath, content, 'utf8')
    } catch (e) {
        console.error(e);
    }
}


