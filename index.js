"use strict";
const XLSX = require('xlsx');
const Sequelize = require("sequelize");
const excelbuilder = require('msexcel-builder');
const express = require('express');
const formidable = require('formidable');
const fs = require('fs');
const bodyParser = require('body-parser');

const app = express();
const sequelize = new Sequelize('postgres://postgres:11111@localhost:5432/mvd', {logging: false});

app.use(express.static("public"));
app.use(bodyParser.urlencoded({ extended: true }, {defer: true}));


app.post('/upload', function(req, res){
   const form = new formidable.IncomingForm();
    form.keepExtensions = true;
    form.parse(req, function(err, fields, files){
        const workbook = XLSX.readFile(files.fileUploaded.path);
        const list = workbook.Sheets["Лист1"];

        const promises = [], resultData = [];
        parseData(list, promises, resultData);

        Promise.all(promises).then(result => {
            let error_num = 0, success_num = 0;
            res.writeHead(200, {"Content-Type": "text/html; charset=utf-8"});
            res.write('<table border="1px" width="50%"><tr><td align="center">id</td><td align="center">fullName</td><td align="center">Error text</td></tr>');
            for(let i =0; i< resultData.length; i++){
                if(resultData[i].type == 'SUCCESS')
                    success_num++;
                else{
                    error_num++;
                    res.write('<tr><td>' + resultData[i].id + '</td><td>' + resultData[i].fullName + '</td><td>' + resultData[i].text + '</td></tr>');
                }
            }
            res.write('</table></br></br></br>');

            res.write('Total number of rows: ' + resultData.length + '</br>');
            res.write('Number of successful rows: ' + success_num + '</br>');
            res.write('Number of error rows: ' + error_num);
            res.end();
        })
    });
});

function parseData(sourse, promiseArr, resultArr) {
    for (let row = 3;; row++){
        if(!sourse['A' + row])
            break;
        else{
            promiseArr.push(sequelize.query('INSERT into policemans (full_name, post, email, work_phone, mobile_phone, location, work_territory, image_url, created_at, updated_at) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)',
                { bind: [sourse['B' + row].h + ' ' + sourse['C' + row].h + ' ' + sourse['D' + row].h,
                    sourse['G' + row].h,
                    sourse['I' + row].h,
                    sourse['J' + row].h,
                    sourse['K' + row].h,
                    sourse['H' + row].h,
                    sourse['L' + row].h,
                    sourse['M' + row].h,
                    new Date(),
                    new Date()]})
                .then(result => {
                    resultArr.push({type: 'SUCCESS', id: sourse['A' + row].h, fullName: sourse['B' + row].h + ' ' + sourse['C' + row].h + ' ' + sourse['D' + row].h, text: "Line " + row + " has been added successfully"});
                })
                .catch(err => {
                    resultArr.push({type: 'ERROR', id: sourse['A' + row].h, fullName: sourse['B' + row].h + ' ' + sourse['C' + row].h + ' ' + sourse['D' + row].h, text: "Line " + row + ": " + err.message});
                })
            );
        }
    }
}

app.listen(3000, function () {
    console.log('listening on port 3000');
});