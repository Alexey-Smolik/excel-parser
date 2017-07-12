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

const result_arr = [];
const promise_arr = [];

app.post('/upload', function(req, res){
   var form = new formidable.IncomingForm();
    form.keepExtensions = true;
    form.parse(req, function(err, fields, files){
        var workbook = XLSX.readFile(files.fileUploaded.path);
        var list = workbook.Sheets["Лист1"];
        parseData(list);
        Promise.all(promise_arr).then(result => {
            var error_num = 0, success_num = 0;
            for(let i =0; i<=result_arr; i++){
                res.write(result_arr[i].type + ': ' + result_arr[i].text);

                if(result_arr[i].type == 'success')
                    success_num++;
                else
                    error_num++;
            }

            res.write('Number of successful rows: ' + success_num);
            res.write('Number of error rows: ' + error_num);
            res.end();
        })
    });
});

function parseData(list) {
    for (let row = 3; row <= 4; row++){
        promise_arr.push(sequelize.query('INSERT into policemans (full_name, post, email, work_phone, mobile_phone, location, work_territory, image_url, created_at, updated_at) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)',
            { bind: [list['B' + row].h + ' ' + list['C' + row].h + ' ' + list['D' + row].h,
                list['G' + row].h,
                list['I' + row].h,
                list['J' + row].h,
                list['K' + row].h,
                list['H' + row].h,
                list['L' + row].h,
                list['M' + row].h,
                new Date(),
                new Date()]})
                .then(result => {
                    result_arr.push({type: 'SUCCESS', text: "Line " + row + " has been added successfully"});
                })
                .catch(err => {
                    result_arr.push({type: 'ERROR', text: "Line " + row + ": " + err.message});
                })
        );
    }
};

app.listen(3000, function () {
    console.log('listening on port 3000');
});