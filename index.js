"use strict";
const XLSX = require('xlsx');
const Sequelize = require("sequelize");
const excelbuilder = require('msexcel-builder');

const workbook = XLSX.readFile('test.xlsx');
const list = workbook.Sheets["Лист1"];
const sequelize = new Sequelize('postgres://postgres:11111@localhost:5432/mvd', {logging: false});
const xlsxfile = excelbuilder.createWorkbook('./', 'error.xlsx');
const sheet = xlsxfile.createSheet('Лист1', 13, 100);



for(let column = 1; column <= 13; column++){
    sheet.set(column, 1, list[String.fromCharCode(64 + column) + 2].h);
}

var error_row = 2;
for (let row = 3; row <= 4; row++) {
    sequelize.query('INSERT into policemans (full_name, post, email, work_phone, mobile_phone, location, work_territory, image_url, created_at, updated_at) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)',
        { bind: [list['B' + row].h + ' ' + list['C' + row].h + ' ' + list['D' + row].h,
            list['G' + row].h,
            list['I' + row].h,
            'jgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjgjg',
            //list['J' + row].h,
            list['K' + row].h,
            list['H' + row].h,
            list['L' + row].h,
            list['M' + row].h,
            new Date(),
            new Date()]})
        .then(result => {
            console.log("The line has been added successfully");
        })
        .catch(err => {
            for(let error_col = 1; error_col <= 13; error_col++){
                //sheet.set(error_col, error_row, 'test');
                //console.log(list[String.fromCharCode(64 + error_col) + row].h);
                sheet.set(error_col, error_row, list[String.fromCharCode(64 + error_col) + row].h);
            }
            error_row++;
            console.error(err.message);
        })
}


xlsxfile.save(function(result){
    if (result == null)
        console.log('error book created');
    else{
        console.error('error book was not created');
        xlsxfile.cancel();
    }
});