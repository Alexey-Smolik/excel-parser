"use strict";
const XLSX = require('xlsx');
const pg = require('pg');

const workbook = XLSX.readFile('test.xlsx');
const list = workbook.Sheets["Лист1"];

const policemen = [];

for (let row = 3; row <= 3; row++) {
    var obj = {
        id: parseInt(list['A' + row].h.substring(0, list['A' + row].h.length - 1)),
        full_name: list['B' + row].h + ' ' + list['C' + row].h + '. ' + list['D' + row].h + '.',
        post: list['G' + row].h,
        email: list['I' + row].h,
        work_phone: list['J' + row].h,
        mobile_phone: list['K' + row].h,
        location:list['H' + row].h,
        work_territory:list['L' + row].h,
        image_url: list['M' + row].h
    };
    policemen.push(obj);
};


var conString = "postgres://postgres:11111@localhost:5432/mvd";

var client = new pg.Client(conString);
client.connect();

//queries are queued and executed one after another once the connection becomes available
client.query('INSERT into policemans (id, full_name, post, email, work_phone, mobile_phone, location, work_territory, image_url, created_at, updated_at) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)RETURNING id',
    [policemen[0].id, policemen[0].full_name, policemen[0].post,
        policemen[0].email, policemen[0].work_phone, policemen[0].mobile_phone,
        policemen[0].location, policemen[0].work_territory, policemen[0].image_url,
        new Date(), new Date()]);

var query = client.query("SELECT * FROM policemans");
//fired after last row is emitted

query.on('row', function(row) {
    console.log("success");
    console.log(row);
});

query.on('end', function() {
    console.log("ending");
    client.end();
});



/*var pool = new pg.Pool(conString);

// connection using created pool
pool.connect(function(err, client, done) {
    client.query('INSERT into policemans (id, full_name, post, email, work_phone, mobile_phone, location, work_territory, image_url, created_at, updated_at) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)',
        [policemen[0].id, policemen[0].full_name, policemen[0].post,
            policemen[0].email, policemen[0].work_phone, policemen[0].mobile_phone,
            policemen[0].location, policemen[0].work_territory, policemen[0].image_url,
            new Date(), new Date()]);
    done();
});

pool.end()*/