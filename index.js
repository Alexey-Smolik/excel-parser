"use strict";
const XLSX = require('xlsx');
const Sequelize = require("sequelize");
const express = require('express');
const formidable = require('formidable');
const fs = require('fs');
const bodyParser = require('body-parser');

const app = express();
const sequelize = new Sequelize('postgres://postgres:11111@localhost:5432/mvd', {logging: false});


const divisions_model = sequelize.define('divisions',
    {
        id: {
            type: Sequelize.INTEGER,
            autoIncrement: true,
            primaryKey: true
        },
        created_at: Sequelize.DATE,
        updated_at: Sequelize.DATE,
    }, {
        timestamps: false
    }
);


const divisions_translations = sequelize.define('divisions_translations',
    {
        id: {
            type: Sequelize.INTEGER,
            autoIncrement: true,
            primaryKey: true
        },
        name: Sequelize.STRING,
        created_at: Sequelize.DATE,
        updated_at: Sequelize.DATE,
        divisions_id: Sequelize.INTEGER,
        languages_id: {
            type: Sequelize.INTEGER,
            allowNull: true
        }
    }, {
        timestamps: false
    }
);

const subdivisions_model = sequelize.define('subdivisions',
    {
        id: {
            type: Sequelize.INTEGER,
            autoIncrement: true,
            primaryKey: true
        },
        created_at: Sequelize.DATE,
        updated_at: Sequelize.DATE,
    }, {
        timestamps: false
    }
);


const subdivisions_translations = sequelize.define('subdivisions_translations',
    {
        id: {
            type: Sequelize.INTEGER,
            autoIncrement: true,
            primaryKey: true
        },
        name: Sequelize.STRING,
        created_at: Sequelize.DATE,
        updated_at: Sequelize.DATE,
        subdivisions_id: Sequelize.INTEGER,
        languages_id: {
            type: Sequelize.INTEGER,
            allowNull: true
        }
    }, {
        timestamps: false
    }
);



var a = 1;





app.use(express.static("public"));
app.use(bodyParser.urlencoded({extended: true}, {defer: true}));


app.post('/upload', function (req, res) {
    const form = new formidable.IncomingForm();
    form.keepExtensions = true;
    form.parse(req, function (err, fields, files) {
        const workbook = XLSX.readFile(files.fileUploaded.path);
        const list = workbook.Sheets["Лист1"];
        const resultData = [], promises = [];

        parseData(list, promises, resultData)
            .then(() => Promise.all(promises))
            .then(result => {
                let error_num = 0, success_num = 0;
                res.writeHead(200, {"Content-Type": "text/html; charset=utf-8"});

                res.write('<table border="1px" width="30%" bordercolor="#9e9e9e" style="float: left; border-collapse: collapse"><tr><td align="center" colspan="3">Строки с ошибками</td></tr><tr><td align="center">id</td><td align="center">полное имя</td><td align="center">текс ошибки</td></tr>');
                for (let i = 0; i < resultData.length; i++) {
                    if (resultData[i].type == 'ERROR') {
                        error_num++;
                        res.write('<tr><td>' + resultData[i].id + '</td><td>' + resultData[i].fullName + '</td><td>' + resultData[i].text + '</td></tr>');
                    }
                }
                res.write('</table>');

                res.write('<table border="1px" width="60%" bordercolor="#9e9e9e" style="float: right; border-collapse: collapse"><tr><td align="center" colspan="3">Строки без ошибок</td></tr><tr><td align="center">id</td><td align="center">полное имя</td><td align="center">текст успеха</td></tr>');
                for (let i = 0; i < resultData.length; i++) {
                    if (resultData[i].type == 'SUCCESS') {
                        success_num++;
                        res.write('<tr><td>' + resultData[i].id + '</td><td>' + resultData[i].fullName + '</td><td>' + resultData[i].text + '</td></tr>');
                    }
                }
                res.write('</table></br></br></br></br></br></br></br></br></br></br></br></br></br></br></br>');
                res.write('Общее колличесвто строк: ' + resultData.length + '</br>');
                res.write('Коллиство успешных строк: ' + success_num + '</br>');
                res.write('Колличество ошибочных строк: ' + error_num);
                res.end();
            })
    });
});

async function parseData(sourse, promiseArr, resultArr) {
    for (let row = 2; ; row++) {
        if (!sourse['A' + row] && !sourse['B' + row] && !sourse['C' + row] && !sourse['D' + row])
            break;
        else {
            var id = '', lastName = '', firstName = '', patronymic = '', divisionName = '', subdivisionName = '', post = '', location = '', workPhone = '', territory = '';

            if(sourse['A' + row])
                id = sourse['A' + row].v;
            if(sourse['B' + row])
                lastName = sourse['B' + row].v;
            if(sourse['C' + row])
                firstName = sourse['C' + row].v;
            if(sourse['D' + row])
                patronymic = sourse['D' + row].v;
            if(sourse['E' + row])
                divisionName = sourse['E' + row].v;
            if(sourse['F' + row])
                subdivisionName = sourse['F' + row].v;
            if(sourse['G' + row])
                post = sourse['G' + row].v;
            if(sourse['H' + row])
                location = sourse['H' + row].v;
            if(sourse['I' + row])
                workPhone = sourse['I' + row].v;
            if(sourse['J' + row])
                territory = sourse['J' + row].v;

            // ---------- DIVISIONS ----------
            var division_id;
            const findDivision = await divisions_translations.findOne({where: { name: divisionName }})
                .then(result => {
                    return result;
                });

            if(findDivision){
                division_id = findDivision.dataValues.divisions_id;
            }
            else {
                division_id = await divisions_model.build({created_at: new Date(), updated_at: new Date()}).save()
                    .then(result => {
                       return result.dataValues.id;
                    });

                await divisions_translations.build({ name: divisionName, created_at: new Date(), updated_at: new Date(), divisions_id: division_id, languages_id: 1}).save()
                    .catch(err => {
                        console.error(err.message);
                    });
            }





            // ---------- SUB_DIVISIONS ----------
            var subdivision_id;
            const findSubdivision = await subdivisions_translations.findOne({where: { name: subdivisionName }})
                .then(result => {
                    return result;
                });

            if(findSubdivision){
                subdivision_id = findSubdivision.dataValues.subdivisions_id;
            }
            else {
                subdivision_id = await subdivisions_model.build({created_at: new Date(), updated_at: new Date()}).save()
                    .then(result => {
                        return result.dataValues.id;
                    });

                await subdivisions_translations.build({ name: subdivisionName, created_at: new Date(), updated_at: new Date(), subdivisions_id: subdivision_id, languages_id: 1}).save()
                    .catch(err => {
                        console.error(err.message);
                    });
            }




            // ---------- POLICEMANS ----------
            await sequelize.query('insert into policemans (full_name, post, work_phone, location, work_territory, created_at, updated_at, subdivisions_id, divisions_id) VALUES($1, $2, $3, $4, $5, $6, $7, $8, $9)',
                { bind: [lastName + ' ' + firstName + ' ' + patronymic,
                    post,
                    workPhone,
                    location,
                    territory,
                    new Date(),
                    new Date(),
                    subdivision_id,
                    division_id
                ]})
                .then(result => {
                    resultArr.push({type: 'SUCCESS', id: id, fullName: lastName + ' ' + firstName + ' ' + patronymic, text: "Строка " + row + " has been added successfully"});
                })
                .catch(err => {
                    resultArr.push({type: 'ERROR', id: id, fullName: lastName + ' ' + firstName + ' ' + patronymic, text: "Строка " + row + ": " + err.message});
                });
        }
    }
}

app.listen(3000, function () {
    console.log('listening on port 3000');
});