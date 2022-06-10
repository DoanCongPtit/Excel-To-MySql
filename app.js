const express = require('express')
const fs = require('fs')
const path = require('path')
const bodyparser = require('body-parser')
const mysql = require('mysql')
const multer = require('multer')
const app = express()
const XLSX = require("xlsx")

app.set('view engine', 'ejs');
app.use(express.static('./public'))
app.use(bodyparser.json())
app.use(
    bodyparser.urlencoded({
        extended: true,
    }),
)

const db = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '123456789',
    database: 'excel',
})

db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message)
    }
    console.log('Database connected.')
})

// SET STORAGE
let storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads')
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname)
    }
})


let upload = multer({ storage: storage })

app.get('/', (req, res) => {
    res.render(__dirname + '/views/index.ejs')
})

app.post('/uploadfile', upload.single('myFile'), (req, res, next) => {
    const file = req.file
    if (!file) {
        const error = new Error('Please upload a file')
        error.httpStatusCode = 400
        return next(error)
    }

    let workbook = XLSX.readFile("uploads/node.xlsx");
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    let rowsCount = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[workbook.SheetNames[0]]);
    console.log(rowsCount.length)
    for (let index = 2; index < rowsCount.length+2; index++) {
        let id = worksheet[`A${index}`].v;
        let name = worksheet[`B${index}`].v;
        let code = worksheet[`C${index}`].v;

        console.log({ id: id, name: name, code: code })
        let query = "INSERT INTO user (id, name, code) VALUES (?)";
        let values = [id, name, code];

        db.query(query, [values], (error, response) => {
            console.log(error || response)
        })
    }

    res.render(__dirname + '/views/index.ejs')
})

app.get('/search', (req, res) => {
    let name = req.query.name;
    let code = req.query.code;

    let query = "select * from user where name LIKE '%" + name + "%' AND code LIKE '%" + code + "%'";
    db.query(query, (err, result) => {
        if (err) console.log(err);
        console.log(result)
        res.render(__dirname + '/views/result.ejs', { user: result });
    })
})


let nodeServer = app.listen(4000, function () {
    let port = nodeServer.address().port
    let host = nodeServer.address().address
    console.log('App working on: ', host, port)
})