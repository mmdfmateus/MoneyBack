var express = require("express");
var firebase = require("firebase");
var xl = require('excel4node');

var router = express.Router();

// Initialize Firebase
var config = {
  apiKey: "AIzaSyCCitS6JiGyv26-eB7VuuQisyrr8LTRkc0",
  authDomain: "moneyback-60e24.firebaseapp.com",
  databaseURL: "https://moneyback-60e24.firebaseio.com",
  projectId: "moneyback-60e24",
  storageBucket: "moneyback-60e24.appspot.com",
  messagingSenderId: "677380462491"
};

firebase.initializeApp(config);

function parseIdentity(identity) {
  return identity
    .split("@")[0]
    .replace(/_/g, "")
    .replace(/\./g, "")
    .replace(/-/g, "");
}

router.post("/", function(req, res, next) {
  var identity = parseIdentity(req.body.identity);

  firebase
    .database()
    .ref("travel/" + identity)
    .set(req.body);

  res.json(req.body);
});

router.post("/items", function(req, res, next) {
  var identity = parseIdentity(req.body.identity);

  firebase
    .database()
    .ref("travel/" + identity)
    .child("items")
    .push(req.body);

  res.json(req.body);
});

router.post("/close", function(req, res, next) {
  var identity = parseIdentity(req.body.identity);

  var viagem = firebase
                .database()
                .ref("travel/" + identity)
                .toJSON();
  
  var wb = new xl.Workbook();

  var style = wb.createStyle({
    font: {
        color: '#FF0800',
        size: 12
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -'
  });

  var ws = wb.addWorksheet('Sheet 1');

  ws.cell(1,1).string('teste1').style(style);
  ws.cell(1,2).string('teste2').style(style);
  wb.write('Excel.xlsx');

  res.json('req.body');
});

module.exports = router;
