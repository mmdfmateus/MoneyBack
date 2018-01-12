var express = require("express");
var firebase = require("firebase");
var xl = require("excel4node");
var nodemailer = require("nodemailer");

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

router.post("/diversos", function(req, res, next) {
  var wb = new xl.Workbook();

  var headerStyle = wb.createStyle({
    font: {
      color: "#212121",
      bold: true,
      size: 14
    }
  });

  var style = wb.createStyle({
    font: {
      color: "#212121",
      size: 12
    }
  });

  var ws = wb.addWorksheet("Diversos");

  ws
    .cell(1, 1)
    .string("Natureza")
    .style(headerStyle);

  ws
    .cell(2, 1)
    .string(req.body.natureza)
    .style(style);
  ws
    .cell(1, 2)
    .string("Valor")
    .style(headerStyle);

  ws
    .cell(2, 2)
    .number(parseFloat(req.body.valor))
    .style(
      wb.createStyle({
        font: {
          color: "#212121",
          size: 12
        },
        numberFormat: "R$#,##0.00; (R$#,##0.00); -"
      })
    );

  ws
    .cell(1, 3)
    .string("Nota fiscal")
    .style(headerStyle);

  ws
    .cell(2, 3)
    .string(req.body.notaFiscal)
    .style(style);

  ws
    .cell(1, 4)
    .string("Solicitante")
    .style(headerStyle);

  ws
    .cell(2, 4)
    .string(req.body.nome)
    .style(style);

  wb.writeToBuffer().then(function(buffer) {
    var transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: "moneybackbot@gmail.com",
        pass: "moneybackpass"
      }
    });

    var mailOptions = {
      from: "moneybackbot@gmail.com",
      to: "brenof@take.net",
      subject: "Solicitação de reembolso - " + req.body.nome,
      text: "Segue despesa em anexo",
      attachments: [
        {
          filename: "solicitacao-reembolso.xlsx",
          content: buffer
        }
      ]
    };

    transporter.sendMail(mailOptions, function(error, info) {
      if (error) {
        console.log(error);
      } else {
        console.log("Email sent: " + info.response);
      }
    });
  });

  res.json(req.body);
});

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

  firebase
    .database()
    .ref("travel/" + identity)
    .once("value")
    .then(function(snapshot) {
      var viagem = snapshot.val();
      var wb = new xl.Workbook();

      var headerStyle = wb.createStyle({
        font: {
          color: "#212121",
          bold: true,
          size: 14
        }
      });

      var style = wb.createStyle({
        font: {
          color: "#212121",
          size: 12
        }
      });

      var ws = wb.addWorksheet("Reembolso");

      ws
        .cell(1, 1)
        .string("Data inicio")
        .style(headerStyle);

      ws
        .cell(2, 1)
        .string(viagem.dataInicio)
        .style(style);

      ws
        .cell(1, 2)
        .string("Data fim")
        .style(headerStyle);

      ws
        .cell(2, 2)
        .string(req.body.dataFim)
        .style(style);

      ws
        .cell(1, 3)
        .string("Destino")
        .style(headerStyle);

      ws
        .cell(2, 3)
        .string(viagem.destino)
        .style(style);

      ws
        .cell(1, 4)
        .string("Centro de custo")
        .style(headerStyle);

      ws
        .cell(2, 4)
        .string(viagem.centroDeCusto)
        .style(style);

      ws
        .cell(1, 5)
        .string("Solicitante")
        .style(headerStyle);

      ws
        .cell(2, 5)
        .string(viagem.nome)
        .style(style);
      ws
        .cell(4, 1)
        .string("Natureza")
        .style(headerStyle);

      ws
        .cell(4, 2)
        .string("Valor")
        .style(headerStyle);

      ws
        .cell(4, 3)
        .string("Nota fiscal")
        .style(headerStyle);

      var keys = Object.keys(viagem.items);
      for (var i = 0; i < keys.length; i++) {
        var item = viagem.items[keys[i]];

        ws
          .cell(5 + i, 1)
          .string(item.natureza)
          .style(style);

        ws
          .cell(5 + i, 2)
          .number(parseFloat(item.valorGasto))
          .style(
            wb.createStyle({
              font: {
                color: "#212121",
                size: 12
              },
              numberFormat: "R$#,##0.00; (R$#,##0.00); -"
            })
          );

        ws
          .cell(5 + i, 3)
          .string(item.notaFiscal)
          .style(style);
      }

      wb.writeToBuffer().then(function(buffer) {
        var transporter = nodemailer.createTransport({
          service: "gmail",
          auth: {
            user: "moneybackbot@gmail.com",
            pass: "moneybackpass"
          }
        });

        var mailOptions = {
          from: "moneybackbot@gmail.com",
          to: "brenof@take.net",
          subject: "Solicitação de reembolso - " + viagem.nome,
          text: "Segue despesas em anexo",
          attachments: [
            {
              filename: "solicitacao-reembolso.xlsx",
              content: buffer
            }
          ]
        };

        transporter.sendMail(mailOptions, function(error, info) {
          if (error) {
            console.log(error);
          } else {
            console.log("Email sent: " + info.response);

            var ref = firebase.database().ref("travel");
            ref.child(identity).remove();
          }
        });
      });

      res.json("done");
    });
});

module.exports = router;
