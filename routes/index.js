var express = require("express");
var firebase = require("firebase");

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

/* POST home page. */
router.post("/", function(req, res, next) {
  firebase
    .database()
    .ref("travel/" + encodeURIComponent(req.body.identity.split('@')[0]))
    .set(req.body); 

  res.json(req.body);
});



module.exports = router;
