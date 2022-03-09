const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const excel = require('exceljs');
const fs = require('fs');
var pdf = require('dynamic-html-pdf');
var html = fs.readFileSync('./template/html-pdf-template.html', 'utf8');

const wellsfargoJson = fs.readFileSync('./json/wellsfargo.json', 'utf8');
let wellsfargo = JSON.parse(wellsfargoJson);
var express = require('express');
var app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(express.static('public'));


var PORT = process.env.PORT || 80;

app.get('/', function (req, res) {
  console.log("Calling - /api/pdf");
  res.send('Hello World!!!!!!')
});

app.get('/api/pdf', cors(), function (req, res) {  
    
  console.log("Calling - /api/pdf");
  
  var options = {
      format: "A3",
      orientation: "landscape",
      border: "10mm",
      // phantomPath: "node_modules/phantomjs-prebuilt/lib/phantom/bin/phantomjs",
      // phantomPath: "node_modules/phantomjs-prebuilt/bin/phantomjs" 
  };

  var document = {
      type: 'buffer',     // 'file' or 'buffer'
      template: html,
      context: {
          Wellsfargo:wellsfargo
      },
      // path: "./html-pdf-template.pdf"    // it is not required if type is buffer
  };

  if (document === null) {
      return null;

  } else {
    pdf.create(document, options).then(response => {
      res.writeHead(200, {
          "Content-Disposition": "attachment;filename=" + "wellsFargo.pdf",
          'Content-Type': 'application/pdf'
      });
      return res.end(response);
  }).catch(error => {
      console.error(error)
  });
};

});

app.listen(PORT, function () {
  console.log('App listening on port ' + PORT);
});