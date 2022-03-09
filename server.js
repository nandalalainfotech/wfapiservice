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


app.get('/', cors(), function (req, res) {
  console.log("Calling1----->",req);
  console.log("Calling2------>",res);
  res.send('Hello World!!!!!!')
});



app.get('/api/pdf', cors(), function (req, res) {     
  console.log("Calling1----->",req);
  console.log("Calling2------>",res);
  
  var options = {
      format: "A3",
      orientation: "landscape",
      border: "10mm",
  };

  var document = {
      type: 'buffer',
      template: html,
      context: {
          Wellsfargo:wellsfargo
      },
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



app.get('/api/excel', cors(), function (req, res) {

console.log("calling----------->excel",);

let workbook = new excel.Workbook();

let worksheet = workbook.addWorksheet('Quote Form');
  res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
      "Content-Disposition",
      "attachment; filename=" + "wellfargo" + ".xlsx"
  );
  return workbook.xlsx.write(res).then(function () {
      res['status'](200).end();
  });

});

app.listen(PORT, function () {
  console.log('App listening on port ' + PORT);
});



