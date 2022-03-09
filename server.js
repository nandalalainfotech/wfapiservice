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
  res.send('Hello World!!!!!!')
});



app.get('/api/pdf', cors(), function (req, res) {     
  // console.log("Calling1----->",req);
  // console.log("Calling2------>",res);
  
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
  // console.log("Calling3----->",req);
  // console.log("Calling4------>",res);

let workbook = new excel.Workbook();
let worksheet = workbook.addWorksheet('Quote Form');

// border none
worksheet.views = [{ showGridLines: false }];

worksheet.getRow(1).height = 40;
worksheet.getRow(2).height = 30;
 // worksheet.getRow(3).height = 25;
 // worksheet.getRow(4).height = 20;
 // worksheet.getRow(5).height = 20;
 // worksheet.getRow(6).height = 20;
 // worksheet.getRow(7).height = 20;
 // worksheet.getRow(8).height = 20;
 // worksheet.getRow(9).height = 20;
 worksheet.getRow(10).height = 40;
  worksheet.getRow(11).height = 20;
 worksheet.getRow(12).height = 25;
 worksheet.getRow(13).height = 25;
 worksheet.getRow(14).height = 80;
 // worksheet.getRow(15).height = 20;
 // worksheet.getRow(16).height = 20;
 worksheet.getRow(17).height = 170;
 worksheet.getRow(18).height = 40;
 worksheet.getRow(19).height = 120;
 worksheet.getRow(20).height = 60;
 worksheet.getRow(21).height = 60;
 worksheet.getRow(22).height = 70;
 worksheet.getRow(23).height = 70;
worksheet.getRow(24).height = 70;
 worksheet.getRow(25).height = 70;
 worksheet.getRow(26).height = 60;
 worksheet.getRow(27).height = 60;
 worksheet.getRow(28).height = 60;
 worksheet.getRow(29).height = 60;
 worksheet.getRow(30).height = 60;
 worksheet.getRow(31).height = 60;
 worksheet.getRow(32).height = 60;
 worksheet.getRow(33).height = 60;
 worksheet.getRow(34).height = 60;
 worksheet.getRow(35).height = 60;
 worksheet.getRow(36).height = 100;
 worksheet.getRow(37).height = 30;
 worksheet.getRow(39).height = 30;
 worksheet.getRow(40).height = 30;
 worksheet.getRow(41).height = 40;
 worksheet.getRow(42).height = 20;
 worksheet.getRow(43).height = 20;
 worksheet.getRow(44).height = 20;
 worksheet.getRow(45).height = 20;
 worksheet.getRow(46).height = 20;
 worksheet.getRow(47).height = 30;
 worksheet.getRow(49).height = 30;
 worksheet.getRow(50).height = 150;
 worksheet.getRow(52).height = 25;
 worksheet.getRow(53).height = 30;
 worksheet.getRow(54).height = 30;
 worksheet.getRow(55).height = 30;
worksheet.getRow(56).height = 30;
 worksheet.getRow(57).height = 30;
 worksheet.getRow(58).height = 30;
 worksheet.getRow(59).height = 30;
 worksheet.getRow(60).height = 30;
 worksheet.getRow(61).height = 30;
 worksheet.getRow(62).height = 30;
 worksheet.getRow(63).height = 30;




 worksheet.columns = [{ key: 'A', width: 5.0 }, { key: 'B', width: 10.0 }, { key: 'C', width: 18.0 },
 { key: 'D', width: 18.0 }, { key: 'E', width: 15.0 }, { key: 'F', width: 20.0 }, { key: 'G', width: 15.0 },
 { key: 'H', width: 15.0 }, { key: 'I', width: 15.0 }, { key: 'J', width: 20.0 }, { key: 'K', width: 15.0 },
 { key: 'L', width: 23.0 }, { key: 'M', width: 15.0 }, { key: 'N', width: 2.0 }, { key: 'O', width: 15.0 },
 { key: 'P', width: 15.0 }, { key: 'Q', width: 15.0 }, { key: 'R', width: 23.0 }, { key: 'S', width: 17.0 },
 { key: 'T', width: 15.0 }, { key: 'U', width: 20.0 }];

// worksheet.columns.forEach((col) => {
//     col.style.font = {
//         size: 10,
//         bold: true
//     };
//     col.style.alignment = { vertical: 'middle', horizontal: 'center' };
//     col.style.border = { top: { style: 'thin' }, left: { style: 'thin' },
//      bottom: { style: 'thin' }, right: { style: 'thin' } };
// })


//insert an image B2:D6 with width
// worksheet.addImage(imageId1,
//     {
//         tl: { col: 1, row: 0 },
//         ext: {
//             width: 50, height: 50
//         }
//     }
// );
// add image to workbook by filename
const imageId1 = workbook.addImage({
    filename: './images/wellsfargo.png',
    extension: 'png',
});
// insert an image over B2:D6
worksheet.addImage(imageId1, 'B1:B1',);

worksheet.mergeCells('B1:U1');
worksheet.getCell('B1:U1').value = wellsfargo.mainTitle;
worksheet.getCell('B1:U1').font = {
    size: 28,
    name: 'Verdana',
    family: 1

};
worksheet.getCell('B1:U1').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('B1:U1').border = {
    // top: {style:'thin'},
    // left: {style:'none'},
    // bottom: {style:'thin'},
    right: { style: 'thin' }
};
// --------------------------COMMON-------------------------

// worksheet.mergeCells('K3:K10');

['B18:E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'O18', 'P18', 'Q18', 'R18', 'S18', 'T18', 'U18',
    'B20:E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'O20', 'P20', 'Q20', 'R20', 'S20', 'T20', 'U20',
    'B22:E22', 'F22', 'G22', 'H22', 'I22', 'J22', 'K22', 'L22', 'M22', 'O22', 'P22', 'Q22', 'R22', 'S22', 'T22', 'U22',
    'B24:E24', 'F24', 'G24', 'H24', 'I24', 'J24', 'K24', 'L24', 'M24', 'O24', 'P24', 'Q24', 'R24', 'S24', 'T24', 'U24',
    'B26:E26', 'F26', 'G26', 'H26', 'I26', 'J26', 'K26', 'L26', 'M26', 'O26', 'P26', 'Q26', 'R26', 'S26', 'T26', 'U26',
    'B28:E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'O28', 'P28', 'Q28', 'R28', 'S28', 'T28', 'U28',
    'B30:E30', 'F30', 'G30', 'H30', 'I30', 'J30', 'K30', 'L30', 'M30', 'O30', 'P30', 'Q30', 'R30', 'S30', 'T30', 'U30',
    'B32:E32', 'F32', 'G32', 'H32', 'I32', 'J32', 'K32', 'L32', 'M32', 'O32', 'P32', 'Q32', 'R32', 'S32', 'T32', 'U32',
    'B34:E34', 'F34', 'G34', 'H34', 'I34', 'J34', 'K34', 'L34', 'M34', 'O34', 'P34', 'Q34', 'R34', 'S34', 'T34', 'U34',
    'B42:G42', 'H42', 'I42', 'J42', 'K42', 'L42', 'M42', 'O42', 'P42', 'Q42', 'R42', 'S42', 'T42', 'U42', 'L43', 'U43',
    'B44:G44', 'H44', 'I44', 'J44', 'K44', 'L44', 'M44', 'O44', 'P44', 'Q44', 'R44', 'S44', 'T44', 'U44', 'L45', 'U45',
    'B46:G46', 'H46', 'I46', 'J46', 'K46', 'L46', 'M46', 'O46', 'P46', 'Q46', 'R46', 'S46', 'T46', 'U46',
    'T4:U4', 'T5:U5', 'T6:U6', 'T7:U7', 'T8:U8'
].map(key => {
    worksheet.getCell(key).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
});


['A1', 'A2'].map(key => {
    worksheet.getCell(key).font = {
        size: 11,
        name: 'Verdana',
        family: 1

    };
});
// -----------------------------------------------------------


worksheet.mergeCells('B2:Q2');
worksheet.getCell('B2:Q2').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};

worksheet.mergeCells('R2:S2');
worksheet.getCell('R2:S2').value = "Date Submitted";
worksheet.getCell('R2:S2').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R2:S2').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('R2:S2').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('T2:U2');
worksheet.getCell('T2:U2').value = wellsfargo.dateSubmitted;
worksheet.getCell('T2:U2').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('T2:U2').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('T2:U2').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('T2:U2').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('B3:J3');
worksheet.getCell('B3:J3').value = wellsfargo.tableOneTitle;
worksheet.getCell('B3:J3').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B3:J3').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B3:J3').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('B3:J3').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('B4:D4');
worksheet.getCell('B4:D4').value = "Project # or Work Order #";
worksheet.getCell('B4:D4').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B4:D4').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B4:D4').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('E4:J4');
worksheet.getCell('E4:J4').value = wellsfargo.projectOrWorkOrder;
worksheet.getCell('E4:J4').font = {
    size: 14,
    name: 'Verdana',
    family: 1
    // bold: true
};
worksheet.getCell('E4:J4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('E4:J4').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E4:J4').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};




worksheet.mergeCells('B5:D5');
worksheet.getCell('B5:D5').value = "WF Project/ Property Manager";
worksheet.getCell('B5:D5').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B5:D5').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B5:D5').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};


worksheet.mergeCells('E5:J5');
worksheet.getCell('E5:J5').value = wellsfargo.wfProjectOrPropertyManager;
worksheet.getCell('E5:J5').font = {
    size: 14,
    name: 'Verdana',
    family: 1
    // bold: true
};
worksheet.getCell('E5:J5').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('E5:J5').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E5:J5').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('B6:D6');
worksheet.getCell('B6:D6').value = "BE Number: ";
worksheet.getCell('B6:D6').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B6:D6').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B6:D6').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('E6:J6');
worksheet.getCell('E6:J6').value = wellsfargo.beNumber;
worksheet.getCell('E6:J6').font = {
    size: 14,
    name: 'Verdana',
    family: 1
    // bold: true
};
worksheet.getCell('E6:J6').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('E6:J6').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E6:J6').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('B7:D7');
worksheet.getCell('B7:D7').value = "Building / Project Name:";
worksheet.getCell('B7:D7').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B7:D7').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B7:D7').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('E7:J7');
worksheet.getCell('E7:J7').value = wellsfargo.buildingOrProjectName;
worksheet.getCell('E7:J7').font = {
    size: 14,
    name: 'Verdana',
    family: 1
    // bold: true
};
worksheet.getCell('E7:J7').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('E7:J7').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E7:J7').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('B8:D8');
worksheet.getCell('B8:D8').value = "BE Service or Delivery Address";
worksheet.getCell('B8:D8').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B8:D8').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B8:D8').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('E8:J8');
worksheet.getCell('E8:J8').value = wellsfargo.beServiceOrDeliveryAddress;
worksheet.getCell('E8:J8').font = {
    size: 14,
    name: 'Verdana',
    family: 1
    // bold: true
};
worksheet.getCell('E8:J8').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E8:J8').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};
worksheet.getCell('E8:J8').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};

worksheet.mergeCells('B9:D9');
worksheet.getCell('B9:D9').value = "Project Area (sq.ft.):";
worksheet.getCell('B9:D9').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B9:D9').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('B9:D9').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('E9:J9');
worksheet.getCell('E9:J9').value = wellsfargo.projectArea;
worksheet.getCell('E9:J9').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    // bold: true
};
worksheet.getCell('E9:J9').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('E9:J9').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('E9:J9').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('B10:C10');
worksheet.getCell('B10:C10').value = "Estimated Start Date:";
worksheet.getCell('B10:C10').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B10:C10').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('B10:C10').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('D10:E10');
worksheet.getCell('D10:E10').value = wellsfargo.estimatedStartDate;
worksheet.getCell('D10:E10').font = {
    size: 14,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('D10:E10').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('D10:E10').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('D10:E10').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('F10:G10');
worksheet.getCell('F10:G10').value = "Estimated Complete Date:";
worksheet.getCell('F10:G10').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('F10:G10').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('F10:G10').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('H10:J10');
worksheet.getCell('H10:J10').value = wellsfargo.estimatedCompleteDate;
worksheet.getCell('H10:J10').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('H10:J10').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('H10:J10').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('H10:J10').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};
//    ---------------------------2 row---------------------------

worksheet.mergeCells('L3:U3');
worksheet.getCell('L3:U3').value = wellsfargo.tableTwoTitle;
worksheet.getCell('L3:U3').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('L3:U3').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('L3:U3').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('L3:U3').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('L4:M4');
worksheet.getCell('L4:M4').value = "Company Name:";
worksheet.getCell('L4:M4').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L4:M4').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L4:M4').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N4:Q4');
worksheet.getCell('N4:Q4').value = wellsfargo.companyName;
worksheet.getCell('N4:Q4').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N4:Q4').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N4:Q4').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N4:Q4').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R4:S4');
worksheet.getCell('R4:S4').value = "WF Vendor Number:";
worksheet.getCell('R4:S4').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R4:S4').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('R4:S4').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T4:U4');
worksheet.getCell('T4:U4').value = wellsfargo.wfVendOrNumber;
worksheet.getCell('T4:U4').font = {
    size: 14,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('T4:U4').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('T4:U4').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L5:M5');
worksheet.getCell('L5:M5').value = "Remit To Address:";
worksheet.getCell('L5:M5').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L5:M5').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L5:M5').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N5:Q5');
worksheet.getCell('N5:Q5').value = wellsfargo.remitToAddress;
worksheet.getCell('N5:Q5').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N5:Q5').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N5:Q5').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N5:Q5').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R5:S5');
worksheet.getCell('R5:S5').value = "Proposal Number:";
worksheet.getCell('R5:S5').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R5:S5').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('R5:S5').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T5:U5');
worksheet.getCell('T5:U5').value = wellsfargo.proposalNumber;
worksheet.getCell('T5:U5').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('T5:U5').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('T5:U5').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L6:M6');
worksheet.getCell('L6:M6').value = " City, State, Zip :";
worksheet.getCell('L6:M6').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L6:M6').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L6:M6').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N6:Q6');
worksheet.getCell('N6:Q6').value = wellsfargo.cityStateZip;
worksheet.getCell('N6:Q6').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N6:Q6').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N6:Q6').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N6:Q6').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R6:S6');
worksheet.getCell('R6:S6').value = "WF Contract Number:";
worksheet.getCell('R6:S6').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R6:S6').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('R6:S6').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T6:U6');
worksheet.getCell('T6:U6').value = wellsfargo.wfContractNumber;
worksheet.getCell('T6:U6').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('T6:U6').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('T6:U6').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};



worksheet.mergeCells('L7:M7');
worksheet.getCell('L7:M7').value = "Contact Name:";
worksheet.getCell('L7:M7').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L7:M7').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L7:M7').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N7:Q7');
worksheet.getCell('N7:Q7').value = wellsfargo.contactName;
worksheet.getCell('N7:Q7').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N7:Q7').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N7:Q7').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N7:Q7').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R7:S7');
worksheet.getCell('R7:S7').value = "Change Order #:";
worksheet.getCell('R7:S7').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R7:S7').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('R7:S7').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T7:U7');
worksheet.getCell('T7:U7').value = wellsfargo.changeOrder;
worksheet.getCell('T7:U7').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('T7:U7').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('T7:U7').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L8:M8');
worksheet.getCell('L8:M8').value = "Phone";
worksheet.getCell('L8:M8').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L8:M8').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L8:M8').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N8:Q8');
worksheet.getCell('N8:Q8').value = wellsfargo.phone;
worksheet.getCell('N8:Q8').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N8:Q8').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N8:Q8').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N8:Q8').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R8:S8');
worksheet.getCell('R8:S8').value = "Change Order Previous PO#:";
worksheet.getCell('R8:S8').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R8:S8').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('R8:S8').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T8:U8');
worksheet.getCell('T8:U8').value = wellsfargo.changeOrderPreviousPO;
worksheet.getCell('T8:U8').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('T8:U8').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('T8:U8').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L9:M9');
worksheet.getCell('L9:M9').value = "Cell:";
worksheet.getCell('L9:M9').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L9:M9').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L9:M9').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N9:U9');
worksheet.getCell('N9:U9').value = wellsfargo.cell;
worksheet.getCell('N9:U9').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N9:U9').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N9:U9').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N9:U9').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L10:M10');
worksheet.getCell('L10:M10').value = "Email:";
worksheet.getCell('L10:M10').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L10:M10').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('L10:M10').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('N10:U10');
worksheet.getCell('N10:U10').value = wellsfargo.email;
worksheet.getCell('N10:U10').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('N10:U10').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('N10:U10').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('N10:U10').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

//    ------------------------------------------------------

// ------------------------------3 row--------------------------
worksheet.mergeCells('B11:U11');

// ------------------------------4 row--------------------------

worksheet.mergeCells('B12:U12');
worksheet.getCell('B12:U12').value = wellsfargo.scopeTitle;
worksheet.getCell('B12:U12').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B12:U12').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B12:U12').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('B12:U12').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('B13:U13');
worksheet.getCell('B13:U13').value = wellsfargo.scopeSubHeadOne;
worksheet.getCell('B13:U13').font = {
    size: 11,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B13:U13').alignment = { vertical: 'top', horizontal: 'center', wrapText: true };
worksheet.getCell('B13:U13').border = {
    top: { style: 'none' },
    left: { style: 'thick' },
    bottom: { style: 'none' },
    right: { style: 'thick' }
};



worksheet.mergeCells('B14:U14');
worksheet.getCell('B14:U14').value = {
    'richText': [

        { 'font': { 'size': 14, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.scopeSubDescription_D1 },

        { 'font': { 'size': 14, 'name': 'Verdana', 'family': 1, 'color': { 'argb': 'FF0000' } }, 'text': wellsfargo.scopeSubDescription_D2 }

    ]
};
// worksheet.getCell('B14:U14').value = wellsfargo.scopeSubDescription;
worksheet.getCell('B14:U14').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    // bold: true
};
worksheet.getCell('B14:U14').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'C5D9F1' },
    bgColor: { argb: 'C5D9F1' }
};
worksheet.getCell('B14:U14').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('B14:U14').border = {
    top: { style: 'none' },
    left: { style: 'thick' },
    bottom: { style: 'none' },
    right: { style: 'thick' }
};

worksheet.mergeCells('B15:U15');
worksheet.getCell('B15:U15').value = wellsfargo.installHeading;
worksheet.getCell('B15:U15').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B15:U15').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B15:U15').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('B15:U15').border = {
    top: { style: 'none' },
    left: { style: 'thick' },
    bottom: { style: 'none' },
    right: { style: 'thick' }
};

worksheet.getRow(16).height = 25;
worksheet.mergeCells('B16:E16');
worksheet.getCell('B16:E16').value = wellsfargo.detailSubHeadOne;
worksheet.getCell('B16:E16').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B16:E16').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B16:E16').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('B16:E16').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('F16:M16');
worksheet.getCell('F16:M16').value = wellsfargo.productSubHeadTwo;
worksheet.getCell('F16:M16').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('F16:M16').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('F16:M16').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('F16:M16').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O16:T16');
worksheet.getCell('O16:T16').value = wellsfargo.laborSubHeadThree;
worksheet.getCell('O16:T16').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('O16:T16').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('O16:T16').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('O16:T16').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};



worksheet.mergeCells('B17:E17');
// worksheet.getCell('B17:E17').value = wellsfargo.installHeadingColumn1;
worksheet.getCell('B17:E17').value = {
    'richText': [
        { 'font': { 'bold': true, 'size': 11, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D1 },
        { 'font': { 'bold': true, 'size': 11, 'color': { 'argb': 'FF0000' }, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D2 },
        { 'font': { 'bold': true, 'size': 11, 'color': { 'argb': 'FF0000' }, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D3 }
    ]
};
worksheet.getCell('B17:E17').font = {
    size: 11,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B17:E17').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('B17:E17').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};




worksheet.mergeCells('F17');
worksheet.getCell('F17').value = wellsfargo.installHeadingColumn2;
worksheet.getCell('F17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('F17').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('F17').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};


worksheet.mergeCells('G17');
worksheet.getCell('G17').value = wellsfargo.installHeadingColumn3;
worksheet.getCell('G17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('G17').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('G17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};


worksheet.mergeCells('H17');
worksheet.getCell('H17').value = wellsfargo.installHeadingColumn4;
worksheet.getCell('H17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('H17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('H17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('I17');
worksheet.getCell('I17').value = wellsfargo.installHeadingColumn5;
worksheet.getCell('I17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('I17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('I17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('J17');
worksheet.getCell('J17').value = wellsfargo.installHeadingColumn6;
worksheet.getCell('J17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('J17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('J17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('K17');
worksheet.getCell('K17').value = wellsfargo.installHeadingColumn7;
worksheet.getCell('K17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('K17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('K17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('L17');
worksheet.getCell('L17').value = wellsfargo.installHeadingColumn8;
worksheet.getCell('L17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('L17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('M17');
worksheet.getCell('M17').value = wellsfargo.installHeadingColumn9;
worksheet.getCell('M17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('M17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('M17').border = {
    top: { style: 'thick' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('O17');
worksheet.getCell('O17').value = wellsfargo.installHeadingColumn10;
worksheet.getCell('O17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('O17').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('O17').border = {
    // top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};


worksheet.mergeCells('P17');
worksheet.getCell('P17').value = wellsfargo.installHeadingColumn11;
worksheet.getCell('P17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('P17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('P17').border = {
    // top: { style: 'thick' },
    // left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('Q17');
worksheet.getCell('Q17').value = wellsfargo.installHeadingColumn12;
worksheet.getCell('Q17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('Q17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('Q17').border = {
    // top: { style: 'thick' },
    // left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('R17');
worksheet.getCell('R17').value = wellsfargo.installHeadingColumn13;
worksheet.getCell('R17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('R17').border = {
    // top: { style: 'thick' },
    // left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S17');
worksheet.getCell('S17').value = wellsfargo.installHeadingColumn14;
worksheet.getCell('S17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('S17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('S17').border = {
    // top: { style: 'thick' },
    // left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T17');
worksheet.getCell('T17').value = wellsfargo.installHeadingColumn15;
worksheet.getCell('T17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('T17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('T17').border = {
    // top: { style: 'thick' },
    // left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('U16:U17');
worksheet.getCell('U16:U17').value = wellsfargo.installHeadingColumn16;
worksheet.getCell('U16:U17').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('U16:U17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('U16:U17').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A18');
worksheet.getCell('A18').value = "1.1";
worksheet.getCell('A18').font = {
    size: 11,
    name: 'Verdana',
    family: 1,
    // bold: true
};
worksheet.getCell('A18').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A19').value = "1.2";
worksheet.getCell('A19').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A19').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A20').value = "1.3";
worksheet.getCell('A20').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A20').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A21').value = "1.4";
worksheet.getCell('A21').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A21').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A22').value = "1.5";
worksheet.getCell('A22').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A22').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A23').value = "1.6";
worksheet.getCell('A23').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A23').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A24').value = "1.7";
worksheet.getCell('A24').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A24').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A25').value = "1.8";
worksheet.getCell('A25').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A25').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('A26').value = "1.9";
worksheet.getCell('A26').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A26').alignment = { vertical: 'middle', horizontal: 'left' };



// -----------------------------------value starting---------------------------

for (let i = 0; i < wellsfargo.installColumns.length; i++) {

    let temp = i + 18;
    const row = worksheet.getRow(temp);
    // row.height = 100;
    // row.width = 200;
    // console.log("width", row.width, "height", row.height);
    // row.height = row.height * 20 / row.width;
    // console.log("row.height", row.height);
    // console.log("row.width", row.width);

    // let height = 20;
    // const row = worksheet.getRow(temp);
    // console.log("row", row);

    // row.height = row.height * 20 / row.width;
    // console.log("  row.height", row.height)

    // const row = worksheet.getRow(temp);
    // worksheet.getRow(temp).height = 20 * row.width;
    // const row = worksheet.getRow(temp);
    // row.height = row.height * 150 / row.width;
    // const row = worksheet.getRow(temp + 4);
    // console.log("row.height------",row);
    // worksheet.views = [{}]
    // worksheet.properties.defaultRowHeight = 150;
    // row.height = row.height + 100;

    worksheet.mergeCells(temp, 2, temp, 5);
    worksheet.getCell(temp, 2, temp, 5).value = wellsfargo.installColumns[i].coloumn1;
    worksheet.getCell(temp, 2, temp, 5).font = {
        size: 11,
        name: 'Calibri',
        family: 1
    };
    worksheet.getCell(temp, 2, temp, 5).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell(temp, 2, temp, 5).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('F' + temp);
    worksheet.getCell('F' + temp).value = wellsfargo.installColumns[i].coloumn2;
    worksheet.getCell('F' + temp).font = {
        size: 11,
        name: 'Calibri',
        family: 1
    };
    worksheet.getCell('F' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('F' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('G' + temp);
    worksheet.getCell('G' + temp).value = wellsfargo.installColumns[i].coloumn3;
    worksheet.getCell('G' + temp).font = {
        size: 11,
        name: 'Calibri',
        family: 1
    };
    worksheet.getCell('G' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('G' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('H' + temp);
    worksheet.getCell('H' + temp).value = wellsfargo.installColumns[i].coloumn4;
    worksheet.getCell('H' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('H' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('H' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('I' + temp);
    worksheet.getCell('I' + temp).value = wellsfargo.installColumns[i].coloumn5;
    worksheet.getCell('I' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('I' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('I' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('J' + temp);
    worksheet.getCell('J' + temp).value = wellsfargo.installColumns[i].coloumn6;
    worksheet.getCell('J' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('J' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('J' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('K' + temp);
    worksheet.getCell('K' + temp).value = wellsfargo.installColumns[i].coloumn7;
    worksheet.getCell('K' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('K' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('K' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('L' + temp);
    worksheet.getCell('L' + temp).value = wellsfargo.installColumns[i].coloumn8;
    worksheet.getCell('L' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('L' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('L' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('M' + temp);
    worksheet.getCell('M' + temp).value = wellsfargo.installColumns[i].coloumn9;
    worksheet.getCell('M' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('M' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('M' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O' + temp);
    worksheet.getCell('O' + temp).value = wellsfargo.installColumns[i].coloumn10;
    worksheet.getCell('O' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('O' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('O' + temp).border = {
        // top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('P' + temp);
    worksheet.getCell('P' + temp).value = wellsfargo.installColumns[i].coloumn11;
    worksheet.getCell('P' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('P' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('P' + temp).border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('Q' + temp);
    worksheet.getCell('Q' + temp).value = wellsfargo.installColumns[i].coloumn12;
    worksheet.getCell('Q' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('Q' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('Q' + temp).border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R' + temp);
    worksheet.getCell('R' + temp).value = "";
    worksheet.getCell('R' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('R' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('R' + temp).border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('S' + temp);
    worksheet.getCell('S' + temp).value = wellsfargo.installColumns[i].coloumn14;
    worksheet.getCell('S' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + temp).border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + temp);
    worksheet.getCell('T' + temp).value = wellsfargo.installColumns[i].coloumn15;
    worksheet.getCell('T' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('T' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('T' + temp).border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('U' + temp);
    worksheet.getCell('U' + temp).value = wellsfargo.installColumns[i].coloumn16;
    worksheet.getCell('U' + temp).font = {
        size: 12,
        bold: true
    };
    worksheet.getCell('U' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell('U' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


}

// -----------------------------------value ending---------------------------


worksheet.mergeCells('B37:K37');
// worksheet.getCell('B37:K37').border = {
//     top: { style: 'thick' },
//     left: { style: 'thick' },
//     bottom: { style: 'none' },
//     right: { style: 'thick' }
// };

worksheet.mergeCells('L37');
worksheet.getCell('L37').value = wellsfargo.totalProduct;
worksheet.getCell('L37').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('L37').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('L37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O37');
worksheet.getCell('O37').value = wellsfargo.totalPeople;
worksheet.getCell('O37').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('O37').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('O37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('P37');
worksheet.getCell('P37').value = wellsfargo.totalHoursPerPerson;
worksheet.getCell('P37').font = {
    size: 12,
    bold: true
};
worksheet.getCell('P37').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('P37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('Q37');
worksheet.getCell('Q37').value = wellsfargo.totalHourlyBillRate;
worksheet.getCell('Q37').font = {
    size: 13,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('Q37').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('Q37').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('Q37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R37');
worksheet.getCell('R37').value = wellsfargo.totalUnionRate;
worksheet.getCell('R37').font = {
    size: 13,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R37').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('R37').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('R37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('S37');
worksheet.getCell('S37').value = wellsfargo.totalLabor;
worksheet.getCell('S37').font = {
    size: 13,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('S37').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('T37');
worksheet.getCell('T37').value = wellsfargo.totalLaborTaxable;
worksheet.getCell('T37').font = {
    size: 13,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('T37').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('T37').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('T37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('U37');
worksheet.getCell('U37').value = wellsfargo.totalProductAndLabor;
worksheet.getCell('U37').font = {
    size: 13,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('U37').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('U37').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};



worksheet.mergeCells('B38:U38');
worksheet.getCell('B38:U38').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};


worksheet.mergeCells('B39:T39');
worksheet.getCell('B39:T39').value = wellsfargo.demoHeading;
worksheet.getCell('B39:T39').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B39:T39').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B39:T39').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
worksheet.getCell('B39:T39').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};



worksheet.mergeCells('B40:M40');
worksheet.getCell('B40:M40').value = wellsfargo.demoSubHeadOne;
worksheet.getCell('B40:M40').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B40:M40').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B40:M40').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('B40:M40').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O40:T40');
worksheet.getCell('O40:T40').value = wellsfargo.demoSubHeadTwo;
worksheet.getCell('O40:T40').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('O40:T40').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('O40:T40').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('O40:T40').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.getRow(41).height = 50;
worksheet.mergeCells('B41:G41');
worksheet.getCell('B41:G41').value = wellsfargo.demoHeadingColumn1;
worksheet.getCell('B41:G41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FF0000' },
    bold: true
};
worksheet.getCell('B41:G41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('B41:G41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('H41');
worksheet.getCell('H41').value = "";
worksheet.getCell('H41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('H41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('H41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('I41');
worksheet.getCell('I41').value = "";
worksheet.getCell('I41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('I41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('I41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('J41');
worksheet.getCell('J41').value = wellsfargo.demoHeadingColumn4;
worksheet.getCell('J41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('J41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('J41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('K41');
worksheet.getCell('K41').value = wellsfargo.demoHeadingColumn5;
worksheet.getCell('K41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('K41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('K41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('L41');
worksheet.getCell('L41').value = wellsfargo.demoHeadingColumn6;
worksheet.getCell('L41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('L41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('M41');
worksheet.getCell('M41').value = wellsfargo.demoHeadingColumn7;
worksheet.getCell('M41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('M41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('M41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O41');
worksheet.getCell('O41').value = wellsfargo.demoHeadingColumn8;
worksheet.getCell('O41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('O41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('O41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('P41');
worksheet.getCell('P41').value = wellsfargo.demoHeadingColumn9;
worksheet.getCell('P41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('P41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('P41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('Q41');
worksheet.getCell('Q41').value = wellsfargo.demoHeadingColumn10;
worksheet.getCell('Q41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('Q41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('Q41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R41');
worksheet.getCell('R41').value = wellsfargo.demoHeadingColumn11;
worksheet.getCell('R41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('R41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('S41');
worksheet.getCell('S41').value = wellsfargo.demoHeadingColumn12;
worksheet.getCell('S41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('S41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('S41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('T41');
worksheet.getCell('T41').value = wellsfargo.demoHeadingColumn13;
worksheet.getCell('T41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('T41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('T41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('U39:U41');
worksheet.getCell('U39:U41').value = wellsfargo.demoHeadingColumn14;
worksheet.getCell('U39:U41').font = {
    size: 12,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('U39:U41').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('U39:U41').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('A42');
worksheet.getCell('A42').value = "20.1";
worksheet.getCell('A42').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A42').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A42').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};

worksheet.mergeCells('A43');
worksheet.getCell('A43').value = "20.2";
worksheet.getCell('A43').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A43').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A43').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};

worksheet.mergeCells('A44');
worksheet.getCell('A44').value = "21.1";
worksheet.getCell('A44').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A44').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A44').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};

worksheet.mergeCells('A45');
worksheet.getCell('A45').value = "21.2";
worksheet.getCell('A45').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('A45').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A45').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};
//    ----------------------------SECOND Table-----------------

for (let i = 0; i < wellsfargo.demoColumns.length; i++) {
    // console.log("second table");
    let temp = i + 42;

    worksheet.mergeCells(temp, 2, temp, 7);
    worksheet.getCell(temp, 2, temp, 7).value = wellsfargo.demoColumns[i].coloumn1;
    worksheet.getCell(temp, 2, temp, 7).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell(temp, 2, temp, 7).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell(temp, 2, temp, 7).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('H' + temp);
    worksheet.getCell('H' + temp).value = wellsfargo.demoColumns[i].coloumn2;
    worksheet.getCell('H' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('H' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('H' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('I' + temp);
    worksheet.getCell('I' + temp).value = wellsfargo.demoColumns[i].coloumn3;
    worksheet.getCell('I' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('I' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('I' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('J' + temp);
    worksheet.getCell('J' + temp).value = wellsfargo.demoColumns[i].coloumn4;
    worksheet.getCell('J' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('J' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('J' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('K' + temp);
    worksheet.getCell('K' + temp).value = wellsfargo.demoColumns[i].coloumn5;
    worksheet.getCell('K' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('K' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('K' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('L' + temp);
    worksheet.getCell('L' + temp).value = wellsfargo.demoColumns[i].coloumn6;
    worksheet.getCell('L' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('L' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('L' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('M' + temp);
    worksheet.getCell('M' + temp).value = wellsfargo.demoColumns[i].coloumn7;
    worksheet.getCell('M' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('M' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('M' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + temp);
    worksheet.getCell('O' + temp).value = wellsfargo.demoColumns[i].coloumn8;
    worksheet.getCell('O' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('O' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('O' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('P' + temp);
    worksheet.getCell('P' + temp).value = wellsfargo.demoColumns[i].coloumn9;
    worksheet.getCell('P' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('P' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('P' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('Q' + temp);
    worksheet.getCell('Q' + temp).value = wellsfargo.demoColumns[i].coloumn10;
    worksheet.getCell('Q' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('Q' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('Q' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R' + temp);
    worksheet.getCell('R' + temp).value = wellsfargo.demoColumns[i].coloumn11;
    worksheet.getCell('R' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('R' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + temp);
    worksheet.getCell('S' + temp).value = wellsfargo.demoColumns[i].coloumn12;
    worksheet.getCell('S' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('S' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + temp);
    worksheet.getCell('T' + temp).value = wellsfargo.demoColumns[i].coloumn13;
    worksheet.getCell('T' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('T' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('U' + temp);
    worksheet.getCell('U' + temp).value = wellsfargo.demoColumns[i].coloumn14;
    worksheet.getCell('U' + temp).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('U' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('U' + temp).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

}


//    ----------------------------SECOND ROW-----------------


worksheet.mergeCells('B47:K47');
// worksheet.getCell('B47:K47').border = {
//     top: { style: 'thick' },
//     left: { style: 'thick' },
//     bottom: { style: 'none' },
//     right: { style: 'thick' }
// };

worksheet.mergeCells('L47');
worksheet.getCell('L47').value = wellsfargo.totalDemo;
worksheet.getCell('L47').font = {
    size: 12,
    bold: true
};
worksheet.getCell('L47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('L47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O47');
worksheet.getCell('O47').value = wellsfargo.totalDemoPeople;
worksheet.getCell('O47').font = {
    size: 12,
    bold: true
};
worksheet.getCell('O47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('O47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('P47');
worksheet.getCell('P47').value = wellsfargo.totalDemoHoursPerPerson;
worksheet.getCell('P47').font = {
    size: 12,
    bold: true
};
worksheet.getCell('P47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('P47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('Q47');
worksheet.getCell('Q47').value = wellsfargo.totalDemoHourlyBillRate;
worksheet.getCell('Q47').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('Q47').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('Q47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('Q47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R47');
worksheet.getCell('R47').value = wellsfargo.totalDemoUnionRate;
worksheet.getCell('R47').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('R47').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('R47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('S47');
worksheet.getCell('S47').value = wellsfargo.totalDemoLabor;
worksheet.getCell('S47').font = {
    size: 12,
    bold: true
};
worksheet.getCell('S47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('T47');
worksheet.getCell('T47').value = wellsfargo.totalDemoLaborTaxable;
worksheet.getCell('T47').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('T47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('U47');
worksheet.getCell('U47').value = wellsfargo.totalDemoProductAndLabor;
worksheet.getCell('U47').font = {
    size: 13,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('U47').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('U47').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};



worksheet.mergeCells('B48:U48');
worksheet.getCell('B48:U48').border = {
    top: { style: 'none' },
    left: { style: 'none' },
    bottom: { style: 'none' },
    right: { style: 'none' }
};


worksheet.mergeCells('B49:M49');
worksheet.getCell('B49:M49').value = wellsfargo.clarificationsHeading;
worksheet.getCell('B49:M49').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    color: { argb: 'FFFFFF' },
    bold: true
};
worksheet.getCell('B49:M49').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '808080' },
    bgColor: { argb: '808080' }
};
worksheet.getCell('B49:M49').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('B49:M49').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A50');
worksheet.getCell('A50').value = "30.1";
worksheet.getCell('A50').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A50').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};

worksheet.mergeCells('B50:M50');
worksheet.getCell('B50:M50').value = wellsfargo.clarificationDescription;
worksheet.getCell('B50:M50').font = {
    size: 16,
    name: 'Verdana',
    family: 1
};
worksheet.getCell('B50:M50').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('B50:M50').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};




worksheet.mergeCells('B51:U51');
worksheet.getCell('B51:U51').value = wellsfargo.taxHeading;
worksheet.getCell('B51:U51').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('B51:U51').alignment = { vertical: 'bottom', horizontal: 'right', wrapText: true };



worksheet.mergeCells('B52:K52');

worksheet.mergeCells('L52:M52');
worksheet.getCell('L52:M52').value = wellsfargo.taxHeadingColoumn1;
worksheet.getCell('L52:M52').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L52:M52').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('L52:M52').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O52:Q52');

worksheet.mergeCells('R52');
worksheet.getCell('R52').value = wellsfargo.taxHeadingColoumn2;
worksheet.getCell('R52').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('R52').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('R52').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('S52');
worksheet.getCell('S52').value = wellsfargo.taxHeadingColoumn3;
worksheet.getCell('S52').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('S52').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('S52').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('T52:U52');
worksheet.getCell('T52:U52').value = wellsfargo.totalTaxHeading;
worksheet.getCell('T52:U52').font = {
    size: 16,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('T52:U52').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('T52:U52').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};


worksheet.mergeCells('L53:M53');
worksheet.getCell('L53:M53').value = wellsfargo.taxRate;
worksheet.getCell('L53:M53').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('L53:M53').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('L53:M53').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('L53:M53').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('O53:Q53');
worksheet.getCell('O53:Q53').value = wellsfargo.product;
worksheet.getCell('O53:Q53').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O53:Q53').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O53:Q53').border = {
    top: { style: 'thick' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R53');
worksheet.getCell('R53').value = wellsfargo.productPreTax;
worksheet.getCell('R53').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R53').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R53').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S53');
worksheet.getCell('S53').value = wellsfargo.productTax;
worksheet.getCell('S53').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S53').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S53').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S53').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T53:U53');
worksheet.getCell('T53:U53').value = wellsfargo.productTotal;
worksheet.getCell('T53:U53').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T53:U53').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T53:U53').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A54');
worksheet.getCell('A54').value = "40";
worksheet.getCell('A54').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A54').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};

worksheet.mergeCells('O54:Q54');
worksheet.getCell('O54:Q54').value = wellsfargo.labor;
worksheet.getCell('O54:Q54').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O54:Q54').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O54:Q54').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R54');
worksheet.getCell('R54').value = wellsfargo.laborPreTax;
worksheet.getCell('R54').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R54').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R54').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S54');
worksheet.getCell('S54').value = wellsfargo.laborTax;
worksheet.getCell('S54').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S54').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S54').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S54').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T54:U54');
worksheet.getCell('T54:U54').value = wellsfargo.laborTotal;
worksheet.getCell('T54:U54').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T54:U54').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T54:U54').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};



worksheet.mergeCells('A55');
worksheet.getCell('A55').value = "40";
worksheet.getCell('A55').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A55').font = {
    size: 12,
    bold: true
};

worksheet.mergeCells('O55:Q55');
worksheet.getCell('O55:Q55').value = wellsfargo.subTotal1;
worksheet.getCell('O55:Q55').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('O55:Q55').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O55:Q55').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R55');
worksheet.getCell('R55').value = wellsfargo.subTotal1PreTax;
worksheet.getCell('R55').font = {
    size: 14,
    bold: true
};
worksheet.getCell('R55').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R55').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S55');
worksheet.getCell('S55').value = wellsfargo.subTotal1Tax;
worksheet.getCell('S55').font = {
    size: 14,
    bold: true
};
worksheet.getCell('S55').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S55').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T55:U55');
worksheet.getCell('T55:U55').value = wellsfargo.subTotal1Total;
worksheet.getCell('T55:U55').font = {
    size: 14,
    bold: true
};
worksheet.getCell('T55:U55').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T55:U55').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A56');
worksheet.getCell('A56').value = "40";
worksheet.getCell('A56').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A56').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};

worksheet.mergeCells('O56:Q56');
worksheet.getCell('O56:Q56').value = wellsfargo.demoProduct;
worksheet.getCell('O56:Q56').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O56:Q56').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O56:Q56').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R56');
worksheet.getCell('R56').value = wellsfargo.demoProductPreTax;
worksheet.getCell('R56').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R56').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R56').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S56');
worksheet.getCell('S56').value = wellsfargo.demoProductTax;
worksheet.getCell('S56').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S56').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S56').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S56').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T56:U56');
worksheet.getCell('T56:U56').value = wellsfargo.demoProductTotal;
worksheet.getCell('T56:U56').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T56:U56').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T56:U56').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A57');
worksheet.getCell('A57').value = "41";
worksheet.getCell('A57').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A57').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};

worksheet.mergeCells('O57:Q57');
worksheet.getCell('O57:Q57').value = wellsfargo.demoLabor;
worksheet.getCell('O57:Q57').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O57:Q57').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O57:Q57').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R57');
worksheet.getCell('R57').value = wellsfargo.demoLaborPreTax;
worksheet.getCell('R57').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R57').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R57').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S57');
worksheet.getCell('S57').value = wellsfargo.demoLaborTax;
worksheet.getCell('S57').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S57').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S57').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S57').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T57:U57');
worksheet.getCell('T57:U57').value = wellsfargo.demoLaborTotal;
worksheet.getCell('T57:U57').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T57:U57').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T57:U57').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('A58');
worksheet.getCell('A58').value = "41";
worksheet.getCell('A58').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
worksheet.getCell('A58').font = {
    size: 11,
    name: 'Verdana',
    family: 1
};
// ------------------------------------------.....to be continued----------------------------
worksheet.mergeCells('O58:Q58');
worksheet.getCell('O58:Q58').value = wellsfargo.subTotal2;
worksheet.getCell('O58:Q58').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('O58:Q58').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O58:Q58').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R58');
worksheet.getCell('R58').value = wellsfargo.subTotal2PreTax;
worksheet.getCell('R58').font = {
    size: 14,
    bold: true
};
worksheet.getCell('R58').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R58').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S58');
worksheet.getCell('S58').value = wellsfargo.subTotal2Tax;
worksheet.getCell('S58').font = {
    size: 14,
    bold: true
};
worksheet.getCell('S58').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S58').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T58:U58');
worksheet.getCell('T58:U58').value = wellsfargo.subTotal2Total;
worksheet.getCell('T58:U58').font = {
    size: 14,
    bold: true
};
worksheet.getCell('T58:U58').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T58:U58').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('O59:Q59');
worksheet.getCell('O59:Q59').value = wellsfargo.freight;
worksheet.getCell('O59:Q59').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O59:Q59').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O59:Q59').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R59');
worksheet.getCell('R59').value = wellsfargo.freightPreTax;
worksheet.getCell('R59').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R59').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('R59').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R59').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S59');
worksheet.getCell('S59').value = wellsfargo.freightTax;
worksheet.getCell('S59').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S59').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S59').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S59').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T59:U59');
worksheet.getCell('T59:U59').value = wellsfargo.freightTotal;
worksheet.getCell('T59:U59').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T59:U59').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T59:U59').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('O60:Q60');
worksheet.getCell('O60:Q60').value = wellsfargo.shipping;
worksheet.getCell('O60:Q60').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O60:Q60').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O60:Q60').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R60');
worksheet.getCell('R60').value = wellsfargo.shippingPreTax;
worksheet.getCell('R60').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R60').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('R60').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R60').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S60');
worksheet.getCell('S60').value = wellsfargo.shippingTax;
worksheet.getCell('S60').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S60').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S60').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S60').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T60:U60');
worksheet.getCell('T60:U60').value = wellsfargo.shippingTotal;
worksheet.getCell('T60:U60').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T60:U60').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T60:U60').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('O61:Q61');
worksheet.getCell('O61:Q61').value = wellsfargo.profit;
worksheet.getCell('O61:Q61').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O61:Q61').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O61:Q61').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R61');
worksheet.getCell('R61').value = wellsfargo.profitPreTax;
worksheet.getCell('R61').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R61').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('R61').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R61').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S61');
worksheet.getCell('S61').value = wellsfargo.profitTax;
worksheet.getCell('S61').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S61').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S61').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S61').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T61:U61');
worksheet.getCell('T61:U61').value = wellsfargo.profitTotal;
worksheet.getCell('T61:U61').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T61:U61').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T61:U61').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};


worksheet.mergeCells('O62:Q62');
worksheet.getCell('O62:Q62').value = wellsfargo.insurance;
worksheet.getCell('O62:Q62').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
};
worksheet.getCell('O62:Q62').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O62:Q62').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R62');
worksheet.getCell('R62').value = wellsfargo.insurancePreTax;
worksheet.getCell('R62').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('R62').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('R62').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R62').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S62');
worksheet.getCell('S62').value = wellsfargo.insuranceTax;
worksheet.getCell('S62').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2DCDB' },
    bgColor: { argb: 'F2DCDB' }
};
worksheet.getCell('S62').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('S62').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S62').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T62:U62');
worksheet.getCell('T62:U62').value = wellsfargo.insuranceTotal;
worksheet.getCell('T62:U62').font = {
    size: 14,
    // bold: true
};
worksheet.getCell('T62:U62').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T62:U62').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thick' }
};



worksheet.mergeCells('A63');
worksheet.getCell('A63').value = "41";
worksheet.getCell('A63').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
worksheet.getCell('A63').font = {
    size: 12,
    bold: true
};

worksheet.mergeCells('O63:Q63');
worksheet.getCell('O63:Q63').value = wellsfargo.allTotal;
worksheet.getCell('O63:Q63').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('O63:Q63').alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
worksheet.getCell('O63:Q63').border = {
    top: { style: 'thin' },
    left: { style: 'thick' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('R63');
worksheet.getCell('R63').value = wellsfargo.allTotalTaxPreTax;
worksheet.getCell('R63').font = {
    size: 14,
    bold: true
};
worksheet.getCell('R63').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('R63').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('S63');
worksheet.getCell('S63').value = wellsfargo.allTotalTax;
worksheet.getCell('S63').font = {
    size: 14,
    bold: true
};
worksheet.getCell('S63').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('S63').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thin' }
};

worksheet.mergeCells('T63:U63');
worksheet.getCell('T63:U63').value = wellsfargo.allMaterialTotal;
worksheet.getCell('T63:U63').font = {
    size: 14,
    bold: true
};
worksheet.getCell('T63:U63').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
worksheet.getCell('T63:U63').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thick' },
    right: { style: 'thick' }
};

worksheet.mergeCells('M59:N62');
worksheet.getCell('M59:N62').value = wellsfargo.rotationText;
worksheet.getCell('M59:N62').font = {
    size: 14,
    name: 'Verdana',
    family: 1,
    bold: true
};
worksheet.getCell('M59:N62').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F2F2F2' },
    bgColor: { argb: 'F2F2F2' }
};
worksheet.getCell('M59:N62').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, textRotation: 90 };
// worksheet.getCell('M59:M62').alignment = { textRotation: 30 };
// worksheet.getCell('M59:M62').alignment = { textRotation: -45 };
// worksheet.getCell('M59:M62').alignment = { textRotation: 'vertical' };


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



