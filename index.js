// Require library
var xl = require('excel4node');
const fs = require('fs');
const sharp = require('sharp');
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
const path = require('path');

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');

// Create a reusable style
var style = wb.createStyle({
    font: {
        color: '#FF0800',
        size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -',
});

// Set value of cell A1 to 100 as a number type styled with paramaters of style
// ws.cell(1, 1)
//   .number(100)
//   .style(style);


const imagesfolder = './images';


// fs.readdirSync(imagesolder).forEach((file, index) => {
//     console.log(file)
//     sharp(file).resize({
//         height: 780
//     }).toFile(__dirname + "/output.jpg")
// });





fs.readdirSync(imagesfolder).forEach((file, index) => {
    let width = 2;
    let height = 14;
    let step;
    switch (index) {
        case 0:
            step = 1
            break;
        case 1:
            step = width + 2
            break;
        default:
            step = index * width + index + 1
            break;
    }

    console.log(index , step);


    ws.addImage({
        path: `images/${file}`,
        type: 'picture',
        position: {
            type: 'twoCellAnchor',
            from: {
                col: step,
                colOff: 0,
                row: 1,
                rowOff: 0,
            },
            to: {
                col: step + width,
                colOff: 0,
                row: height,
                rowOff: 0,
            },
        },
    });

});




wb.write('Excel.xlsx');