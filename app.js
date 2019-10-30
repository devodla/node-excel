const express = require('express');
const xl = require('excel4node');

const app = express();

app.get('/', (req, res) => {
  try {
    const data = [
      {   "FirstName": "John", 
          "LastName": "Parker", 
          "Age": "23", 
          "Cat": "23g",
          "SOP": "Active"
      },
      {   "FirstName": "Rose", 
          "LastName": "Jackson", 
          "Age": "44", 
          "Cat": "44g",
          "SOP": "InActive"
      }
    ];

    const wb = new xl.Workbook();

    // Add Worksheets to the workbook
    const ws = wb.addWorksheet('Report');

    ws.cell(1, 1).string('FirstName');
    ws.cell(1, 2).string('LastName');
    ws.cell(1, 3).string('Age');
    ws.cell(1, 4).string('Cat');
    ws.cell(1, 5).string('SOP');

    for (let i = 0; i < data.length; i += 1) {
      ws.cell(i + 2, 1).string(data[i].FirstName);
      ws.cell(i + 2, 2).string(data[i].LastName);
      ws.cell(i + 2, 3).string(data[i].Age);
      ws.cell(i + 2, 4).string(data[i].Cat);
      ws.cell(i + 2, 5).string(data[i].SOP);
    }

    const fileName = `Report_${Date.now().toString()}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
    wb.write(fileName, res);

  } catch (err) {
      console.error(res, err);
  }
});

app.listen(3000, () => console.log('Example app listening on port 3000!'))