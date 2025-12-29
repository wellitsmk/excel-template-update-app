
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.static('public'));
app.use(require('cors')());


const upload = multer({ dest: 'uploads/' });

app.post('/upload', upload.single('template'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);

    
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    console.log('sheets', sheetNames);

    const summary = workbook.getWorksheet('Cover');
    summary.getCell('F3').value = 'Bechtel India';     
    summary.getCell('F4').value = 'mkumar37';
    summary.getCell('F5').value = new Date().toISOString();

    
    const dataSheet = workbook.getWorksheet('Data');
    
    const tableData = [
      { name: 'Amit', age: 29, salary: 80000 },
      { name: 'Raj', age: 31, salary: 90000 },
      { name: 'Arjun', age: 28, salary: 80000 }
    ];


    const startRow = 5;
    tableData.forEach((row, index) => {
      const excelRow = dataSheet.getRow(startRow + index);
      excelRow.getCell(1).value = row.name;
      excelRow.getCell(2).value = row.age;
      excelRow.getCell(3).value = row.salary;
      excelRow.commit();
    });

    
    const outputPath = path.join(
      __dirname,
      'uploads',
      `output-${Date.now()}.xlsx`
    );

    await workbook.xlsx.writeFile(outputPath);

    // fs.unlinkSync(req.file.path);

    res.download(outputPath, 'filled-template.xlsx', () => {
      fs.unlinkSync(outputPath);
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Excel processing failed', message: error.message });
  }
});

app.listen(3000, () => {
  console.log('Started: http://localhost:3000');
});

