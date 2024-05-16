// Import Library
const ExcelJS = require('exceljs');

// Variables
const INPUT_FILE_PATH = './excel/score_2024-05-13 - 시트1.xlsx';
const OUTPUT_FILE_PATH = './excel/score_2024-05-13 - 시트1-updated.xlsx';
const EXCEL_SHEET_NUMBER = 1;
const TARGET_COLUMN_INDEX = 'G';
const NEW_COLUMN_INDEX = 'H';

/** Read Excel row of target column line by line and write a color column */
async function addColorColumn() {
  // 1. Read File
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(INPUT_FILE_PATH);

  // 2. Read Sheet
  const worksheet = workbook.getWorksheet(EXCEL_SHEET_NUMBER);

  // 3. Read Row line by line
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const cell = row.getCell(TARGET_COLUMN_INDEX);
    const fgColorIndexed = cell?.fill?.fgColor?.indexed;

    // Log
    // if (cell.row === 15) console.log(fgColorIndexed);

    // Daily Evaluation
    // if (fgColorIndexed === 16) row.getCell(NEW_COLUMN_INDEX).value = 'RED'
    // if (fgColorIndexed === 17) row.getCell(NEW_COLUMN_INDEX).value = 'ORANGE'
    // if (fgColorIndexed === 19) row.getCell(NEW_COLUMN_INDEX).value = 'YELLOW'
    // if (fgColorIndexed === 18) row.getCell(NEW_COLUMN_INDEX).value = 'GREEN'

    // Score Sheet
    if (fgColorIndexed === 18) row.getCell(NEW_COLUMN_INDEX).value = 'RED'
    if (fgColorIndexed === 17) row.getCell(NEW_COLUMN_INDEX).value = 'ORANGE'
    if (fgColorIndexed === 15) row.getCell(NEW_COLUMN_INDEX).value = 'YELLOW'
    if (fgColorIndexed === 16) row.getCell(NEW_COLUMN_INDEX).value = 'GREEN'
  });

  await workbook.xlsx.writeFile(OUTPUT_FILE_PATH);
}

addColorColumn();
