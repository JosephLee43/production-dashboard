const ExcelJS = require('exceljs');
const dayjs = require('dayjs');

const FILE_PATH = 'C:\\Users\\jleemil\\Azenta, Inc\\Instrument Team - Ziath Build Plan\\Weekly Build Plan for Instruments Cell Wotton.xlsx';

async function test() {
    const workbook = new ExcelJS.Workbook();
    try {
        console.log('Reading file...');
        await workbook.xlsx.readFile(FILE_PATH);
        console.log('Success! Worksheets:', workbook.worksheets.map(ws => ws.name));
        
        const currentMonthName = dayjs().format('MMMM-YYYY');
        console.log('Current month format:', currentMonthName);
        
        const sheet = workbook.getWorksheet(currentMonthName) || workbook.worksheets[0];
        console.log('Using sheet:', sheet.name);
        console.log('Row count:', sheet.rowCount);
    } catch (err) {
        console.error('Error:', err);
    }
}

test();
