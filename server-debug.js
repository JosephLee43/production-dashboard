const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const chokidar = require('chokidar');
const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

const FILE_PATH = 'C:\\Users\\jleemil\\Azenta, Inc\\Instrument Team - Ziath Build Plan\\Weekly Build Plan for Instruments Cell Wotton.xlsx';
const PORT = 3000;
const LOG_FILE = 'server-debug.log';

function log(message) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] ${message}\n`;
    fs.appendFileSync(LOG_FILE, logMessage);
    console.log(message);
}

app.use(express.static('public'));

async function parseExcel() {
    const workbook = new ExcelJS.Workbook();
    try {
        log('Attempting to read file: ' + FILE_PATH);
        await workbook.xlsx.readFile(FILE_PATH);
        log('File read successfully');
        
        const currentMonthName = dayjs().format('MMMM-YYYY');
        log('Looking for sheet: ' + currentMonthName);
        log('Available sheets: ' + workbook.worksheets.map(ws => ws.name).join(', '));
        
        const sheet = workbook.getWorksheet(currentMonthName) || workbook.worksheets[0];
        log('Using sheet: ' + (sheet ? sheet.name : 'NONE FOUND'));
        
        let data = [];
        const headers = [];
        
        sheet.getRow(1).eachCell((cell) => {
            headers.push(cell.value ? cell.value.toString().trim() : null);
        });

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const rowData = {};
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const header = headers[colNumber - 1];
                if (header) {
                    let val = cell.value;
                    
                    if (val instanceof Date || (val && val.constructor.name === 'Date')) {
                        val = dayjs(val).format('YYYY-MM-DD');
                    } else if (val && typeof val === 'object' && val.result) {
                        val = val.result instanceof Date ? dayjs(val.result).format('YYYY-MM-DD') : val.result;
                    }
                    rowData[header] = val;
                }
            });

            if (rowData['Item'] || rowData['Work Order']) {
                data.push(rowData);
            }
        });

        log('Parsed rows: ' + data.length);
        return {
            month: currentMonthName,
            rows: data,
            serverToday: dayjs().format('YYYY-MM-DD'),
            lastUpdated: dayjs().format('HH:mm:ss')
        };
    } catch (err) {
        log("Dashboard Error: " + err.message);
        log("Full error: " + JSON.stringify(err, null, 2));
        return null;
    }
}

const watcher = chokidar.watch(FILE_PATH, {
    persistent: true,
    ignoreInitial: true
});

watcher.on('change', async () => {
    log('File update detected. Syncing...');
    const d = await parseExcel();
    if (d) io.emit('update', d);
});

watcher.on('error', error => {
    log('Watcher error: ' + error);
});

io.on('connection', async (socket) => {
    log('Client connected');
    const d = await parseExcel();
    if (d) {
        log('Sending data to client - rows: ' + d.rows.length);
        socket.emit('update', d);
    } else {
        log('No data to send');
    }
});

server.listen(PORT, () => {
    log(`Purple Dashboard Live: http://localhost:${PORT}`);
    log(`Watching: ${FILE_PATH}`);
});

process.on('uncaughtException', (err) => {
    log('Uncaught Exception: ' + err);
});

process.on('unhandledRejection', (reason, promise) => {
    log('Unhandled Rejection: ' + reason);
});
