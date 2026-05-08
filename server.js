const express = require('express');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { Server } = require('socket.io');
const chokidar = require('chokidar');
const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');

dayjs.extend(customParseFormat);

const app = express();
const server = http.createServer(app);
const io = new Server(server);

// --- CONFIGURATION ---
// IMPORTANT: Update this to your actual OneDrive path
const FILE_PATH = 'C:\\Users\\jleemil\\Azenta, Inc\\Instrument Team - Ziath Build Plan\\2026 Weekly Build Plan - Instruments Cell.xlsx';
const PORT = 3000;
const FILE_DIR = path.dirname(FILE_PATH);
const FILE_NAME = path.basename(FILE_PATH).toLowerCase();

// Cache the last parsed data to avoid re-reading on every connection
let cachedData = null;
let isReading = false;
let refreshInProgress = false;
let refreshQueued = false;
let lastKnownFileSignature = null;

function resolveWorksheet(workbook) {
    const currentMonthName = dayjs().format('MMMM-YYYY');
    const directMatch = workbook.getWorksheet(currentMonthName);

    if (directMatch) {
        return {
            monthLabel: currentMonthName,
            worksheet: directMatch,
            resolution: 'exact-current-month'
        };
    }

    const monthNamedSheets = workbook.worksheets
        .map((worksheet) => {
            const parsed = dayjs(worksheet.name, 'MMMM-YYYY', true);
            if (!parsed.isValid()) {
                return null;
            }

            return {
                worksheet,
                parsedMonth: parsed
            };
        })
        .filter(Boolean)
        .sort((left, right) => left.parsedMonth.valueOf() - right.parsedMonth.valueOf());

    if (monthNamedSheets.length > 0) {
        const latestMonthSheet = monthNamedSheets[monthNamedSheets.length - 1];
        return {
            monthLabel: latestMonthSheet.worksheet.name,
            worksheet: latestMonthSheet.worksheet,
            resolution: 'latest-month-sheet'
        };
    }

    return {
        monthLabel: workbook.worksheets[0]?.name || currentMonthName,
        worksheet: workbook.worksheets[0],
        resolution: 'first-worksheet-fallback'
    };
}

function normalizeRowDate(value) {
    if (!value) {
        return null;
    }

    if (value instanceof Date || (value && value.constructor?.name === 'Date')) {
        const parsed = dayjs(value);
        return parsed.isValid() ? parsed.format('YYYY-MM-DD') : null;
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        const parsed = dayjs(new Date(excelEpoch.getTime() + Math.round(value) * 86400000));
        return parsed.isValid() ? parsed.format('YYYY-MM-DD') : null;
    }

    const text = value.toString().trim();
    const isoMatch = text.match(/\d{4}-\d{2}-\d{2}/);
    if (isoMatch) {
        return isoMatch[0];
    }

    const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
        const month = slashMatch[1].padStart(2, '0');
        const day = slashMatch[2].padStart(2, '0');
        return `${slashMatch[3]}-${month}-${day}`;
    }

    const parsed = dayjs(text);
    return parsed.isValid() ? parsed.format('YYYY-MM-DD') : null;
}

function deriveServerToday(rows, resolution) {
    const realToday = dayjs().format('YYYY-MM-DD');

    if (resolution === 'exact-current-month') {
        return realToday;
    }

    const dateFields = [
        'Planned Finish Date',
        'Planned Start Date',
        'Build Date',
        'Date',
        'Finish Date',
        'Date completed'
    ];

    const candidateDates = rows
        .flatMap((row) => dateFields.map((field) => normalizeRowDate(row[field])))
        .filter(Boolean)
        .sort();

    if (candidateDates.length === 0) {
        return realToday;
    }

    const latestOnOrBeforeToday = [...candidateDates]
        .reverse()
        .find((candidateDate) => candidateDate <= realToday);

    return latestOnOrBeforeToday || candidateDates[candidateDates.length - 1];
}

app.use(express.static('public'));

async function parseExcel(forceRefresh = false) {
    if (!forceRefresh && cachedData) {
        console.log('Returning cached data');
        return cachedData;
    }
    
    if (isReading) {
        console.log('Already reading file, waiting...');
        // Wait for the current read to complete
        while (isReading) {
            await new Promise(resolve => setTimeout(resolve, 100));
        }
        return cachedData;
    }
    
    isReading = true;
    const workbook = new ExcelJS.Workbook();
    try {
        console.log('Attempting to read file:', FILE_PATH);
        await workbook.xlsx.readFile(FILE_PATH);
        console.log('File read successfully');
        
        // Match sheet name to current month (e.g., "December-2025")
        const currentMonthName = dayjs().format('MMMM-YYYY');
        console.log('Looking for sheet:', currentMonthName);
        console.log('Available sheets:', workbook.worksheets.map(ws => ws.name));
        const { worksheet: sheet, monthLabel, resolution } = resolveWorksheet(workbook);
        console.log('Using sheet:', sheet ? sheet.name : 'NONE FOUND', `(${resolution})`);
        
        if (!sheet) {
            throw new Error('No worksheet found in workbook');
        }
        
        let data = [];
        const headers = [];
        
        // Capture headers from row 1
        sheet.getRow(1).eachCell((cell) => {
            headers.push(cell.value ? cell.value.toString().trim() : null);
        });
        
        console.log('Excel Headers:', headers);

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const rowData = {};
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const header = headers[colNumber - 1];
                if (header) {
                    let val = cell.value;
                    
                    // Handle Excel Date Objects to prevent Timezone/0-result bugs
                    if (val instanceof Date || (val && val.constructor.name === 'Date')) {
                        val = dayjs(val).format('YYYY-MM-DD');
                    } else if (val && typeof val === 'object' && val.result) {
                        // Handle formula results
                        val = val.result instanceof Date ? dayjs(val.result).format('YYYY-MM-DD') : val.result;
                    } else if (val && typeof val === 'object') {
                        // Handle rich text format from Excel
                        if (val.richText && Array.isArray(val.richText)) {
                            val = val.richText.map(rt => rt.text).join('');
                        } else if (val.text !== undefined) {
                            val = val.text;
                        } else if (val.toString && val.toString() !== '[object Object]') {
                            val = val.toString();
                        } else {
                            val = '';
                        }
                    }
                    rowData[header] = val;
                }
            });

            // Filter out empty rows
            if (rowData['Item'] || rowData['Work Order']) {
                data.push(rowData);
            }
        });

        console.log('Parsed rows:', data.length);
        
        // Group by workstation and sort by priority
        const groupedData = {};
        data.forEach(row => {
            // Only include rows with actual item data and valid workstation
            if (!row['Item'] || row['Item'].toString().trim() === '') return;
            
            const workstation = row['Workstation'] || row['Work Station'] || '';
            if (!workstation || workstation.toString().trim() === '' || workstation.toString().trim() === 'null') return;
            
            if (!groupedData[workstation]) {
                groupedData[workstation] = [];
            }
            groupedData[workstation].push(row);
        });
        
        // Sort each workstation's items by priority
        Object.keys(groupedData).forEach(station => {
            groupedData[station].sort((a, b) => {
                const priorityA = parseInt(a['Priority'] || a['priority'] || a['Workstation Priority'] || 999);
                const priorityB = parseInt(b['Priority'] || b['priority'] || b['Workstation Priority'] || 999);
                return priorityA - priorityB;
            });
        });
        
        const effectiveServerToday = deriveServerToday(data, resolution);
        console.log('Using reference date:', effectiveServerToday);

        cachedData = {
            month: monthLabel,
            rows: data,
            groupedByStation: groupedData,
            serverToday: effectiveServerToday,
            lastUpdated: dayjs().format('HH:mm:ss')
        };
        isReading = false;
        return cachedData;
    } catch (err) {
        console.error("Dashboard Error:", err.message);
        console.error("Full error:", err);
        isReading = false;
        return null;
    }
}

async function requestDataRefresh(source = 'manual') {
    if (refreshInProgress) {
        refreshQueued = true;
        console.log(`Refresh already in progress; queueing another pass (${source})`);
        return;
    }

    refreshInProgress = true;
    try {
        do {
            refreshQueued = false;
            console.log(`Refreshing dashboard data (${source})...`);
            const data = await parseExcel(true);
            if (data) {
                io.emit('update', data);
            }
        } while (refreshQueued);
    } finally {
        refreshInProgress = false;
    }
}

function getFileSignature() {
    try {
        const stats = fs.statSync(FILE_PATH);
        return `${stats.size}:${stats.mtimeMs}:${stats.ctimeMs}`;
    } catch (error) {
        return null;
    }
}

function startFallbackFilePolling(intervalMs = 8000) {
    lastKnownFileSignature = getFileSignature();

    setInterval(async () => {
        const currentSignature = getFileSignature();

        if (!currentSignature || currentSignature === lastKnownFileSignature) {
            return;
        }

        lastKnownFileSignature = currentSignature;
        console.log('File signature changed via fallback poll. Syncing...');
        await requestDataRefresh('poll:signature');
    }, intervalMs);
}

function isTargetWorkbookPath(changedPath) {
    if (!changedPath) {
        return false;
    }

    try {
        return path.basename(changedPath).toLowerCase() === FILE_NAME;
    } catch (error) {
        return false;
    }
}

// Push updates when file is saved
const watcher = chokidar.watch(FILE_DIR, {
    persistent: true,
    ignoreInitial: true,
    depth: 0,
    usePolling: true,
    interval: 1000,
    binaryInterval: 1500,
    awaitWriteFinish: {
        stabilityThreshold: 1800,
        pollInterval: 150
    }
});

watcher.on('change', async (changedPath) => {
    if (!isTargetWorkbookPath(changedPath)) {
        return;
    }

    console.log('File update detected. Syncing...', changedPath);
    lastKnownFileSignature = getFileSignature();
    await requestDataRefresh('watch:change');
});

watcher.on('add', async (changedPath) => {
    if (!isTargetWorkbookPath(changedPath)) {
        return;
    }

    console.log('File add detected. Syncing...', changedPath);
    lastKnownFileSignature = getFileSignature();
    await requestDataRefresh('watch:add');
});

watcher.on('unlink', async (changedPath) => {
    if (!isTargetWorkbookPath(changedPath)) {
        return;
    }

    console.log('File unlink detected. Waiting for re-add...', changedPath);
    lastKnownFileSignature = null;
    await requestDataRefresh('watch:unlink');
});

watcher.on('error', error => {
    console.error('Watcher error:', error);
});

startFallbackFilePolling(3000);

// Check for day change every minute and refresh cache if needed
let lastCheckedDate = dayjs().format('YYYY-MM-DD');
setInterval(async () => {
    const currentDate = dayjs().format('YYYY-MM-DD');
    if (currentDate !== lastCheckedDate) {
        console.log('New day detected. Refreshing cache...');
        lastCheckedDate = currentDate;
        console.log('Broadcasting update for new day:', currentDate);
        await requestDataRefresh('day-rollover');
    }
}, 60000); // Check every minute

io.on('connection', async (socket) => {
    console.log('Client connected');
    
    // Send loading state immediately
    socket.emit('loading', { message: 'Reading Excel file from OneDrive...' });
    
    const d = await parseExcel();
    if (d) {
        console.log('Sending data to client');
        socket.emit('update', d);
    } else {
        console.log('No data to send');
        socket.emit('error', { message: 'Failed to read Excel file' });
    }
});

server.listen(PORT, '0.0.0.0', () => {
    console.log(`\x1b[35m%s\x1b[0m`, `Purple Dashboard Live: http://localhost:${PORT}`);
    console.log(`Network access: http://<your-local-ip>:${PORT}`);
    console.log(`Watching: ${FILE_PATH}`);
});

// Add error handlers
process.on('uncaughtException', (err) => {
    console.error('Uncaught Exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});