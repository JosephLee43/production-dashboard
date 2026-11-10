const dayjs = require('dayjs');

function normalizeCellValue(value) {
    if (value === null || value === undefined) {
        return '';
    }

    if (typeof value === 'string') {
        return value.replace(/\u00a0/g, ' ').trim();
    }

    if (typeof value === 'number' || typeof value === 'boolean') {
        return value;
    }

    if (value instanceof Date || (value && value.constructor && value.constructor.name === 'Date')) {
        return dayjs(value).isValid() ? dayjs(value).format('YYYY-MM-DD') : '';
    }

    if (value && typeof value === 'object') {
        if (value.result !== undefined) {
            return normalizeCellValue(value.result);
        }

        if (Array.isArray(value.richText)) {
            return value.richText.map((part) => normalizeCellValue(part && part.text ? part.text : '')).join('');
        }

        if (value.text !== undefined) {
            return normalizeCellValue(value.text);
        }

        if (value.formula !== undefined && value.result === undefined) {
            return '';
        }

        if (value.toString && value.toString() !== '[object Object]') {
            return normalizeCellValue(value.toString());
        }

        return '';
    }

    return String(value).trim();
}

function hasMeaningfulRow(row) {
    if (!row || typeof row !== 'object') {
        return false;
    }

    const item = normalizeCellValue(row['Item'] ?? row['item']);
    const workOrder = normalizeCellValue(row['Work Order'] ?? row['WorkOrder'] ?? row['work order']);
    const workstation = normalizeCellValue(row['Workstation'] ?? row['Work Station'] ?? row['workstation']);
    const status = normalizeCellValue(row['Status'] ?? row['status']);

    return [item, workOrder, workstation, status].some((value) => value !== '');
}

module.exports = {
    normalizeCellValue,
    hasMeaningfulRow
};
