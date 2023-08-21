const fs = require("fs");
const xlsx = require('xlsx');

/**
 * function to get the file downloaded name
 * based on the name of the report
 * @param {string} reportName name of the report
 * @returns {string}
 * **/
function getExportedFileName(reportName) {
    return reportName.replace(/ /g, '_');
}

/**
 * function to create a buffer of an excel file,
 * then return the Object with all the sheets content
 * @param {string} reportName name of the report
 * @returns {Object}
 * **/
exports.getFileContent = (reportName) => {
    const workbook = xlsx.readFile(reportName);
    return workbook
}

/**
 * function to return the Object of one excel sheet
 * @param {Object} file already parsed excel file Object
 * @param {string} sheetName name of the sheet to be returned
 * @returns {Object}
 * **/
exports.getSheetContent = (file, sheetName) => {
    return file.Sheets[sheetName];
}

/**
 * function to return an Array<string> of all the values inside a column
 * of an excel sheet
 * @param {Object} sheet already parsed excel sheet Object
 * @param {string} columnLetter column letter to be parsed
 * @returns {Array}
 * **/
exports.getColumnValues = (sheet, columnLetter) => {
    const filteredColumn = Object.entries(sheet).filter(column => column[0].includes(columnLetter));
    return filteredColumn.map(value => {
        const columnRawStringValue = value[1].r;
        return columnRawStringValue;
    });
}
