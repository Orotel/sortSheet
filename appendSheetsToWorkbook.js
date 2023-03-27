
function appendSheetsToWorkbook(workbook, sheets) {
    sheets.forEach(({ sheetName, data }) => {
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, ws, sheetName);
    });
    return workbook;
}

module.exports = appendSheetsToWorkbook;