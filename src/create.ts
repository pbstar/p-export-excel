import { Config } from './types'
export function createXlsx(e: Config) {
  let xlsxFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
  xlsxFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
  xlsxFile += "<head>";
  xlsxFile += "<!--[if gte mso 9]>";
  xlsxFile += "<xml>";
  xlsxFile += "<x:ExcelWorkbook>";
  xlsxFile += "<x:ExcelWorksheets>";
  for (let i = 0; i < e.sheets.length; i++) {
    let sheet = e.sheets[i];
    xlsxFile += "<x:ExcelWorksheet>";
    xlsxFile += "<x:Name>";
    xlsxFile += sheet.sheetName || 'sheet' + (i + 1);
    xlsxFile += "</x:Name>";
    xlsxFile += "<x:WorksheetOptions>";
    xlsxFile += "<x:DisplayGridlines/>";
    xlsxFile += "</x:WorksheetOptions>";
    xlsxFile += "</x:ExcelWorksheet>";
  }
  xlsxFile += "</x:ExcelWorksheets>";
  xlsxFile += "</x:ExcelWorkbook>";
  xlsxFile += "</xml>";
  xlsxFile += "<![endif]-->";
  xlsxFile += "</head>";
  xlsxFile += "<body>";
  for (let i = 0; i < e.sheets.length; i++) {
    let sheet = e.sheets[i];
    let sheetStyle = ''
    if (e.sheetStyle) sheetStyle += e.sheetStyle
    if (sheet.style) sheetStyle += sheet.style
    xlsxFile += "<table style='font-family:宋体;" + sheetStyle + ";vnd.ms-excel.numberformat:@'>";
    for (let j = 0; j < sheet.rows.length; j++) {
      let row = sheet.rows[j];
      let rowStyle = ''
      if (e.rowStyle) rowStyle += e.rowStyle
      if (sheet.rowStyle) rowStyle += sheet.rowStyle
      if (row.style) rowStyle += row.style
      xlsxFile += "<tr style='" + rowStyle + "'>";
      for (let k = 0; k < row.cells.length; k++) {
        let cell = row.cells[k];
        let cellStyle = ''
        if (e.cellStyle) cellStyle += e.cellStyle
        if (sheet.cellStyle) cellStyle += sheet.cellStyle
        if (row.cellStyle) cellStyle += row.cellStyle
        if (cell instanceof Object) {
          xlsxFile += "<td";
          if (cell.colspan) xlsxFile += " colspan=" + cell.colspan;
          if (cell.rowspan) xlsxFile += " rowspan=" + cell.rowspan;
          if (cell.style) cellStyle += cell.style
          xlsxFile += " style='" + cellStyle + "'>" + cell.text + "</td>";
        } else {
          xlsxFile += "<td style='" + cellStyle + "'>" + cell + "</td>";
        }
      }
      xlsxFile += "</tr>";
    }
    xlsxFile += "</table>";
  }
  xlsxFile += "</body>";
  xlsxFile += "</html>";
  return 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(xlsxFile);
}
export function createCsv(e: Config) {
  let csvFile = '';
  if (e.sheets.length > 0) {
    if (e.sheets.length > 1) {
      console.warn('csv格式下多表导出仅会导出第一个表')
    }
    let sheet = e.sheets[0];
    for (let j = 0; j < sheet.rows.length; j++) {
      let row = sheet.rows[j];
      if (j > 0) csvFile += "\n";
      for (let k = 0; k < row.cells.length; k++) {
        let cell = row.cells[k];
        if (k > 0) csvFile += ",";
        if (cell instanceof Object) {
          csvFile += cell.text;
        } else {
          csvFile += cell;
        }
      }
    }
  }
  return 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvFile);
}