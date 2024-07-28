import { Sheet } from './types'
export default function create(sheets: Sheet[]) {
  let excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel';
  excelFile += '; charset=UTF-8">';
  excelFile += "<head>";
  excelFile += "<!--[if gte mso 9]>";
  excelFile += "<xml>";
  excelFile += "<x:ExcelWorkbook>";
  excelFile += "<x:ExcelWorksheets>";
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    excelFile += "<x:ExcelWorksheet>";
    excelFile += "<x:Name>";
    excelFile += sheet.name || 'sheet' + (i + 1);
    excelFile += "</x:Name>";
    excelFile += "<x:WorksheetOptions>";
    excelFile += "<x:DisplayGridlines/>";
    excelFile += "</x:WorksheetOptions>";
    excelFile += "</x:ExcelWorksheet>";
  }
  excelFile += "</x:ExcelWorksheets>";
  excelFile += "</x:ExcelWorkbook>";
  excelFile += "</xml>";
  excelFile += "<![endif]-->";
  excelFile += "</head>";
  excelFile += "<body>";
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    excelFile += "<table style='font-family:SimSun;vnd.ms-excel.numberformat:@'>";
    for (let j = 0; j < sheet.table.rows.length; j++) {
      let row = sheet.table.rows[j];
      excelFile += "<tr align='center'>";
      for (let k = 0; k < row.cells.length; k++) {
        let cell = row.cells[k];
        if (cell instanceof Object) {
          excelFile += "<td colspan=" + cell.colSpan + " rowspan=" + cell.rowSpan + " style=" + cell.style + " align=" + cell.align + ">" + cell.text + "</td>";
        } else {
          excelFile += "<td>" + cell + "</td>";
        }
      }
      excelFile += "</tr>";
    }
    excelFile += "</table>";
  }
  excelFile += "</body>";
  excelFile += "</html>";
  return 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelFile);
}