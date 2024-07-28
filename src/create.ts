import { Config } from './types'
export default function create(e: Config) {
  let excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel';
  excelFile += '; charset=UTF-8">';
  excelFile += "<head>";
  excelFile += "<!--[if gte mso 9]>";
  excelFile += "<xml>";
  excelFile += "<x:ExcelWorkbook>";
  excelFile += "<x:ExcelWorksheets>";
  for (let i = 0; i < e.sheets.length; i++) {
    let sheet = e.sheets[i];
    excelFile += "<x:ExcelWorksheet>";
    excelFile += "<x:Name>";
    excelFile += sheet.sheetName || 'sheet' + (i + 1);
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
  for (let i = 0; i < e.sheets.length; i++) {
    let sheet = e.sheets[i];
    let sheetStyle = e.sheetStyle + sheet.style
    excelFile += "<table style='font-family:宋体;" + sheetStyle + ";vnd.ms-excel.numberformat:@'>";
    for (let j = 0; j < sheet.rows.length; j++) {
      let row = sheet.rows[j];
      let rowStyle = e.rowStyle + row.style
      excelFile += "<tr style='" + rowStyle + "'>";
      for (let k = 0; k < row.cells.length; k++) {
        let cell = row.cells[k];
        if (cell instanceof Object) {
          let cellStyle = e.cellStyle + cell.style
          excelFile += "<td colspan=" + cell.colSpan + " rowspan=" + cell.rowSpan + " style='" + cellStyle + "'>" + cell.text + "</td>";
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
  console.log(excelFile);
  
  return 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelFile);
}