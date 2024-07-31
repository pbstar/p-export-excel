import { Config, Sheet } from './types'
export default function create(e: Config) {
  let base64 = ''
  let xlsxFile = ''
  let xlsxtables: string[] = [];
  let csvFile = ''

  if (e.fileType == 'xlsx') {
    xlsxtables = createTable(e)
    xlsxFile = tablesToExcel(xlsxtables, e)
    base64 = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(xlsxFile);
  } else if (e.fileType == 'csv') {
    if (e.sheets.length > 0) {
      if (e.sheets.length > 1) {
        console.warn('csv格式下多表导出仅会导出第一个表')
      }
      csvFile = createCsv(e.sheets[0])
    }
    base64 = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvFile);
  }
  return {
    base64: base64,
    csv: csvFile,
    xlsx: xlsxtables
  }
}
function createTable(e: Config) {
  let xlsxtables = []
  for (let i = 0; i < e.sheets.length; i++) {
    let sheet = e.sheets[i];
    let sheetStyle = ''
    if (e.sheetStyle) sheetStyle += e.sheetStyle
    if (sheet.style) sheetStyle += sheet.style
    let xlsxTable = "<table style='border-spacing: 0;font-family:宋体;" + sheetStyle + ";vnd.ms-excel.numberformat:@'>";
    for (let j = 0; j < sheet.rows.length; j++) {
      let row = sheet.rows[j];
      let rowStyle = ''
      if (e.rowStyle) rowStyle += e.rowStyle
      if (sheet.rowStyle) rowStyle += sheet.rowStyle
      if (row.style) rowStyle += row.style
      xlsxTable += "<tr style='" + rowStyle + "'>";
      for (let k = 0; k < row.cells.length; k++) {
        let cell = row.cells[k];
        let cellStyle = ''
        if (e.cellStyle) cellStyle += e.cellStyle
        if (sheet.cellStyle) cellStyle += sheet.cellStyle
        if (row.cellStyle) cellStyle += row.cellStyle
        if (cell instanceof Object) {
          xlsxTable += "<td";
          if (cell.colspan) xlsxTable += " colspan=" + cell.colspan;
          if (cell.rowspan) xlsxTable += " rowspan=" + cell.rowspan;
          if (cell.style) cellStyle += cell.style
          xlsxTable += " style='" + cellStyle + "'>" + cell.text + "</td>";
        } else {
          xlsxTable += "<td style='" + cellStyle + "'>" + cell + "</td>";
        }
      }
      xlsxTable += "</tr>";
    }
    xlsxTable += "</table>";
    xlsxtables.push(xlsxTable);
  }
  return xlsxtables
}
function tablesToExcel(xlsxtables: any, e: Config) {
  const html_start = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>`
  const template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>`
  const template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`
  const template_HTMLWorksheet = `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
  <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
  <link id="Main-File" rel="Main-File" href="../WorkBook.htm">
  <link rel="File-List" href="filelist.xml">
</head>
<body>{SheetContent}</body>
</html>`

  const template_WorkBook = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook>
  <x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
  <x:ActiveSheet>0</x:ActiveSheet>
</x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
  <frame src="sheet0.htm" name="frSheet">
  <noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
  <o:MainFile HRef="../WorkBook.htm"/>
  {ListWorksheets}
  <o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`

  const format = (s: any, c: any) => {
    return s.replace(/{(\w+)}/g, function (m: any, p: any) {
      return c[p]
    })
  }

  const context_WorkBook = {
    ExcelWorksheets: '',
    HTMLWorksheets: '',
    ListWorksheets: ''
  }

  e.sheets.forEach((p, SheetIndex) => {
    const SheetName = p.sheetName || 'Sheet' + SheetIndex
    context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
      SheetIndex: SheetIndex,
      SheetName: SheetName
    })
    let content = xlsxtables[SheetIndex]
    context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
      SheetIndex: SheetIndex,
      SheetContent: content
    })
    context_WorkBook.ListWorksheets += format(template_ListWorksheet, {
      SheetIndex: SheetIndex
    })
  })
  return format(template_WorkBook, context_WorkBook)
}
function createCsv(sheet: Sheet) {
  let str = ''
  for (let j = 0; j < sheet.rows.length; j++) {
    let row = sheet.rows[j];
    if (j > 0) str += " \n ";
    for (let k = 0; k < row.cells.length; k++) {
      let cell = row.cells[k];
      if (k > 0) str += ",";
      if (cell instanceof Object) {
        str += cell.text;
      } else {
        str += cell;
      }
    }
  }
  return str;
}