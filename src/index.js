function pExportExcel(objInfo) {
  let obj = {
    fileName: objInfo.fileName || '文件名',
    footName: objInfo.footName || 'sheet1',
    theadList: objInfo.theadList || [],
    tbodyList: objInfo.tbodyList || []
  }
  let excel = '<table style="font-size:15px;font-family:Microsoft YaHei;vnd.ms-excel.numberformat:@">';
  for (let i = 0; i < obj.theadList.length; i++) {
    excel += "<tr align='center'>"
    let th = obj.theadList[i]
    for (let h = 0; h < th.length; h++) {
      excel += "<th colspan=" + th[h].colspan + " rowspan=" + th[h].rowspan + " style=" + th[h].style + " align=" + th[h].align + ">" + th[h].text + "</th>"
    }
    excel += "</tr>";
  }
  for (let i = 0; i < obj.tbodyList.length; i++) {
    excel += "<tr align='center'>"
    let td = obj.tbodyList[i]
    for (let d = 0; d < td.length; d++) {
      excel += "<td colspan=" + td[d].colspan + " rowspan=" + td[d].rowspan + " style=" + td[d].style + " align=" + td[d].align + ">" + td[d].text + "</td>"
    }
    excel += "</tr>";
  }
  excel += "</table>";
  let excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
  excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel';
  excelFile += '; charset=UTF-8">';
  excelFile += "<head>";
  excelFile += "<!--[if gte mso 9]>";
  excelFile += "<xml>";
  excelFile += "<x:ExcelWorkbook>";
  excelFile += "<x:ExcelWorksheets>";
  excelFile += "<x:ExcelWorksheet>";
  excelFile += "<x:Name>";
  excelFile += obj.footName;
  excelFile += "</x:Name>";
  excelFile += "<x:WorksheetOptions>";
  excelFile += "<x:DisplayGridlines/>";
  excelFile += "</x:WorksheetOptions>";
  excelFile += "</x:ExcelWorksheet>";
  excelFile += "</x:ExcelWorksheets>";
  excelFile += "</x:ExcelWorkbook>";
  excelFile += "</xml>";
  excelFile += "<![endif]-->";
  excelFile += "</head>";
  excelFile += "<body>";
  excelFile += excel;
  excelFile += "</body>";
  excelFile += "</html>";
  let url = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelFile);
  let link = document.createElement("a");
  link.href = url;
  link.style = "display:none";
  link.download = obj.fileName + ".xlsx";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
if (typeof window !== 'undefined') window.pExportExcel = pExportExcel
export default pExportExcel