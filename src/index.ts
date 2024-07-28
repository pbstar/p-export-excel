import { Config } from './types'
import create from './create'
function pExportExcel(e: Config) {
  let obj = {
    fileName: e.fileName || '文件',
    sheets: e.sheets || [],
    sheetStyle: e.sheetStyle || '',
    rowStyle: e.rowStyle || '',
    cellStyle: e.cellStyle || ''
  }
  let excelFile = create(obj);
  let link: any = document.createElement("a");
  link.href = excelFile;
  link.download = obj.fileName + ".xlsx";
  link.click();
}
export default pExportExcel