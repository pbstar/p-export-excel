import { Config } from './types'
import {createXlsx} from './create'
function pExportExcel(e: Config) {
  let obj = {
    fileName: e.fileName || '文件',
    fileType: e.fileType || 'xlsx',
    sheets: e.sheets || [],
    sheetStyle: e.sheetStyle || '',
    rowStyle: e.rowStyle || '',
    cellStyle: e.cellStyle || '',
    resType: e.resType || 'download'
  }
  let excelFile = createXlsx(obj);
  let link: any = document.createElement("a");
  link.href = excelFile;
  link.download = obj.fileName + "." + obj.fileType;
  link.click();
}
export default pExportExcel