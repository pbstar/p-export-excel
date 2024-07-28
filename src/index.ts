import { Config } from './types'
import create from './create'
function pExportExcel(e: Config) {
  let obj = {
    fileName: e.fileName || '文件名',
    sheets: e.sheets || []
  }
  let excelFile = create(obj.sheets);
  let link: any = document.createElement("a");
  link.href = excelFile;
  link.style = "display:none";
  link.download = obj.fileName + ".xlsx";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
export default pExportExcel