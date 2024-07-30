import { Config } from './types'
import { createXlsx, createCsv } from './create'
import { checkConfig } from './check'
import result from './result';
function pExportExcel(e: any) {
  return new Promise((resolve, reject) => {
    const checkResult: any = checkConfig(e);
    if (typeof checkResult == 'boolean') return reject('error:config error');
    const config: Config = checkResult;
    const resBase64 = config.fileType == 'xlsx' ? createXlsx(config) : createCsv(config);
    if (!resBase64) return reject('error: create file error');
    const res = result(config, resBase64)
    resolve(res)
  })
}
export default pExportExcel