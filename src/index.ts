import { Config } from './types'
import create from './create'
import { checkConfig } from './check'
import result from './result';
import preview from './preview'
function pExportExcel(e: any) {
  return new Promise((resolve, reject) => {
    const checkResult: any = checkConfig(e);
    if (typeof checkResult == 'boolean') return reject('error:config error');

    const config: Config = checkResult;

    const createRes = create(config);
    if (!createRes.base64) return reject('error: create file error');


    if (config.isPreview) {
      preview(config, createRes).then((e) => {
        if (e) {
          const res = result(config, createRes.base64)
          resolve(res)
        } else {
          reject('error: preview error')
        }
      })
    } else {
      const res = result(config, createRes.base64)
      resolve(res)
    }
  })
}
export default pExportExcel