import { Config } from './types'
export default function result(config: Config, resBase64: string) {
  if (config.resType == 'base64') {
    return resBase64
  } else if (config.resType == 'blob') {
    let blob = new Blob([resBase64], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob
  } else if (config.resType == 'file') {
    let file = new File([resBase64], config.fileName + "." + config.fileType, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return file
  } else if (config.resType == 'download') {
    let link: any = document.createElement("a");
    link.href = resBase64;
    link.download = config.fileName + "." + config.fileType;
    link.click();
    return 'success'
  } else if (config.resType == 'bloburl') {
    let blob = new Blob([resBase64], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return URL.createObjectURL(blob);
  }
}