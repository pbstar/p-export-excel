import { Config } from './types'
export function checkConfig(e: Config) {
  const config: Config = {
    fileName: '文件',
    fileType: 'xlsx',
    sheets: [],
    sheetStyle: '',
    rowStyle: '',
    cellStyle: '',
    resType: 'download',
    isPreview: false
  }
  if (!e) {
    console.error('请传入配置参数')
    return false
  }
  if (e.fileName) {
    if (typeof e.fileName !== 'string') {
      console.error('fileName参数类型错误，请传入字符串')
      return false
    }
    if (e.fileName.length > 50) {
      console.error('fileName参数长度不能超过50')
      return false
    }
    if (!/^[a-zA-Z0-9\u4e00-\u9fa5_-]+$/.test(e.fileName)) {
      console.error('fileName参数不合法，请传入合法的文件名')
      return false
    }
    config.fileName = e.fileName
  }
  if (e.fileType) {
    if (typeof e.fileType !== 'string') {
      console.error('fileType参数类型错误，请传入字符串')
      return false
    }
    if (!['xlsx', 'csv'].includes(e.fileType)) {
      console.error('fileType参数不合法，请传入xlsx或csv')
      return false
    }
    config.fileType = e.fileType
  }
  if (e.sheets) {
    if (!Array.isArray(e.sheets)) {
      console.error('sheets参数类型错误，请传入数组')
      return false
    }
    config.sheets = e.sheets
  }
  if (e.sheetStyle) {
    if (typeof e.sheetStyle !== 'string') {
      console.error('sheetStyle参数类型错误，请传入字符串')
      return false
    }
    config.sheetStyle = e.sheetStyle
  }
  if (e.rowStyle) {
    if (typeof e.rowStyle !== 'string') {
      console.error('rowStyle参数类型错误，请传入字符串')
      return false
    }
    config.rowStyle = e.rowStyle
  }
  if (e.cellStyle) {
    if (typeof e.cellStyle !== 'string') {
      console.error('cellStyle参数类型错误，请传入字符串')
      return false
    }
    config.cellStyle = e.cellStyle
  }
  if (e.resType) {
    if (typeof e.resType !== 'string') {
      console.error('resType参数类型错误，请传入字符串')
      return false
    }
    if (!['download', 'blob', 'bloburl', 'file', 'base64'].includes(e.resType)) {
      console.error('resType参数不合法，请传入download、blob、bloburl、file或base64')
      return false
    }
    config.resType = e.resType
  }
  if (e.isPreview || e.isPreview === false) {
    if (typeof e.isPreview !== 'boolean') {
      console.error('isPreview参数类型错误，请传入布尔值')
      return false
    }
    config.isPreview = e.isPreview
  }
  return config
}