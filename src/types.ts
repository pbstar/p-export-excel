export type Config = {
  [key: string]: any,
  fileName: string,
  fileType: 'xlsx' | 'csv',
  sheets: Sheet[],
  sheetStyle: string,
  rowStyle: string,
  cellStyle: string,
  resType: 'download' | 'file' | 'base64' | 'blob'
}
export type Sheet = {
  [key: string]: any,
  sheetName: string,
  style: string,
  rows: Row[]
}
export type Row = {
  [key: string]: any,
  style: string,
  cells: Cell[]
}
export type Cell = {
  [key: string]: any,
  text: string | number,
  style: string,
  colspan: number,
  rowspan: number,
} | string | number;