export type Config = {
  [key: string]: any,
  fileName: string,
  sheets: Sheet[],
  sheetStyle: string,
  rowStyle: string,
  cellStyle: string
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
  colSpan: number,
  rowSpan: number,
} | string | number;