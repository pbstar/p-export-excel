export type Config = {
  [key: string]: any,
  fileName: string,
  sheets: Sheet[]
}
export type Sheet = {
  [key: string]: any,
  sheetName: string,
  table: Table
}
export type Table = {
  [key: string]: any,
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
  text: string,
  style: string,
  colSpan: number,
  rowSpan: number,
} | string | number;