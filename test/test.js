/**
 * title 表名称 string
 * column 表头 string
 * {
 *    label: string 名称,
 *    prop: string 对应data字段,
 *    placeholder: boolean 是否把单元格变为占位符，用于头部行合并,
 *    headerslot: boolean 是否在表头上方插入数据,
 *    bottomslot: boolean 是否在表底部插入数据,
 *    style: string 单元格样式,
 *    colspan: number 列合并,
 *    rowspan: number 行合并,
 *    noExport: boolean 是否导出此列,
 *    exceltype: string 列类型
 * }
 */

const excel = {
  title: '',
  sheets: [],
  /**
   * name: sheet名称
   * sheets: 页签list
   * column: 表头
   * data: 数据
   * content table html
   * retract 缩进字段 string | array
   * @param option { title: string, sheets: [{ name: string, column：Array, retract: string | array, data：Array, content: table }] }
   */
  option: function (option = { title: '', sheets: [] }) {
    if (!option.sheets.length) {
      console.error('the column is null')
      return false
    }
    this.title = option.title
    this.sheets = option.sheets
    this.sheets.forEach(item => {
      item.column = this.filterColumn(item.column)
    })
    this.download(this.sheets)
  },

  /**
   * 导出多个sheet页（content 和 tableId必须要有一个）
   * @param contents 数组对象：[{ content:表格内容带table, sheetName: sheet页名称 }]
   */
  tablesToExcel(contents) {
    console.log(contents);
    const uri = 'data:application/vnd.ms-excel;base64,'
    const html_start = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>`
    const template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>`
    const template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`

    const template_HTMLWorksheet = `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
    <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
    <link id="Main-File" rel="Main-File" href="../WorkBook.htm">
    <link rel="File-List" href="filelist.xml">
</head>
<body>{SheetContent}</body>
</html>`

    const template_WorkBook = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
    <x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
    <x:ActiveSheet>0</x:ActiveSheet>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
    <frame src="sheet0.htm" name="frSheet">
    <noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
    <o:MainFile HRef="../WorkBook.htm"/>
    {ListWorksheets}
    <o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`
    const base64 = (s) => {
      return window.btoa(unescape(encodeURIComponent(s)))
    }

    const format = (s, c) => {
      return s.replace(/{(\w+)}/g, function (m, p) {
        return c[p]
      })
    }

    const context_WorkBook = {
      ExcelWorksheets: '',
      HTMLWorksheets: '',
      ListWorksheets: ''
    }

    contents.forEach((p, SheetIndex) => {
      const SheetName = p.sheetName || 'Sheet' + SheetIndex
      context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
        SheetIndex: SheetIndex,
        SheetName: SheetName
      })
      let content = p.content
      if (p.tableId) {
        content = document.getElementById(p.tableId).outerHTML
      }
      context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
        SheetIndex: SheetIndex,
        SheetContent: content
      })
      context_WorkBook.ListWorksheets += format(template_ListWorksheet, {
        SheetIndex: SheetIndex
      })
    })
    return uri + base64(format(template_WorkBook, context_WorkBook))
  },
  /**
   * 下载
   */
  download: async function (sheets) {
    this.generateTable(sheets, (list) => {
      const worksheet = 'ceshi'
      const sheetList = list.map(item => {
        return {
          content: item.content,
          sheetName: item.name
        }
      })
      const a = document.createElement('a')
      a.href = this.tablesToExcel(sheetList)
      a.download = `${worksheet}.xlsx`
      a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }))
      this.done()
    })
  },
  /**
   * 生成table
   */
  generateTable: function (sheets, callback) {
    let executeNumber = 0
    sheets.forEach(async sheet => {
      let tableHead = ''
      let tableBody = ''
      // 判断是否需要复杂表头 （tips：目前只支持表头两层嵌套）
      if (sheet.column.some(item => item.children && item.children.length && !item.headerslot && !item.bottomslot)) {
        tableHead += await this.generateColspanHead(sheet.column)
        sheet.column = await this.flagColumn(sheet.column)
      }
      tableHead += this.generateHead(sheet.column)
      tableBody = this.generateBody(sheet) || ''
      sheet.content = `<table border="1" cellpadding="0" cellspacing="0" style="vnd.ms-excel.numberformat:@;border-collapse:collapse; text-align: center;">
            ${this.generateSlot(sheet.column, 'headerslot')}
            ${tableHead}
            ${tableBody}
            ${this.generateSlot(sheet.column, 'bottomslot')}
          </table>`
      executeNumber++
      if (executeNumber === sheets.length) {
        callback(sheets)
      }
    })
  },
  /**
   * 生成table header
   */
  generateHead: function (column) {
    let tableHead = '<tr>'
    for (let i = 0; i < column.length; i++) {
      const item = column[i]
      if (item.headerslot || item.bottomslot) {
        continue
      }
      if (item.placeholder) {
        continue
      }
      tableHead += `<td style="font-weight: bold">${item.label}</td>`
    }
    tableHead += '</tr>'
    return tableHead
  },
  /**
   * 生成复杂 head
   */
  generateColspanHead: function (column) {
    let tableHead = '<tr>'
    for (let i = 0; i < column.length; i++) {
      const item = column[i]
      if (item.headerslot || item.bottomslot) {
        continue
      }
      if (item.children && item.children.length) {
        tableHead += `<td style="font-weight: bold" colspan="${item.children?.length}">${item.label}</td>`
      } else {
        tableHead += `<td style="font-weight: bold" rowspan="2">${item.label}</td>`
      }
    }
    tableHead += '</tr>'
    return tableHead
  },
  /**
   * 生成table 主体
   */
  generateBody: function (sheet, level = 0) {
    if (!sheet.data.length) {
      return
    }
    let tableBody = ''
    sheet.data.forEach(item => {
      tableBody += `<tr>`
      for (let i = 0; i < sheet.column.length; i++) {
        const col = sheet.column[i]
        if (!col.headerslot && !col.bottomslot) {
          if (isNaN(Number(item[col.prop])) || col.exceltype === 'string') {
            tableBody += `<td style="padding-left: ${sheet.retract && (sheet.retract.includes(col.prop) || sheet.retract === col.prop) ? level * 30 : 0}px">${item[col.prop] || ''}</td>`
          } else {
            tableBody += `<td style="vnd.ms-excel.numberformat:#,##0.00;">${Number(item[col.prop]) || ''}</td>`
          }
        }
      }
      tableBody += `</tr>`
      if (item.children && item.children.length) {
        tableBody += this.generateBody({ ...sheet, data: item.children }, level + 1)
      }
    })
    return tableBody
  },
  /**
   * slottype：[headerslot, bottomslot]
   * 生成底部或者头的插入行
   */
  generateSlot: function (column, slottype) {
    let slot = ''
    for (let i = 0; i < column.length; i++) {
      const item = column[i]
      if (item[slottype]) {
        if (item.children && item.children.length) {
          slot += `<tr>`
          item.children.forEach(d => {
            slot += `<td rowspan="${d.rowspan}" colspan="${d.colspan}" style="${d.style}">${d.label}</td>`
          })
          slot += `</tr>`
          continue
        }
        slot += `
          <tr>
            <td rowspan="${item.rowspan}" colspan="${item.colspan}" style="${item.style}">${item.label}</td>
          </tr>
        `
      }
    }
    return slot
  },
  /**
   * 过滤不需导入的列
   */
  filterColumn(column) {
    const newColumn = []
    for (let i = 0; i < column.length; i++) {
      const item = column[i]
      if (!item.noExport) {
        newColumn.push(item)
      }
    }
    return newColumn
  },
  /**
   * 扁平化head column
   */
  flagColumn: function (column) {
    const newColumn = []
    for (let i = 0; i < column.length; i++) {
      const item = column[i]
      if (item.children && item.children.length && !item.headerslot && !item.bottomslot) {
        newColumn.push(...item.children)
        continue
      }
      item.placeholder = true
      newColumn.push(item)
    }
    return newColumn
  },
  /**
   * 下载完成
   */
  done() {
    this.sheets = []
    this.title = ''
  }
}
