## p-export-excel 官方文档

[![](https://img.shields.io/badge/GitHub-E34C26.svg)](https://github.com/pbstar/p-export-excel)
[![GitHub license](https://img.shields.io/github/license/pbstar/p-export-excel?style=flat&color=109BCD)](https://github.com/pbstar/p-export-excel?tab=MIT-1-ov-file#readme)
[![GitHub stars](https://img.shields.io/github/stars/pbstar/p-export-excel?style=flat&color=d48806)](https://github.com/pbstar/p-export-excel/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/pbstar/p-export-excel?style=flat&color=C6538C)](https://github.com/pbstar/p-export-excel/forks)
[![NPM Version](https://img.shields.io/npm/v/p-export-excel?style=flat&color=d4b106)](https://www.npmjs.com/package/p-export-excel)
[![npm bundle size](https://img.shields.io/bundlephobia/min/p-export-excel?style=flat&color=41B883)](https://www.npmjs.com/package/p-export-excel)

p-export-excel 是一个导出 Excel 的 js 插件。它有着轻量且高效的特性，支持丰富的自定义配置选项。

![请添加图片描述](https://i-blog.csdnimg.cn/direct/5ff6fea5706c4512827af7853e5c813d.png)

### 配置

- fileName: 文件名，字符串。
- fileType: 文件类型，字符串（值为 xlsx、csv，默认为 xlsx）。
- sheetStyle: 工作表样式，支持 css 样式，字符串。
- rowStyle: 数据行样式，支持 css 样式，字符串。
- cellStyle: 单元格样式，支持 css 样式，字符串。
- resType: 资源类型，字符串（值为 download、file、base64、blob 以及 bloburl，默认为 download）。
- sheets: 工作表，数组（值为对象）。
  - sheetName: 工作表名称，字符串。
  - style: 工作表样式，支持 css 样式，字符串。
  - rowStyle: 数据行样式，支持 css 样式，字符串。
  - cellStyle: 单元格样式，支持 css 样式，字符串。
  - rows: 数据行，数组（值为对象）。
    - style: 数据行样式，支持 css 样式，字符串。
    - cellStyle: 单元格样式，支持 css 样式，字符串。
    - cells: 单元格，数组（值为对象或数值或字符串）。
      - text: 单元格内容，字符串或数字。
      - style: 单元格样式，支持 css 样式，字符串。
      - colspan: 单元格横向合并列数，数字。
      - rowspan: 单元格纵向合并行数，数字。

### 安装引入

#### npm 安装

```bash
npm install p-export-excel --save
```

#### esm 引入

```javascript
import pExportExcel from "p-export-excel";
```

#### cdn 引入

```html
<script src="https://unpkg.com/p-export-excel@[version]/lib/p-export-excel.umd.js"></script>
```

### 使用示例

```javascript
// 简约数据表
const option = {
  fileName: "示例数据表",
  sheets: [
    {
      rows: [
        {
          cells: ["Cell 1", "Cell 2", "Cell 3"],
        },
        {
          cells: ["Cell 4", "Cell 5", "Cell 5"],
        },
      ],
    },
  ],
};
// 复杂数据表
const option = {
  fileName: "复杂数据表",
  rowStyle: "height: 40px;vertical-align: middle;",
  cellStyle: "border: 1px solid #eee;width: 100px;text-align: center;",
  sheets: [
    {
      sheetName: "学生名单",
      style: "font-size: 16px;font-family: 微软雅黑;",
      rows: [
        {
          style: "font-weight: bold;color: #fff;font-size: 18px;",
          cellStyle: "background-color: #29B8DB;border: 1px solid #fff;",
          cells: [
            {
              text: "三年级（1）班全体学生名单",
              colspan: 5,
            },
          ],
        },
        {
          style: "font-weight: bold;color: #fff;",
          cellStyle: "background-color: #29B8DB;border: 1px solid #fff;",
          cells: [
            "序号",
            "姓名",
            "性别",
            "年龄",
            {
              text: "备注",
              style: "width: 200px;",
            },
          ],
        },
        {
          cells: ["1", "张三", "男", 12, ""],
        },
        {
          cellStyle: "background-color: #F5F5F5;",
          cells: ["2", "李四", "女", 11, ""],
        },
        {
          cells: ["3", "王五", "男", 12, "校长侄子"],
        },
        {
          cellStyle: "background-color: #F5F5F5;",
          cells: ["4", "赵六", "女", 11, ""],
        },
      ],
    },
  ],
};
pExportExcel(option)
  .then((e) => {
    console.log(e);
  })
  .catch((err) => {
    console.log(err);
  });
```

### 注意事项

- 1.xlsx 文件的字体单位为磅（pt），所以样式中设置的字体大小（px）将被转换，具体公式为 pt≈72\*px/DPI。
- 2.xlsx 文件重复设置样式会按照 css 的原则，后者覆盖前者。
- 3.csv 文件多表导出仅会导出第一个表。
- 4.csv 文件仅导出文本，不支持样式以及单元格合并等。
- 5.所有配置项均为非必填项。
