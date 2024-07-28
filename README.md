## p-export-excel 官方文档

[![](https://img.shields.io/badge/GitHub-E34C26.svg)](https://github.com/pbstar/p-export-excel)
[![GitHub license](https://img.shields.io/github/license/pbstar/p-export-excel?style=flat&color=109BCD)](https://github.com/pbstar/p-export-excel?tab=MIT-1-ov-file#readme)
[![GitHub stars](https://img.shields.io/github/stars/pbstar/p-export-excel?style=flat&color=d48806)](https://github.com/pbstar/p-export-excel/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/pbstar/p-export-excel?style=flat&color=C6538C)](https://github.com/pbstar/p-export-excel/forks)
[![NPM Version](https://img.shields.io/npm/v/p-export-excel?style=flat&color=d4b106)](https://www.npmjs.com/package/p-export-excel)
[![npm bundle size](https://img.shields.io/bundlephobia/min/p-export-excel?style=flat&color=41B883)](https://www.npmjs.com/package/p-export-excel)

p-export-excel 是一个导出 Excel 的 js 插件。它有着轻量且高效的特性，支持丰富的自定义配置选项。

### 配置

- el: 滚动容器的 DOM 元素。
- direction: 滚动方向，可选值包括 'up' (默认) 、 'down' 、 'left' 、 'right'。
- speed: 滚动速度，以毫秒为单位，默认为 100。
- hoverStop: 是否在鼠标移入时停止滚动，默认为 false。
- auto: 是否自动开始滚动，默认为 true。
- loop: 是否循环滚动，默认为 true。
- rest: 在滚动一段距离后停留一段时间，默认为 null，例如{distance: 100, time: 2000}。
  - distance: 停留前滚动的距离，以 px 为单位，必须为 10 的整数倍，默认为 100。
  - time: 停留的时间，以毫秒为单位，默认为 2000。

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
const sheets = [
  {
    table: {
      rows: [
        {
          cells: ["Cell 1", "Cell 2", "Cell 3"],
        },
      ],
    },
  },
];
pExportExcel({
  fileName: "示例数据",
  sheets: sheets,
});
```
