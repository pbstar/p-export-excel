const path = require('path')

module.exports = {
  mode: 'production',
  entry: './src/index.js',
  externals: 'lodash',  //  + 新增
  output: {
    path: path.resolve(__dirname, 'lib'),
    filename: 'p-export-excel.js',
    library: 'library',
    libraryTarget: 'umd'
  }
}
