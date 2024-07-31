import { Config } from './types'
export default function preview(config: Config, createRes: any) {
  return new Promise((resolve, reject) => {
    let res: any = ''
    if (config.fileType == 'xlsx') {
      let arr = createRes.xlsx
      res = `
        <div style="width: 100%;height: 470px;overflow: auto;">
          ${arr.map((item: any, index: number) => {
        return `<div class="pee_tab_content" style="display: ${index === 0 ? 'block' : 'none'};">${item}</div>`
      }).join('')}
        </div>
        <div style="width: 100%;height: 25px;display: flex;align-items: center;margin-top: 4px;box-sizing: border-box;border-left: 1px solid #ccc;border-bottom: 1px solid #ccc;">
          ${config.sheets.map((item: any, index: number) => {
        return `<div class="pee_tab_bar" style="width:100px;height:25px;font-size:12px;cursor: pointer;border: 1px solid #ccc;border-bottom-color: ${index === 0 ? '#1890ff' : 'transparent'};border-left:0;box-sizing: border-box;display:flex;justify-content: center;align-items: center;">${item.sheetName}</div>`
      }).join('')}
        </div>
      `
    } else if (config.fileType == 'csv') {
      res = createRes.csv
      if (res.indexOf("\n") >= 0) res = res.replace("\n", "<br/>");
    }

    let body = document.querySelector('body')
    let view = document.createElement('div')
    view.style.width = '100%'
    view.style.height = '100%'
    view.style.position = 'fixed'
    view.style.top = '0'
    view.style.left = '0'
    view.style.backgroundColor = 'rgba(0,0,0,0.5)'
    view.style.zIndex = '9999'
    view.style.display = 'flex'
    view.style.justifyContent = 'center'
    view.style.alignItems = 'center'
    view.innerHTML = `
    <div style="width: 800px;height: 600px;background-color: white;border-radius: 3px;overflow: hidden;display: flex;flex-direction: column;align-items: center;justify-content: space-between;">
      <div style="width: 100%;height: 50px;padding:0 15px;display: flex;justify-content: space-between;align-items: center;box-sizing: border-box;">
        <div style="font-size: 18px;">预览</div>
        <div id="pee_close" style="cursor: pointer;">X</div>
      </div>
      <div style="width: 100%;height: 500px;overflow: auto;box-sizing: border-box;padding: 0 10px;">
        ${res}
      </div>
      <div style="width: 100%;height: 50px;display: flex;justify-content: center;align-items: center;">
        <button style="width: 100px;height: 30px;lin-height:28px;border: 1px solid #eee;background-color: #fff;color: #333;border-radius: 5px;cursor: pointer;" id="pee_close">取消</button>
        <button style="width: 100px;height: 30px;lin-height:30px;border: none;background-color: #1890ff;color: white;border-radius: 5px;cursor: pointer;margin-left: 15px;" id="pee_ok">确定</button>
      </div>
    </div>
    `
    body?.appendChild(view)
    let close = document.querySelector('#pee_close')
    close?.addEventListener('click', () => {
      toClick(false)
    })
    let ok = document.querySelector('#pee_ok')
    ok?.addEventListener('click', () => {
      toClick(true)
    })
    function toClick(boo: boolean) {
      body?.removeChild(view)
      resolve(boo)
    }
    let pee_tab_bar = document.querySelectorAll('.pee_tab_bar')
    if (pee_tab_bar) {
      pee_tab_bar.forEach((div: any, i: number) => {
        div.addEventListener('click', () => {
          changeTab(i)
        })
      })
    }
    function changeTab(index: number) {
      document.querySelectorAll('.pee_tab_content').forEach((div: any, i: number) => {
        div.style.display = i === index ? 'block' : 'none';
      });
      document.querySelectorAll('.pee_tab_bar').forEach((div: any, i: number) => {
        div.style.borderBottomColor = i === index ? '#1890ff' : '#fff';
      })
    }
  })
}