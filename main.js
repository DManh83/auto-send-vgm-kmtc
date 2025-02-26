// Modules to control application life and create native browser window
const { app, BrowserWindow, ipcMain, globalShortcut, clipboard, dialog } = require('electron')
const path = require('node:path')
const robot = require('robotjs');
const dotenv = require('dotenv');
const XLSX = require('xlsx');
// let robotProcess = null;  // 存储子进程的引用

dotenv.config();

let mainWindow;
let webWindow;

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

function createWindow() {
  // Create the main window.
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,  // 推荐启用上下文隔离
      nodeIntegration: false   // 禁用节点集成
    }
  })
  
  mainWindow.loadFile('index.html')

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

}


async function autoLogin(window) {
  try {
    const loginActions = [
      { type: 'html-input', selector: '#id', content: process.env.ID},
      { type: 'html-input', selector: '#pw', content: process.env.PASSWORD},
      { type: 'html-click', selector: 'a.button.blue.sm'},
    ];
    for (const action of loginActions) {
      await executeAction(action, 1000, window);
    }
    console.log('Login success');
  } catch (error) {
    console.error('autoLogin error:', error);
  }
}

async function searchBooking(window, blNo) {
  const searchActions = [
    { type: 'html-input', selector: '#kind', content: "02" },
    { type: 'html-input', selector: '#keyword', content: blNo },
    { type: 'html-click', selector: 'a.button.blue.sh'},
    // { type: 'enter' },
  ];
  try {
    for (const action of searchActions) {
      await executeAction(action, 1000, window);
    }
    console.log('Search success');
  } catch (error) {
    console.error('searchBooking error:', error);
  }
}

async function inputVGM(window, blNo, filePath) {
  try {
    await executeAction({ type: 'html-click-2', selector: 'a.link', text: 'Input VGM'}, 1000, window);

    const vgmDataList = await getVGMData(blNo, filePath);
    if (!vgmDataList) {
      throw new Error('VGM data not found');
    }

    for (let i = 0; i < vgmDataList.length; i++) {
      const vgmData = vgmDataList[i];
      const inputActions = [
        { type: 'html-input', selector: `#cntr_${i}`, content: vgmData.containerNo },
        { type: 'html-input', selector: `#method_${i}`, content: vgmData.method },
        { type: 'html-input', selector: `#wt_${i}`, content: vgmData.overweight },
        { type: 'html-input', selector: `#sign_${i}`, content: vgmData.signature },
      ];

      for (const action of inputActions) {
        await executeAction(action, 1000, window);
      }
    }

    await executeAction({ x: 716, y: 590 }, 1000);
    //click save button
    // await executeAction({ type: 'html-click-2', selector: 'a.button.blue.lg', text: 'Save'}, 1000, window);
    console.log('Input VGM success');
  } catch (error) {
    console.error('inputVGM error:', error);
  }
}

async function getVGMData(blNo, filePath) {
  try {
    if (!filePath) {
      const response = await fetch(`https://www.dadaex.cn/api/seaOrder/getKmtcOrder?blNo=${blNo}`);
      const data = await response.json();
      // return data;
      if (data.status === 1 && data.data.length > 0) {
      const order = data.data[0].sea_order;
      const vgmData = order.containner_vgms.map(vgm => ({
        containerNo: vgm.boxNub,
        method: "SM2",
        overweight: vgm.overweight ? parseFloat(vgm.overweight).toString() : "",
        signature: vgm.signature || "MYLAN"
      }));
      return vgmData;
    }
    throw new Error('No data found');
  }
  const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    if (data.length > 0) {
      return data.map(row => ({
        containerNo: row['Container No.'],
        method: "SM2",
        overweight: row['Weight(Kg)'],
        signature: "MYLAN"
      }));
    }
    throw new Error('File not contain VGM data');
  } catch (error) {
    console.error('Error getting VGM data:', error);
    return null;
  }
}

async function getSIData(blNo, filePath) {
  try {
    if (filePath) {
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet);
      return data;
    }
    const response = await fetch(`https://www.dadaex.cn/api/seaOrder/getKmtcOrder?blNo=${blNo}`);
    const data = await response.json();
    if (data.status === 1 && data.data.length > 0) {
      const order = data.data[0].sea_order;
      let siData = {};
      const [shMasterMain, shMasterExtra] = (order.sea_order_contacts[0].master || '').split('\n');
      siData.shMaster = shMasterMain;
      siData.shAddress = shMasterExtra ? `${shMasterExtra}\n${order.sea_order_contacts[0].address}` : order.sea_order_contacts[0].address;
      const [cneMasterMain, cneMasterExtra] = (order.sea_order_contacts[1].master || '').split('\n');
      siData.cneMaster = cneMasterMain;
      siData.cneAddress = cneMasterExtra ? `${cneMasterExtra}\n${order.sea_order_contacts[1].address}` : order.sea_order_contacts[1].address;
      const [ntifMasterMain, ntifMasterExtra] = (order.sea_order_contacts[2].master || '').split('\n');
      siData.ntifMaster = ntifMasterMain;
      siData.ntifAddress = ntifMasterExtra ? `${ntifMasterExtra}\n${order.sea_order_contacts[2].address}` : order.sea_order_contacts[2].address;
      siData.cargo = {
        weight: order.goods.reduce((sum, good) => sum + parseFloat(good.weight || 0), 0).toString(),
        amount: order.goods.reduce((sum, good) => sum + parseFloat(good.amount || 0), 0).toString(),
        volume: order.goods.reduce((sum, good) => sum + parseFloat(good.volume || 0), 0).toString(),
        hscode: order.goods[0].hscode.slice(0, 6),
      }
      siData.goods = order.goods;
      siData.containers = order.containner_vgms;
      return siData;
    }
  } catch (error) {
    console.error('Error getting SI data:', error);
    return null;
  }
}

async function inputSIData(window, blNo, filePath) {
  try {
    const siData = await getSIData(blNo, filePath);
    if (!siData) {
      throw new Error('SI data not found');
    }
    // console.log(siData.shMaster);
    // console.log(siData.shAddress);
    // console.log(siData.cneMaster);
    // console.log(siData.cneAddress);
    // console.log(siData.ntifMaster);
    // console.log(siData.ntifAddress);
    // console.log('SI data:', siData);

    // input Cargo Information
    const inputActions1 = [
      { type: 'html-click-2', selector: 'a', text: blNo.slice(4)},
      { type: 'html-click', selector: 'input#agreea1'},
      { type: 'html-click-2', selector: 'a.button.blue.lg', text: 'Confirm'},
      { type: 'html-input', selector: '#shprCstEnm', content: siData.shMaster},
      { type: 'html-input', selector: '#shprCstAddr', content: siData.shAddress},
      { type: 'html-input', selector: '#cneCstEnm', content: siData.cneMaster},
      { type: 'html-input', selector: '#cneCstAddr', content: siData.cneAddress},
      { type: 'html-input', selector: '#ntifCstEnm', content: siData.ntifMaster},
      { type: 'html-input', selector: '#ntifCstAddr', content: siData.ntifAddress},
      { type: 'html-input', selector: '#cneIdCatCd', content: "081"},
      { type: 'html-input', selector: '#grsWt', content: siData.cargo.weight},
      { type: 'html-input', selector: '#pkgQty', content: siData.cargo.amount},
      { type: 'html-input', selector: '#grsCbm', content: siData.cargo.volume},
      { type: 'html-input', selector: '#podHsCd', content: siData.cargo.hscode},
    ];
    for (const action of inputActions1) {
      await executeAction(action, 1000, window);
    }
    // input containers
    for (let i = 0; i < siData.containers.length; i++) {
      const inputActions2 = [
        { type: 'html-input', selector: `#cntrNo_${i}`, content: siData.containers[i].boxNub},
        // { type: 'html-input', selector: `#sealNo_${i}`, content: siData.containers[i].sealNo}, // missing data
        { type: 'html-input', selector: `#wt_${i}`, content: parseFloat(siData.containers[i].weight).toString()},
        // { type: 'html-input', selector: `#cntrPkgQty_${i}`, content: parseFloat(siData.containers[i].amount).toString()}, // missing data
      ];
      for (const action of inputActions2) {
        await executeAction(action, 1000, window);
      }
    }
    // click Auto Say button
    await executeAction({ type: 'html-click-2', selector: 'a.button.sm.ml2', text: 'Auto Say'}, 1000, window);
    // B/L Body Detail
    for (let i = 0; i < siData.goods.length; i++) {
      const inputActions3 = [
        { type: 'html-input', selector: `#mrk_${i}`, content: siData.goods[i].shantou},
        { type: 'html-input', selector: `#dscr_${i}`, content: siData.goods[i].name},
      ];
      for (const action of inputActions3) {
        await executeAction(action, 1000, window);
      }
    }
    console.log('Input SI data success');
  } catch (error) {
    console.error('inputSIData error:', error);
  }
}

async function executeAction(action, delayTime, window = null) {
  robot.moveMouse(action.x, action.y);
  switch (action.type) {
    case 'click':
      robot.mouseClick();
      break;
    case 'text':
      clipboard.writeText(action.content);
      if (process.platform === 'darwin') {
        robot.keyToggle('command', 'down');
        robot.keyTap('a');
        robot.keyTap('v');
        robot.keyToggle('command', 'up');
      } else {
        robot.keyToggle('control', 'down');
        robot.keyTap('a');
        robot.keyTap('v');
        robot.keyToggle('control', 'up');
      }
      break;
    case 'html-input':
      if (!window) {
        throw new Error('WebWindow is not created');
      }
      await window.webContents.executeJavaScript(`
        try {
          console.log('Searching for:', '${action.selector}', 'with text:', '${action.text}');
          const input = document.querySelector('${action.selector}');
          if (input) {
            input.value = \`${action.content.replace(/\\n/g, '\n')}\`;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            // Trigger Angular/React/Vue change detection
            const evt = new InputEvent('input', { bubbles: true, cancelable: true });
            input.dispatchEvent(evt);
          }
          true;
        } catch (err) {
          console.error('Script error:', err);
          false;
        }
      `, true);
      await delay(1000);
      break;
    case 'html-click':
      if (!window) {
        throw new Error('WebWindow is not created');
      }
      await window.webContents.executeJavaScript(`
        document.querySelector('${action.selector}').click();
        `);
        await delay(1000);
        break;
    case 'html-click-2':
      if (!window) {
        throw new Error('WebWindow is not created');
      }
      await window.webContents.executeJavaScript(`
        try {
          const links = Array.from(document.querySelectorAll('${action.selector}'));
          const targetLink = links.find(link => link.innerText === '${action.text}');
          if (targetLink) {
            targetLink.click();
            true;
          } else {
            throw new Error('Link not found: ${action.text}');
          }
        } catch (error) {
          console.error('Click error:', error);
          throw error;
        }
      `, true);
      await delay(1000);
      break;
    case 'enter':
      robot.keyTap('enter');
      break;
    default:
      break;
  }

  await delay(delayTime);
}

async function createWebWindow() {
  // Create the web window.
  webWindow = new BrowserWindow({
    width: 1500,
    height: 800,
    x: 0,
    y: 0,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,  // 推荐启用上下文隔离
      nodeIntegration: false   // 禁用节点集成
    }
  })
  // Open kmtc page
  await webWindow.loadURL('https://www.ekmtc.com/index.html#/working/progress');

  webWindow.on('closed', () => {
    webWindow = null;
  });

  return webWindow;
}

app.whenReady().then(() => {
  createWindow()

  globalShortcut.register('CommandOrControl+Shift+E', () => {
    const windows = BrowserWindow.getAllWindows();
    windows.forEach(window => {
      window.close();
    });
  });

  globalShortcut.register('CommandOrControl+Shift+X', () => {
    const mousePos = robot.getMousePos();
    console.log('Mouse position:', mousePos);
    BrowserWindow.getAllWindows().forEach(window => {
      window.webContents.send('mouse-position', mousePos);
    });
  });

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('send-vgm', async (event, blNo, filePath) => {
  try {
    const window = await createWebWindow();

    await delay(3000);
    await autoLogin(window);
    await delay(3000);
    await searchBooking(window, blNo);
    await delay(3000);
    await inputVGM(window, blNo, filePath);

    await delay(2000);

    // if (window && !window.isDestroyed()) {
    //   window.close();
    // }

    await new Promise(resolve => {
      window.webContents.once('did-finish-load',  resolve);
    });

    return { success: true };
  } catch (error) {
    console.error('Error when processing BlNo:', error);
    if (window && !window.isDestroyed()) {
      window.close();
    }
    return { success: false, error: error.message };
  }
});

ipcMain.handle('submit-si', async (event, blNo, filePath) => {
  try {
    const window = await createWebWindow();

    await delay(3000);
    await autoLogin(window);
    await delay(3000);
    await searchBooking(window, blNo);
    await delay(3000);
    await inputSIData(window, blNo, filePath);

    await new Promise(resolve => {
      window.webContents.once('did-finish-load',  resolve);
    });

  } catch (error) {
    console.error('Error when processing BlNo:', error);
    if (window && !window.isDestroyed()) {
      window.close();
    }
    return { success: false, error: error.message };
  }
});


