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

async function autoLogin() {
  try {
    const loginActions = [
      // { type: 'click', x: 1369, y: 87 },
      { type: 'click', x: 1140, y: 174 },
      { type: 'text', content: process.env.ID },
      { type: 'click', x: 1140, y: 202 },
      { type: 'text', content: process.env.PASSWORD },
      { type: 'click', x: 1280, y: 266 },
    ];
    for (const action of loginActions) {
      await executeAction(action, 1000);
    }
    console.log('Login success');
  } catch (error) {
    console.error('autoLogin error:', error);
  }
}

async function searchBooking(blNo) {
  const searchActions = [
    { type: 'click', x: 420, y: 265 },
    { type: 'click', x: 420, y: 345 },
    { type: 'click', x: 600, y: 265 },
    { type: 'text', content: blNo },
    { type: 'enter' },
  ];
  try {
    for (const action of searchActions) {
      await executeAction(action, 1000);
    }
    console.log('Search success');
  } catch (error) {
    console.error('searchBooking error:', error);
  }
}

async function inputVGM(window, vgmDataList) {
  try {
    await executeAction({ type: 'click', x: 370, y: 660 }, 1000);

    for (let i = 0; i < vgmDataList.length; i++) {
      const vgmData = vgmDataList[i];
      const inputActions = [
        { type: 'html', selector: `#cntr_${i}`, content: vgmData.containerNo },
        { type: 'html', selector: `#method_${i}`, content: vgmData.method },
        { type: 'html', selector: `#wt_${i}`, content: vgmData.overweight },
        { type: 'html', selector: `#sign_${i}`, content: vgmData.signature },
      ];

      for (const action of inputActions) {
        await executeAction(action, 1000, window);
      }
    }

    await executeAction({ x: 716, y: 590 }, 1000);
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
        overweight: vgm.weight ? parseFloat(vgm.weight).toString() : "",
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

async function executeAction(action, delayTime, window = null) {
  robot.moveMouse(action.x, action.y);
  switch (action.type) {
    case 'click':
      robot.mouseClick();
      break;
    case 'double-click':
      robot.doubleClick();
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
    case 'html':
      if (!window) {
        throw new Error('WebWindow is not created');
      }
      await delay(500);
      await window.webContents.executeJavaScript(`
        document.querySelector('${action.selector}').value = '${action.content}';
      `);
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

ipcMain.handle('submit-button', async (event, blNo, filePath) => {
  try {
    const vgmData = await getVGMData(blNo, filePath);
    if (!vgmData) {
      throw new Error('VGM data not found');
    }

    console.log('VGM data:', vgmData);

    const window = await createWebWindow();

    await delay(3000);
    await autoLogin();
    await delay(3000);
    await searchBooking(blNo);
    await delay(3000);
    await inputVGM(window, vgmData);

    await delay(2000);

    if (window && !window.isDestroyed()) {
      window.close();
    }

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