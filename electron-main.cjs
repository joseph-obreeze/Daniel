const { app, BrowserWindow, Menu, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

// 외부 라이브러리 없이 자체 기능으로 개발 모드 확인
const isDev = !app.isPackaged;

// 엑셀 파일 저장 처리기
ipcMain.handle('save-excel', async (event, { content, fileName }) => {
  const { filePath } = await dialog.showSaveDialog({
    title: '엑셀 파일 저장',
    defaultPath: path.join(app.getPath('downloads'), fileName),
    filters: [
      { name: 'Excel Files', extensions: ['xlsx'] }
    ]
  });

  if (filePath) {
    try {
      // Base64 데이터를 버퍼로 변환하여 저장
      const buffer = Buffer.from(content, 'base64');
      fs.writeFileSync(filePath, buffer);
      return { success: true, filePath };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
  return { success: false, canceled: true };
});

function createWindow() {
  const win = new BrowserWindow({
    width: 1300,
    height: 900,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
    title: "판매분석 솔루션",
    icon: path.join(__dirname, isDev ? 'public/icon.png' : 'dist/icon.png')
  });

  // 메뉴 바 제거 (필요시)
  // Menu.setApplicationMenu(null);

  if (isDev) {
    // 개발 모드: Vite 개발 서버 로드
    win.loadURL('http://localhost:5173');
    // win.webContents.openDevTools(); // 개발자 도구 자동 열기
  } else {
    // 프로덕션 모드: 빌드된 파일 로드
    win.loadFile(path.join(__dirname, 'dist/index.html'));
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
