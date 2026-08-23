import { app, BrowserWindow, session, shell } from 'electron';
import path from 'node:path';

import {
  CHROME_USER_AGENT,
  GEMINI_APP_URL,
  SESSION_PARTITION,
  decideWindowOpen,
  sanitizeRequestHeaders,
  shouldReloadAfterAuthNavigation,
} from './auth-policy';
import { LINUX_CHROME_COMPAT_SWITCHES } from './chrome-flags';

if (process.platform === 'linux') {
  for (const [name, value] of LINUX_CHROME_COMPAT_SWITCHES) {
    if (value === undefined) {
      app.commandLine.appendSwitch(name);
    } else {
      app.commandLine.appendSwitch(name, value);
    }
  }
}

let mainWindow: BrowserWindow | null = null;

function geminiSession() {
  return session.fromPartition(SESSION_PARTITION);
}

function asStringHeaders(
  headers: Record<string, string | string[]>,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(headers)) {
    out[key] = Array.isArray(value) ? value.join(',') : value;
  }
  return out;
}

function applyChromeIdentity(ses: Electron.Session): void {
  ses.setUserAgent(CHROME_USER_AGENT);
  ses.webRequest.onBeforeSendHeaders((details, callback) => {
    callback({
      requestHeaders: sanitizeRequestHeaders(asStringHeaders(details.requestHeaders)),
    });
  });
}

function attachAuthWindowPolicy(win: BrowserWindow): void {
  win.webContents.setWindowOpenHandler(({ url }) => {
    const decision = decideWindowOpen(url);
    if (decision === 'deny') {
      if (/^https?:/i.test(url)) {
        void shell.openExternal(url);
      }
      return { action: 'deny' };
    }

    return {
      action: 'allow',
      overrideBrowserWindowOptions: {
        width: 520,
        height: 720,
        autoHideMenuBar: true,
        webPreferences: {
          partition: SESSION_PARTITION,
          contextIsolation: true,
          nodeIntegration: false,
          sandbox: true,
        },
      },
    };
  });

  win.webContents.on('did-create-window', (child) => {
    child.webContents.on('did-navigate', (_event, url) => {
      if (shouldReloadAfterAuthNavigation(url) && mainWindow && !mainWindow.isDestroyed()) {
        child.close();
        void mainWindow.loadURL(GEMINI_APP_URL);
      }
    });
  });
}

function createMainWindow(): BrowserWindow {
  const win = new BrowserWindow({
    width: 1280,
    height: 840,
    minWidth: 880,
    minHeight: 600,
    title: 'Gemini',
    autoHideMenuBar: true,
    backgroundColor: '#0b0f14',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      partition: SESSION_PARTITION,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  attachAuthWindowPolicy(win);
  win.webContents.on('did-fail-load', (_event, code, desc, url, isMainFrame) => {
    if (isMainFrame) {
      console.error(`主窗口加载失败 ${code} ${desc} ${url}`);
    }
  });
  win.webContents.on('did-finish-load', () => {
    console.log(`主窗口已加载: ${win.webContents.getURL()}`);
  });
  void win.loadURL(GEMINI_APP_URL);
  return win;
}

function registerApp(): void {
  app.setName('Gemini');
  app.userAgentFallback = CHROME_USER_AGENT;
  app.commandLine.appendSwitch('disable-features', 'ElectronSerialChooser');

  const ses = geminiSession();
  applyChromeIdentity(ses);
  applyChromeIdentity(session.defaultSession);

  mainWindow = createMainWindow();
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (!mainWindow) {
      return;
    }
    if (mainWindow.isMinimized()) {
      mainWindow.restore();
    }
    mainWindow.focus();
  });

  app.whenReady().then(() => {
    registerApp();
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit();
    }
  });

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      registerApp();
    }
  });
}
