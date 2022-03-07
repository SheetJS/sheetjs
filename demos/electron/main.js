/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* from the electron quick-start */
var electron = require('electron');
var XLSX = require('xlsx');
var app = electron.app;
require('@electron/remote/main').initialize();

var win = null;

function createWindow() {
	if (win) return;
	win = new electron.BrowserWindow({
		width: 800, height: 600,
		webPreferences: {
			worldSafeExecuteJavaScript: true, // required for Electron 12+
			contextIsolation: false, // required for Electron 12+
			nodeIntegration: true,
			enableRemoteModule: true
		}
	});
	win.loadURL("file://" + __dirname + "/index.html");
	require('@electron/remote/main').enable(win.webContents);
	win.webContents.openDevTools();
	win.on('closed', function () { win = null; });
}
if (app.setAboutPanelOptions) app.setAboutPanelOptions({ applicationName: 'sheetjs-electron', applicationVersion: "XLSX " + XLSX.version, copyright: "(C) 2017-present SheetJS LLC" });
app.on('open-file', function () { console.log(arguments); });
app.on('ready', createWindow);
app.on('activate', createWindow);
app.on('window-all-closed', function () { if (process.platform !== 'darwin') app.quit(); });
