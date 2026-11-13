const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const { spawn, spawnSync } = require('child_process')
const fs = require('fs')
const os = require('os')


// ─── Python 路径（运行时检测，避免 isPackaged 在模块加载时出错）───────────
function getPythonPath() {
  if (!app.isPackaged) {
    return process.platform === 'win32' ? 'python' : 'python3'
  }
  // extraResources: python/bundle/ -> Resources/python/bundle/
  const base = path.join(process.resourcesPath, 'python')
  if (process.platform === 'win32') {
    const winPython = path.join(base, 'bundle', 'python.exe')
    if (fs.existsSync(winPython)) return winPython
    const fallback = path.join(base, 'python.exe')
    return fs.existsSync(fallback) ? fallback : 'python'
  }
  // macOS: python-build-standalone: python/bundle/bin/python3
  const macPython = path.join(base, 'bundle', 'bin', 'python3')
  if (fs.existsSync(macPython)) return macPython
  return 'python3'
}

// 启动时确保 openpyxl 已安装（兜底 afterPack 脚本未运行的场景）
function ensurePythonDeps() {
  if (!app.isPackaged) return
  const pythonPath = getPythonPath()
  try {
    require('child_process').execFileSync(pythonPath, ['-c', 'import openpyxl'], {
      timeout: 8000,
      windowsHide: true,
    })
    return // openpyxl already available
  } catch (_) {}

  console.log('[openpyxl] not found, attempting install via pip...')
  // 优先用 pip install --user（不会破坏系统 Python 环境）
  const installCmd =
    `${pythonPath} -m pip install --user openpyxl PyYAML rich uiautomation pywin32 Pillow 2>&1`
  try {
    const out = require('child_process').execSync(installCmd, { timeout: 120 * 1000 })
    console.log('[openpyxl] install output:', out.toString())
  } catch (e) {
    console.error('[openpyxl] pip install --user failed:', e.message)
    // 最后的 fallback：尝试系统 pip 直接安装到用户目录
    try {
      const fallback = require('child_process').execSync(
        `python3 -m pip install --user openpyxl PyYAML rich uiautomation pywin32 Pillow 2>&1`,
        { timeout: 120 * 1000 }
      )
      console.log('[openpyxl] fallback install output:', fallback.toString())
    } catch (e2) {
      console.error('[openpyxl] all install attempts failed:', e2.message)
    }
  }
}

function getScriptPath(name) {
  // extraResources: python/ -> Resources/python/
  if (!app.isPackaged) {
    return path.join(__dirname, '..', 'python', name)
  }
  return path.join(process.resourcesPath, 'python', name)
}

// ─── 加载 dist/（生产）vs localhost:5173（开发）────────────────────────
function loadMainWindow(win) {
  const distIndex = path.join(__dirname, '..', 'dist', 'index.html')
  if (fs.existsSync(distIndex)) {
    win.loadFile(distIndex)
  } else {
    win.loadURL('http://localhost:5173')
  }
}

// ─── macOS Automation 权限检测 + 引导授权 ───────────────────────────
let automationChecked = false
function checkAutomationPermission() {
  if (process.platform !== 'darwin' || automationChecked) return
  automationChecked = true

  // 等窗口加载完再检测，不要阻塞启动
  setTimeout(() => {
    try {
      const { dialog, shell } = require('electron')
      const result = require('child_process').spawnSync(
        'osascript',
        ['-e', 'tell application "System Events"\n keystroke "x"\n end tell'],
        { timeout: 5000 }
      )
      if (result.status === 0) return // 已有权限
    } catch (_) {}

    const win = BrowserWindow.getAllWindows()[0]
    if (!win || win.isDestroyed()) return

    dialog.showMessageBox(win, {
      type: 'warning',
      title: '需要辅助功能权限',
      message: '微信发送助手需要「辅助功能」权限才能自动发送微信消息。',
      detail: '点击「打开系统设置」后，在「辅助功能」列表中找到「微信发送助手」并勾选启用。',
      buttons: ['打开系统设置', '稍后'],
      defaultId: 0,
      cancelId: 1,
    }).then(({ response }) => {
      if (response === 0) {
        shell.openExternal('x-apple.systempreferences:com.apple.preference.security?Privacy_Automation')
      }
    })
  }, 2000) // 窗口加载完再检测
}

// ─── 主进程 ───────────────────────────────────────────────────────────
let daemonProcess = null

function createWindow() {
  const win = new BrowserWindow({
    width: 980, height: 680,
    minWidth: 800, minHeight: 560,
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: '#F5F5F5',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    icon: path.join(__dirname, '..', 'assets', 'icon.png'),
  })
  loadMainWindow(win)
  return win
}

// ─── IPC ──────────────────────────────────────────────────────────────

ipcMain.handle('get-config', async () => {
  try {
    const p = path.join(os.homedir(), '.wechat-sender', 'config.json')
    if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, 'utf8'))
  } catch (_) {}
  return { send_interval: 5, max_per_minute: 8, poll_seconds: 15, dry_run: false }
})

ipcMain.handle('save-config', async (_, cfg) => {
  try {
    const dir = path.join(os.homedir(), '.wechat-sender')
    fs.mkdirSync(dir, { recursive: true })
    fs.writeFileSync(path.join(dir, 'config.json'), JSON.stringify(cfg, null, 2), 'utf8')
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

ipcMain.handle('parse-excel', async (_, bytes) => {
  const tmp = path.join(os.tmpdir(), 'wst_upload_' + Date.now() + '.xlsx')
  try {
    fs.writeFileSync(tmp, Buffer.from(bytes))
    const pythonBin = getPythonPath()
    const scriptPath = getScriptPath('app/parse_excel.py')
    console.log('[parse-excel] python:', pythonBin, 'script:', scriptPath, 'tmp:', tmp)
    const r = spawnSync(pythonBin, [scriptPath, tmp], {
      encoding: 'utf8', timeout: 15000,
      env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8' }),
    })
    fs.unlinkSync(tmp)
    if (r.status !== 0) {
      console.error('[parse-excel] python error:', r.stderr)
      return { error: r.stderr }
    }
    try {
      return JSON.parse(r.stdout)
    } catch (e) {
      console.error('[parse-excel] JSON parse error:', e, 'raw:', r.stdout.slice(0, 200))
      return { error: '返回数据格式错误: ' + e.message }
    }
  } catch (e) {
    console.error('[parse-excel] exception:', e)
    return { error: e.message }
  }
})

ipcMain.handle('download-template', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const result = await dialog.showSaveDialog(win, {
    title: '保存表格模版',
    defaultPath: path.join(os.homedir(), 'Desktop', 'wechat_template.xlsx'),
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  })
  if (result.canceled) return
  const r = spawnSync(getPythonPath(), [getScriptPath('app/template_gen.py'), result.filePath], {
    encoding: 'utf8',
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8' }),
  })
  return r.status === 0
    ? { ok: true, path: result.filePath }
    : { ok: false, err: r.stderr }
})

ipcMain.handle('send-now', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const proc = spawn(getPythonPath(), [getScriptPath('app/cli.py'), 'send'], {
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }),
    windowsHide: true,
  })
  proc.stdout.on('data', function(d) {
    win.webContents.send('send-progress', d.toString('utf8'))
  })
  proc.stderr.on('data', function(d) {
    win.webContents.send('send-progress', '[ERROR] ' + d.toString('utf8'))
  })
  proc.on('close', function(code) {
    win.webContents.send('send-done', { code: code })
  })
  return { ok: true }
})

ipcMain.handle('send-selected', async (event, taskData) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const jsonArg = Array.isArray(taskData) ? JSON.stringify(taskData) : '[]'
  const pythonBin = getPythonPath()
  const scriptPath = getScriptPath('app/cli.py')
  console.log('[send-selected] python:', pythonBin)
  console.log('[send-selected] script:', scriptPath)
  console.log('[send-selected] tasks:', jsonArg.slice(0, 80))

  // ✅ 改为 spawn 异步方式，避免 execFileSync 同步阻塞主进程导致"未响应"
  const proc = spawn(pythonBin, [scriptPath, 'send', '--tasks-json', jsonArg], {
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }),
    windowsHide: true,
  })

  let stdoutBuf = ''

  proc.stdout.on('data', function(d) {
    const text = d.toString('utf8')
    stdoutBuf += text
    if (!win.isDestroyed()) win.webContents.send('send-progress', text)
    // 实时逐行解析任务状态
    text.split('\n').forEach(function(line) {
      if (!line.trim()) return
      const hasSuccess = line.includes('✅')
      const hasFailed = line.includes('❌')
      if (hasSuccess || hasFailed) {
        const taskMatch = line.match(/^\[(\d+)\/(\d+)\] → (.+?) \[(.+?)\]/)
        if (taskMatch && !win.isDestroyed()) {
          win.webContents.send('task-status-update', {
            target: taskMatch[3].trim(),
            msg_type: taskMatch[4].trim(),
            status: hasSuccess ? 'success' : 'failed',
          })
        }
      }
    })
  })

  proc.stderr.on('data', function(d) {
    const text = d.toString('utf8')
    console.error('[send-selected] stderr:', text)
    if (!win.isDestroyed()) win.webContents.send('send-progress', '[ERROR] ' + text)
  })

  proc.on('close', function(code) {
    console.log('[send-selected] process exited with code', code)
    if (!win.isDestroyed()) win.webContents.send('send-done', { code: code || 0 })
  })

  proc.on('error', function(err) {
    console.error('[send-selected] process error:', err.message)
    if (!win.isDestroyed()) {
      win.webContents.send('send-progress', '[ERROR] ' + err.message)
      win.webContents.send('send-done', { code: 1, error: err.message })
    }
  })

  return { ok: true }
})

ipcMain.handle('stop-send', async () => {
  try {
    fs.writeFileSync(path.join(os.homedir(), '.wechat-sender', 'stop_signal'), String(Date.now()))
  } catch (_) {}
  return { ok: true }
})

ipcMain.handle('daemon-start', async (event) => {
  if (daemonProcess) return { ok: false, error: '已在运行' }
  const win = BrowserWindow.fromWebContents(event.sender)
  daemonProcess = spawn(getPythonPath(), [getScriptPath('app/cli.py'), 'daemon'], {
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }),
    windowsHide: true,
  })
  daemonProcess.stdout.on('data', function(d) {
    if (!win.isDestroyed()) win.webContents.send('daemon-log', d.toString('utf8'))
  })
  daemonProcess.stderr.on('data', function(d) {
    if (!win.isDestroyed()) win.webContents.send('daemon-log', '[ERR] ' + d.toString('utf8'))
  })
  daemonProcess.on('close', function() {
    daemonProcess = null
    if (!win.isDestroyed()) win.webContents.send('daemon-stopped')
  })
  return { ok: true }
})

ipcMain.handle('daemon-stop', async () => {
  if (daemonProcess) { daemonProcess.kill(); daemonProcess = null }
  return { ok: true }
})

ipcMain.handle('daemon-status', async () => ({ running: daemonProcess !== null }))

ipcMain.handle('add-schedule', async (_, schedule) => {
  const f = path.join(os.homedir(), '.wechat-sender', 'schedules.json')
  let arr = []
  try { arr = JSON.parse(fs.readFileSync(f, 'utf8')) } catch (_) {}
  arr.push(schedule)
  fs.writeFileSync(f, JSON.stringify(arr, null, 2))
  return { ok: true }
})

ipcMain.handle('remove-schedule', async (_, id) => {
  const f = path.join(os.homedir(), '.wechat-sender', 'schedules.json')
  let arr = []
  try { arr = JSON.parse(fs.readFileSync(f, 'utf8')) } catch (_) {}
  arr = arr.filter(function(s) { return s.id !== id })
  fs.writeFileSync(f, JSON.stringify(arr, null, 2))
  return { ok: true }
})

// ── 任务持久化：GUI 任务同步写回 Excel，供 daemon 读取 ──────────────────


// ── GUI 任务同步：写任务列表到 gui_tasks.json（daemon 会读取）─────────────
ipcMain.handle('save-gui-tasks', async (_, taskList) => {
  try {
    const dir = path.join(os.homedir(), '.wechat-sender')
    fs.mkdirSync(dir, { recursive: true })
    const f = path.join(dir, 'gui_tasks.json')
    fs.writeFileSync(f, JSON.stringify(taskList, null, 2), 'utf8')
    return { ok: true }
  } catch (e) {
    return { ok: false, err: e.message }
  }
})

ipcMain.handle('load-gui-tasks', async () => {
  try {
    const f = path.join(os.homedir(), '.wechat-sender', 'gui_tasks.json')
    if (fs.existsSync(f)) {
      return JSON.parse(fs.readFileSync(f, 'utf8'))
    }
  } catch (_) {}
  return []
})

ipcMain.handle('pick-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const r = await dialog.showOpenDialog(win, {
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  })
  return r.canceled ? null : r.filePaths[0]
})

ipcMain.handle('open-file', async (_, filePath) => {
  await shell.openPath(filePath)
})

ipcMain.handle('get-platform', () => process.platform)

// ─── App lifecycle ────────────────────────────────────────────────────
app.whenReady().then(function() {
  ensurePythonDeps()
  createWindow()
  // 首次启动检测 Automation 权限（macOS）
  checkAutomationPermission()
  app.on('activate', function() {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function() {
  if (daemonProcess) daemonProcess.kill()
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', function() {
  if (daemonProcess) daemonProcess.kill()
})
