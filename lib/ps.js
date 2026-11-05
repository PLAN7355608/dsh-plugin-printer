/**
 * PowerShell 调用层。
 *
 * 约定：ps/*.ps1 的输出一律走 UTF-8（脚本自己设 [Console]::OutputEncoding），
 * 每行是一条记录、字段用 US(\x1f) 分隔；出错时输出 `ERR<US>消息` 并以退出码 1 结束。
 * 脚本参数一律走 -File 的具名参数，值由 argv 传递（CreateProcessW 是 Unicode，
 * 中文路径不会经过代码页转换），因此不需要任何字符串拼接，也就没有注入面。
 *
 * 为什么不用 ctx.subprocess：DSH 的 subprocess 默认把命令交给沙箱执行器（Windows 上是
 * 降为 Low 完整性的受限令牌），而打印后台处理程序、WMI、WIA COM 对 Low 完整性进程一律
 * 拒绝访问。Host 插件运行在宿主进程内、持有普通用户令牌，所以这里用 node:child_process
 * 直接起 powershell.exe。这是本插件唯一绕开 ctx.subprocess 的地方，且仅用于设备操作。
 *
 * 输出捕获优先用管道；若创建匿名管道被环境拒绝（受限沙箱里会得到 spawn EPERM），
 * 自动退回“子进程直接写临时文件”的捕获方式——argv 完全相同，行为一致。
 */
import { spawn } from 'node:child_process'
import { closeSync, openSync, readFileSync, unlinkSync } from 'node:fs'
import { randomUUID } from 'node:crypto'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const PS_DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'ps')
const US = '\x1f'

/**
 * 设备操作失败（消息面向用户，可直接展示）。
 *
 * records：脚本在报错之前已经输出的记录。进纸器可能扫到第 3 页才出错，那时前 2 页已经写盘了，
 *   把这些记录丢掉等于连已经拿到的页一起扔——上层据此可以返回「部分成功」而不是整体失败。
 * code：脚本给的机器可读错误码（输出里单独一行 CODE<US>feederEmpty），用于翻译成人话或分支处理。
 */
export class PrinterError extends Error {
  constructor(message, code) {
    super(message)
    this.name = 'PrinterError'
    this.code = code || ''
    this.records = []
  }
}

function isPipeDenied(error) {
  const code = error && (error.code || error.errno)
  return code === 'EPERM' || code === 'EACCES' || code === 'ENOTSUP' || code === 'UNKNOWN'
}

function buildArgs(scriptName, params) {
  const args = ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', join(PS_DIR, scriptName)]
  for (const [key, value] of Object.entries(params || {})) {
    if (value === undefined || value === null || value === '' || value === false) continue
    args.push('-' + key)
    if (value !== true) args.push(String(value))
  }
  return args
}

function removeFiles(files) {
  if (!files) return
  for (const path of [files.out, files.err]) {
    if (!path) continue
    try { unlinkSync(path) } catch { /* 已经没了 */ }
  }
}

function closeFiles(files) {
  if (!files) return
  for (const fd of [files.fdOut, files.fdErr]) {
    if (fd) { try { closeSync(fd) } catch { /* 已经关了 */ } }
  }
}

function readFileSafe(path) {
  try { return readFileSync(path, 'utf8') } catch { return '' }
}

/**
 * 起一个 PowerShell 进程并等它结束。
 * @returns {Promise<string[][]>} 记录数组（每行按 \x1f 切分后的字段数组）
 */
function attempt(powershell, args, timeoutMs, signal, usePipe) {
  return new Promise((resolve, reject) => {
    let child = null
    let files = null

    try {
      if (usePipe) {
        child = spawn(powershell, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
      } else {
        const stem = join(tmpdir(), 'dsh-printer-' + randomUUID())
        files = { out: stem + '.out', err: stem + '.err', fdOut: 0, fdErr: 0 }
        files.fdOut = openSync(files.out, 'w')
        files.fdErr = openSync(files.err, 'w')
        child = spawn(powershell, args, { windowsHide: true, stdio: ['ignore', files.fdOut, files.fdErr] })
        // 子进程已继承句柄，父进程这边立刻关掉，避免句柄泄漏。
        closeFiles(files)
        files.fdOut = 0
        files.fdErr = 0
      }
    } catch (error) {
      closeFiles(files)
      removeFiles(files)
      reject(error)
      return
    }

    let stdout = ''
    let stderr = ''
    let settled = false
    let timer = null

    const killTree = () => {
      try { child.kill() } catch { /* 已经退出 */ }
      // 超时/取消时把整棵进程树收掉（rundll32、扫描驱动可能挂很久）
      try {
        spawn('taskkill', ['/pid', String(child.pid), '/t', '/f'], { windowsHide: true, stdio: 'ignore' })
      } catch { /* 尽力而为 */ }
    }
    const settle = (fn, value) => {
      if (settled) return
      settled = true
      if (timer) clearTimeout(timer)
      if (signal && signal.removeEventListener) signal.removeEventListener('abort', onAbort)
      removeFiles(files)
      fn(value)
    }
    const onAbort = () => {
      killTree()
      settle(reject, new PrinterError('操作已取消'))
    }

    const finish = (code) => {
      if (settled) return
      if (files) {
        stdout = readFileSafe(files.out)
        stderr = readFileSafe(files.err)
      }
      const records = []
      let errCode = ''
      for (const line of stdout.split(/\r?\n/)) {
        if (!line) continue
        const fields = line.split(US)
        if (fields[0] === 'CODE') { errCode = fields[1] || ''; continue }
        if (fields[0] === 'ERR') {
          const error = new PrinterError(fields.slice(1).join(' ').trim() || '脚本执行失败', errCode)
          // 出错之前的记录要一起交给上层：脚本可能已经成功写完前几页。
          error.records = records
          settle(reject, error)
          return
        }
        records.push(fields)
      }
      if (code !== 0) {
        const detail = stderr.trim().split(/\r?\n/).filter(Boolean).slice(-1)[0] || ('PowerShell 退出码 ' + code)
        const error = new PrinterError(detail, errCode)
        // 非零退出也要带上已经解析出来的记录，交给上层决定算不算「部分成功」。
        error.records = records
        settle(reject, error)
        return
      }
      settle(resolve, records)
    }

    if (signal) {
      if (signal.aborted) { onAbort(); return }
      signal.addEventListener('abort', onAbort, { once: true })
    }
    timer = setTimeout(() => {
      killTree()
      settle(reject, new PrinterError('操作超时（超过 ' + Math.round(timeoutMs / 1000) + ' 秒）'))
    }, timeoutMs)

    if (usePipe) {
      child.stdout.setEncoding('utf8')
      child.stderr.setEncoding('utf8')
      child.stdout.on('data', (chunk) => { stdout += chunk })
      child.stderr.on('data', (chunk) => { stderr += chunk })
    }
    child.on('error', (error) => {
      // 管道被拒（受限沙箱）时把权限类错误交给上层改用临时文件捕获。
      if (usePipe && isPipeDenied(error)) { settle(reject, error); return }
      settle(reject, new PrinterError('无法启动 PowerShell：' + (error && error.message)))
    })
    child.on('close', (code) => finish(code))
  })
}

/**
 * 运行 ps/<name>.ps1。
 * @param {string} name 脚本文件名，如 'list-printers.ps1'
 * @param {Record<string, string|number|boolean|undefined>} params 具名参数：true 表示开关，空值跳过
 * @param {{ powershell?: string, timeoutMs?: number, signal?: AbortSignal }} [options]
 * @returns {Promise<string[][]>} 记录数组
 */
export async function runPs(name, params = {}, options = {}) {
  const powershell = options.powershell || 'powershell.exe'
  const timeoutMs = Number(options.timeoutMs) > 0 ? Number(options.timeoutMs) : 180000
  const signal = options.signal
  const args = buildArgs(name, params)

  try {
    return await attempt(powershell, args, timeoutMs, signal, true)
  } catch (error) {
    if (!isPipeDenied(error)) {
      if (error instanceof PrinterError) throw error
      throw new PrinterError('无法启动 ' + powershell + '：' + (error && error.message))
    }
    return attempt(powershell, args, timeoutMs, signal, false)
  }
}

/** 把记录数组变成对象数组：需要首字段匹配 kind 才算一条。 */
export function pickRecords(records, kind) {
  return records.filter((fields) => fields[0] === kind)
}
