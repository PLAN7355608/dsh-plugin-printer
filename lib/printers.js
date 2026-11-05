/**
 * 打印机与打印队列的高层封装：把 ps/*.ps1 的记录行整理成对象。
 * 这里只做数据，不生成给模型看的文案（文案在 index.js）。
 */
import { runPs, pickRecords, PrinterError } from './ps.js'
import { translateError } from './errors.js'

function toNumber(value) {
  if (value === undefined || value === null || value === '') return null
  const n = Number(value)
  return Number.isFinite(n) ? n : null
}

const translateQueueError = translateError
const translatePrintError = translateError

function boolValue(value) {
  if (value === undefined || value === null || value === '') return null
  return value === '1' || String(value).toLowerCase() === 'true'
}

/** 列出打印机 + 当前默认打印机。 */
export async function listPrinters(options) {
  let records
  try {
    records = await runPs('list-printers.ps1', {}, options)
  } catch (error) {
    // 两个来源都读不到时脚本会报错，而不是拿空列表冒充「没有打印机」。
    throw translateQueueError(error)
  }
  const warnings = pickRecords(records, 'WARN').map((f) => ({
    kind: f[1] || '',
    message: f[3] ? String(f[2] || '') + '：' + f[3] : String(f[2] || ''),
  }))
  const printers = []
  let defaultPrinter = ''
  for (const fields of records) {
    if (fields[0] === 'DEFAULT') { defaultPrinter = fields[1] || ''; continue }
    if (fields[0] !== 'PRINTER') continue
    printers.push({
      name: fields[1] || '',
      driver: fields[2] || '',
      port: fields[3] || '',
      status: fields[4] || '',
      jobs: toNumber(fields[5]),
      shared: fields[6] === 'True',
      note: fields[7] || '',
      source: fields[8] || '',
    })
  }
  if (!defaultPrinter) {
    const hit = printers.find((p) => String(p.note).toLowerCase() === 'default')
    if (hit) defaultPrinter = hit.name
  }
  for (const p of printers) p.default = p.name === defaultPrinter
  return { defaultPrinter, printers, warnings }
}

/**
 * 打印机能力查询：基础信息来自 listPrinters()，硬件能力来自 printer-caps.ps1。
 * 找不到打印机时抛 PrinterError（带 available 名单，便于调用方给出可选值）。
 */
export async function printerInfo(params, options) {
  const { defaultPrinter, printers } = await listPrinters(options)
  if (printers.length === 0) throw new PrinterError('这台电脑上没有安装任何打印机。', 'noPrinter')
  const wanted = String(params.printer || '').trim()
  let target = null
  if (wanted) {
    const lower = wanted.toLowerCase()
    target = printers.find((p) => p.name.toLowerCase() === lower)
      || printers.find((p) => p.name.toLowerCase().includes(lower))
    if (!target) {
      const error = new PrinterError('没有找到打印机「' + wanted + '」。当前可用：' + printers.map((p) => p.name).join('、'), 'printerNotFound')
      error.available = printers.map((p) => p.name)
      throw error
    }
  } else {
    target = printers.find((p) => p.default) || printers.find((p) => p.name === defaultPrinter) || printers[0]
  }

  const capabilities = {
    canDuplex: null,
    duplexOn: null,
    color: null,
    horizontalResolution: null,
    verticalResolution: null,
    paperSizes: [],
    comment: '',
    location: '',
    duplexSource: '',
    colorSource: '',
  }
  let config = null
  let configError = ''
  const warnings = []
  try {
    const records = await runPs('printer-caps.ps1', { Printer: target.name }, options)
    const caps = pickRecords(records, 'CAPS')[0]
    if (caps) {
      capabilities.canDuplex = boolValue(caps[2])
      capabilities.duplexOn = boolValue(caps[3])
      capabilities.color = boolValue(caps[4])
      capabilities.horizontalResolution = toNumber(caps[5])
      capabilities.verticalResolution = toNumber(caps[6])
      capabilities.paperSizes = (caps[7] || '').split(',').map((s) => toNumber(s)).filter((n) => n !== null)
      capabilities.comment = caps[8] || ''
      capabilities.location = caps[9] || ''
      // wmi | capabilities | gdi - where canDuplex / color came from. A value inferred from a
      // capability bit or from the GDI driver is still true, but it is not a fact the WMI
      // properties reported, and the host message says so instead of overstating it.
      capabilities.duplexSource = caps[10] || ''
      capabilities.colorSource = caps[11] || ''
    }
    const conf = pickRecords(records, 'CONFIG')[0]
    if (conf) config = { duplexing: conf[1] || '', color: conf[2] || '', paperSize: conf[3] || '' }
    for (const w of pickRecords(records, 'WARN')) {
      if (w[1] === 'config') configError = w[2] || ''
      else warnings.push(w[2] || '')
    }
  } catch (error) {
    // 查询被拒（权限、WMI 不可用）不等于打印机没有能力：如实报错误，能力字段保持 null。
    const translated = translateError(error)
    configError = translated && translated.message ? String(translated.message) : String(error)
    capabilities.unreadable = true
  }

  return { defaultPrinter, printer: target, capabilities, config, configError, warnings }
}

/** 设为系统默认打印机。返回实测结果：脚本会回读注册表，改不动时给 WARN 而不是假成功。 */
export async function setDefaultPrinter(printer, options) {
  const records = await runPs('printer-action.ps1', { Action: 'set-default', Printer: printer }, options)
  const ok = records.find((f) => f[0] === 'OK')
  if (ok) {
    return { printer, actual: ok[3] || printer, verified: true, method: ok[4] || '' }
  }
  const warn = records.find((f) => f[0] === 'WARN')
  if (warn) {
    return { printer, actual: warn[3] || '', verified: false, method: '' }
  }
  return { printer, actual: '', verified: false, method: '' }
}

/** 打印 Windows 测试页（会真的出纸）。返回 Win32 返回码、设备状态与队列状态，便于如实汇报。 */
export async function printTestPage(printer, options) {
  const records = await runPs('printer-action.ps1', { Action: 'test-page', Printer: printer }, options)
  const ok = records.find((f) => f[0] === 'OK')
  if (ok) {
    return {
      printer,
      unverified: false,
      returnValue: ok[3] || '',
      state: ok[4] || '',
      jobStatus: ok[5] || '',
      detail: ok[3] || '',
    }
  }
  const warn = records.find((f) => f[0] === 'WARN')
  if (warn) {
    return { printer, unverified: true, returnValue: '', state: warn[4] || '', jobStatus: '', detail: warn[3] || '' }
  }
  return { printer, unverified: true, returnValue: '', state: '', jobStatus: '', detail: '' }
}

/**
 * 打印一个文件。kind 取值与含义：
 *   queue         文本走后台队列（Out-Printer）
 *   printto       系统 printto 动词，程序认打印机参数
 *   defaultSwitch 临时把目标机设为系统默认机再打印（allowDefaultSwitch），已还原
 *   file          打印到文件（如 Microsoft Print to PDF）
 *   unknown       脚本没给出可识别的成功记录
 */
export async function printFile(params, options) {
  let records
  try {
    records = await runPs('print-file.ps1', {
      Path: params.path,
      Printer: params.printer,
      Copies: params.copies,
      Mode: params.mode,
      PageRange: params.pageRange,
      Collate: params.collate,
      OutputFile: params.outputFile,
      TextEncoding: params.encoding,
      Duplex: params.duplex,
      AllowDefaultSwitch: params.allowDefaultSwitch ? true : undefined,
      AllowHugeJob: params.allowHugeJob ? true : undefined,
      MaxPages: params.maxPages,
    }, options)
  } catch (error) {
    throw translatePrintError(error)
  }
  const warnings = pickRecords(records, 'WARN').map((f) => ({ kind: f[1] || '', message: f[2] || '' }))
  for (const fields of records) {
    if (fields[0] !== 'OK') continue
    const kind = fields[1]
    if (kind === 'gdi') {
      return {
        kind: 'gdi',
        printer: fields[2] || '',
        file: fields[3] || '',
        pagesTotal: toNumber(fields[4]) || 0,
        pagesSelected: fields[5] || 'all',
        copies: toNumber(fields[6]) || 1,
        collate: fields[7] !== 'false',
        duplex: fields[8] || 'off',
        observed: fields[9] || '',
        sourceKind: fields[10] || '',
        outputFile: fields[11] || '',
        warnings,
      }
    }
    if (kind === 'queue') {
      return {
        kind: 'queue',
        printer: fields[2] || '',
        copies: toNumber(fields[3]) || 1,
        file: fields[4] || '',
        duplex: fields[5] || 'off',
        observed: fields[6] || '',
        warnings,
      }
    }
    if (kind === 'printto') {
      return {
        kind: 'printto',
        printer: fields[2] || '',
        file: fields[3] || '',
        observed: fields[4] || '',
        warnings,
      }
    }
    if (kind === 'defaultSwitch') {
      return {
        kind: 'defaultSwitch',
        printer: fields[2] || '',
        file: fields[3] || '',
        restored: fields[4] || '',
        observed: fields[5] || '',
        duplex: fields[6] || 'off',
        warnings,
      }
    }
    if (kind === 'file') {
      return {
        kind: 'file',
        printer: fields[2] || '',
        outputFile: fields[3] || '',
        bytes: toNumber(fields[4]) || 0,
        relocatedFrom: fields[5] || '',
        method: fields[6] || '',
        warnings,
      }
    }
  }
  return { kind: 'unknown', warnings }
}

/** 打印队列。 */
export async function listJobs(printer, options) {
  let records
  try {
    records = await runPs('printer-jobs.ps1', { Action: 'list', Printer: printer }, options)
  } catch (error) {
    throw translateQueueError(error)
  }
  const jobs = []
  for (const fields of records) {
    if (fields[0] !== 'JOB') continue
    jobs.push({
      printer: fields[1] || '',
      id: fields[2] || '',
      document: fields[3] || '',
      status: fields[4] || '',
      pages: toNumber(fields[5]),
      owner: fields[6] || '',
    })
  }
  const warnings = pickRecords(records, 'WARN').map((f) => ({ kind: f[1] || '', message: f[3] || f[2] || '' }))
  return { jobs, warnings }
}

/**
 * 取消一个作业。脚本会先删再回读，所以 cancelled 表示「确认已从队列消失」；
 * 删除请求发出但作业还挂着时脚本会报 stillQueued，不会假报成功。
 */
export async function cancelJob(printer, jobId, options) {
  let records
  try {
    records = await runPs('printer-jobs.ps1', { Action: 'cancel', Printer: printer, JobId: jobId }, options)
  } catch (error) {
    throw translateQueueError(error)
  }
  const ok = records.find((f) => f[0] === 'OK' && f[1] === 'cancel') || []
  return {
    printer,
    jobId,
    cancelled: ok.length > 0,
    readback: ok[4] || '',
    status: ok[5] || '',
  }
}

/** 清空一台打印机的队列。requested/deleted/remaining 都来自脚本的回读复核。 */
export async function purgeJobs(printer, options) {
  let records
  try {
    records = await runPs('printer-jobs.ps1', { Action: 'purge', Printer: printer }, options)
  } catch (error) {
    throw translateQueueError(error)
  }
  const progress = records.find((f) => f[0] === 'PROGRESS' && f[1] === 'purge') || []
  const ok = records.find((f) => f[0] === 'OK' && f[1] === 'purge') || []
  const requested = ok[3] !== undefined ? toNumber(ok[3]) : toNumber(progress[3])
  const deleted = ok[4] !== undefined ? toNumber(ok[4]) : toNumber(progress[4])
  return {
    printer,
    requested: requested === null ? 0 : requested,
    deleted: deleted === null ? 0 : deleted,
    failures: toNumber(progress[5]) || 0,
    remaining: toNumber(ok[5]) || 0,
    readback: ok[6] || '',
    verified: ok.length > 0,
  }
}
