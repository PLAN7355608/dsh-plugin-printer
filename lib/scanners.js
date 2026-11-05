/**
 * 扫描仪（WIA）封装：把 ps/scan*.ps1 的记录行整理成对象，不生成给模型看的文案。
 *
 * 关键约定：scan.ps1 出图是「边扫边报」的 —— 每页落盘就立刻发一条 PAGE 记录。
 * 所以脚本非零退出（进纸器卡纸、驱动中途报错）不代表这一趟白跑：lib/ps.js 会把出错前
 * 已经解析出来的记录挂在 err.records 上，这里据此返回「部分成功」，让调用方如实汇报。
 */
import { runPs, pickRecords, PrinterError } from './ps.js'
import { translateError } from './errors.js'

const DEVICE_TYPE_TEXT = {
  1: '扫描仪',
  2: '相机',
  3: '摄像机',
  65535: '未知类型（eSCL 等新式驱动常见）',
}

function toNumber(value) {
  if (value === undefined || value === null || value === '') return null
  const n = Number(value)
  return Number.isFinite(n) ? n : null
}

function splitList(value) {
  if (!value) return []
  return String(value)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
}

function numberList(value) {
  return splitList(value)
    .map((s) => toNumber(s))
    .filter((n) => n !== null)
}

/** 文档源能力位（WIA 3086）。 */
export function decodeCaps(raw) {
  const caps = toNumber(raw)
  if (caps === null) return { raw: null, feeder: null, flatbed: null, duplex: null }
  return {
    raw: caps,
    feeder: (caps & 1) === 1,
    flatbed: (caps & 2) === 2,
    duplex: (caps & 4) === 4,
  }
}

function parsePage(fields) {
  return {
    index: toNumber(fields[1]) || 0,
    path: fields[2] || '',
    bytes: toNumber(fields[3]) || 0,
    format: fields[4] || '',
    width: toNumber(fields[5]) || 0,
    height: toNumber(fields[6]) || 0,
    dpi: toNumber(fields[7]) || 0,
    converted: fields[8] === '1',
    kept: fields[9] === undefined ? true : fields[9] === '1',
  }
}

function parseWarnings(records) {
  return pickRecords(records, 'WARN').map((f) => ({ kind: f[1] || '', message: f[2] || '' }))
}

/** 列出 WIA 扫描设备。WIA 打不开时抛 scannerUnreadable，不会拿空列表冒充「没有扫描仪」。 */
export async function listScanners(options) {
  let records
  try {
    records = await runPs('list-scanners.ps1', {}, options)
  } catch (error) {
    throw translateError(error)
  }
  const scanners = []
  const warnings = []
  for (const fields of records) {
    if (fields[0] !== 'SCANNER') continue
    const type = toNumber(fields[2])
    // available === null means the script did not probe this device; false means it was listed
    // but could not be opened right now (the usual stale eSCL entry).
    const available = fields[4] === '1' ? true : fields[4] === '0' ? false : null
    const name = fields[1] || ''
    const error = fields[5] || ''
    if (available === false) warnings.push('「' + name + '」现在打不开：' + (error || '设备没有响应') + '（扫描时会自动跳过它，选别的条目）。')
    scanners.push({
      name,
      type: fields[2] || '',
      typeName: type === null ? '' : DEVICE_TYPE_TEXT[type] || ('类型 ' + type),
      id: fields[3] || '',
      available,
      error,
    })
  }
  return { scanners, warnings }
}

/** 查询扫描仪能力（不扫描）。device 为空时返回全部设备。 */
export async function scannerInfo(params, options) {
  let records
  try {
    records = await runPs('scanner-info.ps1', { Device: params.device }, options)
  } catch (error) {
    throw translateError(error)
  }
  const warned = pickRecords(records, 'NONE')[0]
  const devices = []
  for (const fields of records) {
    if (fields[0] !== 'DETAIL') continue
    const name = fields[1] || ''
    const caps = decodeCaps(fields[2])
    const formats = splitList(fields[3])
    const resolutions = numberList(fields[4])
    const sourceRaw = toNumber(fields[5])
    devices.push({
      name,
      type: '',
      id: '',
      capabilities: caps,
      hasFeeder: caps.feeder,
      hasFlatbed: caps.flatbed,
      hasDuplex: caps.duplex,
      supportedFormats: formats,
      supportedDpi: resolutions,
      currentSource: sourceRaw,
      currentSourceText: sourceRaw === 1 ? 'feeder' : sourceRaw === 2 ? 'flatbed' : sourceRaw === null ? '' : 'unknown',
      documentStatus: toNumber(fields[6]),
      formatListPublished: formats.length > 0,
      note: fields[7] || '',
    })
  }
  for (const fields of records) {
    if (fields[0] !== 'SCANNER') continue
    const type = toNumber(fields[2])
    for (const d of devices) {
      if (d.name !== (fields[1] || '')) continue
      d.type = fields[2] || ''
      d.typeName = type === null ? '' : DEVICE_TYPE_TEXT[type] || ('类型 ' + type)
      d.id = fields[3] || ''
    }
  }
  return { devices, warnings: warned ? [warned[1] || warned[0] || ''] : [] }
}

/**
 * 扫描到图片或 PDF。返回结构：
 *   { device, pages, requestedDpi, effectiveDpi, source, supportedFormats, supportedDpi,
 *     stoppedBy, nativeFormat, crop, pdfPath, pdfBytes, warnings, failure }
 * failure 非空表示「扫到了页面但这一趟中途报错」——pages 仍然可用。
 */
export async function scan(params, options) {
  const args = {
    Device: params.device,
    Output: params.output,
    Dpi: params.dpi,
    Mode: params.mode,
    Format: params.format,
    Source: params.source,
    Crop: params.crop,
    KeepPages: params.keepPages ? true : undefined,
  }
  let records = []
  let failure = null
  try {
    records = await runPs('scan.ps1', args, options)
  } catch (error) {
    if (!(error instanceof PrinterError) || !Array.isArray(error.records)) throw translateError(error)
    records = error.records
    // 中文提示在这里生成；STAT 行说明失败前已经完成了几页（部分成功）。
    const translated = translateError(error)
    const stat = pickRecords(records, 'STAT')[0] || []
    failure = {
      code: error.code || 'scanFailed',
      message: translated && translated.message ? String(translated.message) : String(error.message || error),
      completedPages: toNumber(stat[1]) || 0,
      stoppedBy: stat[2] || '',
      // The OK record is only emitted on a clean finish, so the failure path repeats the device and
      // the source on the STAT line - otherwise a partial scan names no device at all.
      device: stat[3] || '',
      source: stat[4] || '',
    }
  }

  const pages = pickRecords(records, 'PAGE').map(parsePage)
  if (pages.length === 0) {
    if (failure) throw new PrinterError(failure.message, failure.code)
    throw new PrinterError('扫描仪没有返回任何页面。', 'noPage')
  }

  const ok = pickRecords(records, 'OK')[0] || []
  const pdf = pickRecords(records, 'PDF')[0] || []
  const caps = decodeCaps(ok[10])
  const result = {
    device: ok[2] || (failure ? failure.device : '') || '',
    pages,
    requestedDpi: toNumber(ok[4]),
    effectiveDpi: toNumber(ok[5]),
    source: ok[6] || (failure ? failure.source : '') || '',
    supportedFormats: splitList(ok[7]),
    supportedDpi: numberList(ok[8]),
    stoppedBy: ok[9] || 'complete',
    capabilities: caps,
    nativeFormat: ok[11] || '',
    crop: ok[12] || 'off',
    fallback: ok[13] || '',
    pdfPath: pdf[1] || '',
    pdfBytes: toNumber(pdf[2]) || 0,
    pdfPages: toNumber(pdf[3]) || 0,
    warnings: parseWarnings(records),
    failure,
  }
  return result
}
