/**
 * 打印机与扫描仪插件（Host 半边）。
 *
 * 给智能体一套工具：枚举打印机与 WIA 扫描仪、打印文件或测试页、管理打印队列、
 * 扫描到图片或 PDF。真正的 Windows 调用在 ps/*.ps1，本文件只负责参数校验与给模型看的文案。
 *
 * 为什么不用 ctx.subprocess：DSH 的 subprocess 会把命令交给沙箱执行器（Windows 上降为 Low
 * 完整性令牌），打印后台处理程序 / WMI / WIA COM 对 Low 完整性一律拒绝访问。Host 插件运行在
 * 宿主进程内、持有普通用户令牌，因此这里用 node:child_process 直接起 powershell.exe。
 */
import { isAbsolute, resolve } from 'node:path'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { PrinterError } from './lib/ps.js'
import {
  listPrinters,
  printerInfo,
  setDefaultPrinter,
  printTestPage,
  printFile,
  listJobs,
  cancelJob,
  purgeJobs,
} from './lib/printers.js'
import { listScanners, scannerInfo, scan } from './lib/scanners.js'

export const name = 'dsh-plugin-printer'
export const inject = ['tools']

const DEFAULTS = {
  powershell: 'powershell.exe',
  timeoutMs: 180000,
}

const TEXT_OUTPUT = {
  schema: { type: 'json' },
  render: (_args, value) => [{ type: 'text', text: String(value && value.message ? value.message : '') }],
}

function formatBytes(bytes) {
  const n = Number(bytes)
  if (!Number.isFinite(n) || n <= 0) return '0 字节'
  if (n < 1024) return n + ' 字节'
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB'
  if (n < 1024 * 1024 * 1024) return (n / 1024 / 1024).toFixed(2) + ' MB'
  return (n / 1024 / 1024 / 1024).toFixed(2) + ' GB'
}

/** ps 脚本给出的状态标记（如 Idle+LowToner）翻成中文；不认识的标记原样保留。 */
const STATUS_TEXT = {
  Idle: '空闲',
  Printing: '正在打印',
  Warmup: '预热中',
  Stopped: '已停止',
  Offline: '脱机',
  Error: '出错',
  PaperJam: '卡纸',
  PaperOut: '缺纸',
  LowPaper: '纸张不足',
  NoPaper: '无纸',
  ManualFeed: '手动进纸',
  PaperProblem: '纸张问题',
  OutputBinFull: '出纸盒已满',
  NotAvailable: '不可用',
  LowToner: '墨粉不足',
  NoToner: '墨粉用尽',
  UserIntervention: '需要人工干预',
  OutOfMemory: '打印机内存不足',
  DoorOpen: '门未关',
  ServerUnknown: '服务器未知',
  ServiceRequested: '需要维修',
  CannotPrintPage: '无法打印该页',
}

function translateStatus(status) {
  if (!status) return ''
  return String(status)
    .split('+')
    .map((token) => STATUS_TEXT[token] || token)
    .join('，')
}

function fail(error) {
  let message = error instanceof PrinterError
    ? error.message
    : (error && error.message) || String(error)
  // 打印队列、WMI、WIA COM 都要求普通用户令牌：受限权限下 Windows 只回一句
  // "Access is denied."，对用户没有任何指导意义，这里补一句人话。
  if (/access(\s+is)?\s+denied|拒绝访问|0x80070005|80070005/i.test(message) && message.indexOf('权限') === -1) {
    message += '（Windows 拒绝了这次操作：打印队列、WMI 和扫描仪 COM 都需要普通用户权限，权限受限时会失败。）'
  }
  return { ok: false, message: message }
}

function oneOf(value, allowed, label) {
  if (value === undefined || value === null || value === '') return undefined
  const text = String(value).toLowerCase()
  if (allowed.indexOf(text) === -1) {
    throw new PrinterError(label + ' 只能是 ' + allowed.join(' / ') + '，收到：' + value)
  }
  return text
}

/** 把「提交后有没有在队列里看见作业」翻成一句实话。 */
function describeObservation(observed, printer) {
  if (observed === 'submitted') return '已在「' + printer + '」的队列里看到这次作业。'
  if (observed === 'unreadable') return '当前进程读不到打印队列，无法确认作业是否真的进了「' + printer + '」。'
  return '提交后 2.5 秒内没有在队列里看到作业：作业可能已经打完并离开队列（常见的作业在队列里只停留不到一秒），'
    + '也可能根本没进去。要确认请用 printer_jobs 再看一次，或直接看打印机面板。'
}

/** 把双面取值翻成人话。 */
function duplexText(duplex) {
  if (duplex === 'long') return '双面长边翻页'
  if (duplex === 'short') return '双面短边翻页'
  return '单面'
}

function warningLines(warnings) {
  const lines = []
  for (const w of warnings || []) {
    const kind = typeof w === 'string' ? '' : (w && w.kind) || ''
    const message = typeof w === 'string' ? w : (w && w.message) || ''
    if (!message) continue
    if (kind === 'duplex' || kind === 'default' || kind === 'restore' || kind === 'printer') {
      lines.push('注意：' + message)
    } else {
      lines.push('提示：' + message)
    }
  }
  return lines
}

/**
 * @param {import('@deepseek-ai/dsh-tools').Context} ctx
 * @param {{ powershell?: string, timeoutMs?: number }} [config]
 */
export function apply(ctx, config) {
  const settings = Object.assign({}, DEFAULTS, config || {})

  const runOptions = (exec) => ({
    powershell: settings.powershell,
    timeoutMs: settings.timeoutMs,
    signal: exec && exec.signal,
  })

  const sessionCwd = (exec) => {
    const cwd = exec && exec.agent && exec.agent.session && exec.agent.session.header
      ? exec.agent.session.header.cwd
      : ''
    return cwd || process.cwd()
  }

  const resolvePath = (exec, path) => (isAbsolute(path) ? path : resolve(sessionCwd(exec), path))

  ctx.tools.register(defineTool({
    name: 'printer_list',
    description: '列出这台电脑上的打印机：名称、驱动、端口、状态、队列里排了几个作业、是否共享，以及哪一台是系统默认打印机。只读操作，不会打印任何东西。',
    parameters: {},
    output: TEXT_OUTPUT,
    async execute(_args, exec) {
      try {
        const result = await listPrinters(runOptions(exec))
        const lines = []
        if (result.printers.length === 0) {
          lines.push('没有找到任何打印机。可能是打印后台处理程序（Spooler 服务）没启动，或这台电脑没装打印机。')
        } else {
          const head = ['共 ' + result.printers.length + ' 台打印机']
          if (result.defaultPrinter) head.push('默认：' + result.defaultPrinter)
          lines.push(head.join('；') + '：')
          result.printers.forEach((p, index) => {
            lines.push('')
            lines.push((index + 1) + '. ' + p.name + (p.default ? '（默认）' : ''))
            if (p.driver) lines.push('   驱动：' + p.driver)
            if (p.port) lines.push('   端口：' + p.port)
            const bits = []
            if (p.status) bits.push('状态：' + translateStatus(p.status))
            if (p.jobs !== null) bits.push('队列：' + p.jobs + ' 个作业')
            if (p.shared) bits.push('已共享')
            if (bits.length) lines.push('   ' + bits.join('，'))
          })
          lines.push('')
          lines.push('提示：要知道某台机器支不支持双面、彩色、分辨率，用 printer_info。')
        }
        lines.push(...warningLines(result.warnings))
        return { ok: true, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'printer_info',
    description: '查询一台打印机的详细信息与硬件能力：驱动、端口、状态、是不是默认机、支不支持双面/彩色、分辨率、纸张规格，以及当前的打印偏好（双面模式）。只读操作。不填 printer 时查系统默认打印机。',
    parameters: {
      printer: { type: 'string', description: '打印机名称，可只写一部分；不填则查系统默认打印机。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const result = await printerInfo({ printer: args.printer ? String(args.printer) : '' }, runOptions(exec))
        const p = result.printer
        const c = result.capabilities
        const lines = ['「' + p.name + '」' + (p.default ? '（系统默认打印机）' : '')]
        if (p.driver) lines.push('驱动：' + p.driver)
        if (p.port) lines.push('端口：' + p.port)
        if (p.status) lines.push('状态：' + translateStatus(p.status))
        if (p.note) lines.push('备注：' + p.note)
        const caps = []
        const inferred = (source) => {
          if (source === 'capabilities') return '（由驱动能力位推断）'
          if (source === 'gdi') return '（由打印驱动接口推断）'
          return ''
        }
        if (c.canDuplex === true) caps.push('双面：支持' + inferred(c.duplexSource))
        else if (c.canDuplex === false) caps.push('双面：不支持' + inferred(c.duplexSource))
        if (c.color === true) caps.push('彩色：支持' + inferred(c.colorSource))
        else if (c.color === false) caps.push('彩色：不支持（单色机）' + inferred(c.colorSource))
        if (c.horizontalResolution && c.verticalResolution) {
          caps.push('分辨率：' + c.horizontalResolution + '×' + c.verticalResolution + ' dpi')
        }
        if (caps.length) lines.push(caps.join('；'))
        // A property the driver simply does not expose is not the same as a refused query: say
        // which one it is instead of leaving an unexplained gap in the capability line.
        if (!c.unreadable && (c.canDuplex === null || c.color === null)) {
          lines.push('注意：这个驱动没有报告' + (c.canDuplex === null ? '是否支持双面' : '') + (c.canDuplex === null && c.color === null ? '、' : '') + (c.color === null ? '是否支持彩色' : '') + '；留空不代表不支持。')
        }
        if (result.config) {
          const bits = []
          if (result.config.duplexing) bits.push('双面模式：' + result.config.duplexing)
          if (result.config.color) bits.push('彩色：' + result.config.color)
          if (result.config.paperSize) bits.push('纸张：' + result.config.paperSize)
          if (bits.length) lines.push('当前打印偏好：' + bits.join('，'))
        }
        if (c.paperSizes && c.paperSizes.length) lines.push('支持的纸张规格编号（Win32）：' + c.paperSizes.join('、'))
        if (c.location) lines.push('位置：' + c.location)
        if (c.comment) lines.push('注释：' + c.comment)
        if (result.configError) {
          lines.push('')
          lines.push('没能读到打印偏好（双面模式等）：' + result.configError)
        }
        if (c.unreadable) {
          lines.push('')
          lines.push('注意：这台打印机的硬件能力（双面/彩色/分辨率/纸张）没有查到，这一栏为空不代表不支持，多半是查询被系统拒绝。换一个有权限的环境再试，或者直接用 printer_print 打测试页看实际效果。')
        }
        lines.push(...warningLines(result.warnings))
        if (c.canDuplex === true) {
          lines.push('')
          lines.push('提示：printer_print 的 duplex 参数可以按次改双面（long=长边翻页，short=短边翻页），打完自动还原原来的设置。')
        }
        return { ok: true, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'printer_print',
    description: '把一个文件送到打印机打印。文本和图片（txt/log/csv/md/json…、jpg/png/bmp/tif…）默认走插件自己的 GDI 渲染路径：页码范围、份数、逐份/逐页顺序都能精确控制，双面方向随每个作业下发、不改打印机的默认设置。PDF/OOXML/Office 文件走系统「打印」动词，那条路管不了页码与份数顺序。若该文件类型没有 printto 动词、或关联程序会忽略打印机参数（WPS Office 就是），插件会直接拒绝并说明原因，而不是悄悄打到默认打印机——除非显式传 allowDefaultSwitch。给出 outputFile 还能打印到文件（配合 Microsoft Print to PDF 可无纸验证整条打印链路）。gdi 模式只接受文本与图片：文件内容会被嗅探，二进制文件（PDF、Office、压缩包等）一律拒绝；算出来超过 200 页的作业也会先拒绝，避免误判把纸盘打空（确认无误可用 allowHugeJob 放行）。',
    parameters: {
      path: { type: 'string', required: true, description: '要打印的文件路径：绝对路径，或相对当前会话工作目录的路径。' },
      printer: { type: 'string', description: '目标打印机名称；不填用系统默认打印机（打印到文件时默认用 Microsoft Print to PDF）。名称必须先在 printer_list 里确认。' },
      copies: { type: 'integer', description: '打印份数，默认 1，最多 99。和 collate 一起决定多份的顺序。' },
      pageRange: { type: 'string', description: '页码范围：默认 all（全部）。支持 1-3、5、8-（到末页）、-4（首页到第 4 页）、odd（奇数页）、even（偶数页），也可组合成 1-3,5,8-。只对 auto/gdi 模式下的文本与图片生效；PDF/OOXML 不支持，传了会被明确拒绝。' },
      collate: { type: 'boolean', description: '多份时的顺序：true（默认）逐份打印（1,2,3,1,2,3，一份打完再打下一份），false 逐页打印（1,1,2,2,3,3，每页打满份数再换下一页）。只对 auto/gdi 模式下的文本与图片生效。' },
      mode: { type: 'string', description: 'auto（默认）/ gdi / queue / verb。auto：有 outputFile 就走打印到文件、文本与图片走 gdi、其它走 verb；gdi：插件自己渲染页面后打印（页码/份数/逐份/双面都随作业走，最可控）；queue：把文本重排成 UTF-8 后直接送后台打印队列，不弹窗、不依赖文件关联；verb：调用文件类型的系统「打印」动词。' },
      duplex: { type: 'string', description: '双面：off（默认）/ long（长边翻页，常用）/ short（短边翻页）。gdi 模式下随作业下发、不改打印机的设置；queue/verb 模式下是临时改这台打印机的双面偏好、打完还原；对打印到文件无效。' },
      allowDefaultSwitch: { type: 'boolean', description: '默认 false。设为 true 时，如果这个文件类型没有可用的 printto 动词（或关联程序忽略打印机参数），插件会临时把系统默认打印机切成目标机、打印、再还原。仅在你接受「临时改动系统默认打印机」时使用。' },
      allowHugeJob: { type: 'boolean', description: '默认 false。gdi 模式对算出来超过页数上限（默认 200 页）的作业会先拒绝（防止文件判断错误时把纸盘打空）；确认文件无误后设为 true 放行。' },
      maxPages: { type: 'integer', description: '页数熔断上限，默认 200。gdi 渲染出来的作业总页数（含份数）超过这个数就直接拒绝，不发给打印机。' },
      outputFile: { type: 'string', description: '可选。打印到文件的输出路径（自动切到打印到文件模式）。例如配合 printer="Microsoft Print to PDF" 生成 PDF，用来无纸验证。注意 PDF 打印机会自己弹保存框、不认这个路径，插件会把驱动实际写出的文件移到这里。' },
      encoding: { type: 'string', description: 'queue 模式读取文本的编码：auto（默认，按 BOM 判断，不是 UTF-8 就按 GBK 读）/ utf8 / gbk / unicode。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const path = String(args.path || '').trim()
        if (!path) throw new PrinterError('必须给出要打印的文件路径。')
        const mode = oneOf(args.mode, ['auto', 'gdi', 'queue', 'verb'], 'mode')
        const encoding = oneOf(args.encoding, ['auto', 'utf8', 'gbk', 'unicode'], 'encoding')
        const duplex = oneOf(args.duplex, ['off', 'long', 'short'], 'duplex')
        const pageRange = args.pageRange ? String(args.pageRange).trim() : ''
        const collate = args.collate === false ? 'false' : 'true'
        let copies = args.copies === undefined || args.copies === null ? 1 : Number(args.copies)
        if (!Number.isFinite(copies) || copies < 1) copies = 1
        if (copies > 99) throw new PrinterError('打印份数最多 99 份。')
        let maxPages
        if (args.maxPages !== undefined && args.maxPages !== null && args.maxPages !== '') {
          maxPages = Number(args.maxPages)
          if (!Number.isFinite(maxPages) || maxPages < 1) throw new PrinterError('maxPages 必须是大于 0 的整数。')
        }

        const full = resolvePath(exec, path)
        const outputFile = args.outputFile ? resolvePath(exec, String(args.outputFile)) : ''
        const result = await printFile({
          path: full,
          printer: args.printer ? String(args.printer) : '',
          copies,
          mode: mode || 'auto',
          pageRange,
          collate,
          outputFile,
          encoding: encoding || 'auto',
          duplex: duplex || 'off',
          allowDefaultSwitch: args.allowDefaultSwitch === true,
          allowHugeJob: args.allowHugeJob === true,
          maxPages,
        }, runOptions(exec))

        const lines = []
        let ok = true

        if (result.kind === 'gdi') {
          const sel = String(result.pagesSelected)
          const pageText = result.pagesSelected === 'all'
            ? '全部 ' + result.pagesTotal + ' 页'
            : (sel === 'odd' || sel === 'even')
              ? (sel === 'odd' ? '奇数页' : '偶数页') + '（这份文件共 ' + result.pagesTotal + ' 页）'
              : '第 ' + sel.replace(/,/g, '、') + ' 页（这份文件共 ' + result.pagesTotal + ' 页）'
          const orderText = result.copies > 1
            ? result.copies + ' 份，' + (result.collate ? '逐份（一份打完再打下一份）' : '逐页（每页打满份数再换下一页）')
            : '1 份'
          const bits = [pageText, orderText, duplexText(result.duplex)]
          if (result.sourceKind === 'image') bits.push('图片模式')
          lines.push('已把 ' + full + ' 交给「' + result.printer + '」打印：' + bits.join('；') + '。')
          if (result.outputFile) lines.push('这次打印到了文件：' + result.outputFile)
          lines.push('页面是插件自己渲染的：页码、份数顺序和双面都随这一个作业下发，没有改动打印机的默认设置。')
          lines.push(describeObservation(result.observed, result.printer))
        } else if (result.kind === 'queue') {
          lines.push('已把 ' + full + ' 送进打印机「' + result.printer + '」的队列，共 ' + result.copies + ' 份。')
          if (result.duplex && result.duplex !== 'off') lines.push('这次按双面打印（' + result.duplex + '），已还原原来的双面设置。')
          lines.push(describeObservation(result.observed, result.printer))
          lines.push('这是后台打印，命令返回不代表已经出纸；用 printer_jobs 可以看队列状态。')
        } else if (result.kind === 'printto') {
          lines.push('已通过「' + result.printer + '」把 ' + full + ' 交给系统的 printto 动词处理（打印程序已退出）。')
          lines.push(describeObservation(result.observed, result.printer))
          if (result.observed === 'unreadable') ok = false
        } else if (result.kind === 'defaultSwitch') {
          lines.push('这个文件类型不会把打印机参数传给关联程序，所以插件临时把系统默认打印机换成了「' + result.printer + '」，打印完再还原。')
          lines.push('默认打印机还原：' + (result.restored === 'ok' ? '已还原' : result.restored === 'failed' ? '还原失败，请手动检查系统默认打印机' : '无需还原（本来就是这个）') + '。')
          if (result.duplex && result.duplex !== 'off') lines.push('这次按双面打印（' + result.duplex + '），已还原原来的双面设置。')
          lines.push(describeObservation(result.observed, result.printer))
          if (result.observed !== 'submitted') {
            lines.push('这条路依赖关联程序自己弹界面/打印，无法保证它真的打印；如果没出纸，请改用 printer_test_page 验证机器，或把文件转成 PDF/文本再打印。')
          }
          if (result.observed === 'unreadable' || result.restored === 'failed') ok = false
        } else if (result.kind === 'file') {
          let extra = ''
          if (result.relocatedFrom) {
            extra = '\n注意：「Microsoft Print to PDF」走的是 PORTPROMPT: 端口，会自己弹保存框、不认指定的文件名，'
              + '驱动实际把它写到了\n  ' + result.relocatedFrom + '\n已帮你移到上面这个路径。'
          }
          lines.push('已用打印机「' + result.printer + '」把 ' + full + ' 打印到文件：')
          lines.push(result.outputFile + '（' + formatBytes(result.bytes) + '）' + extra)
        } else {
          lines.push('打印命令已执行，但没有收到明确结果。')
        }

        lines.push(...warningLines(result.warnings))
        return { ok, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'printer_test_page',
    description: '让指定打印机打印一张 Windows 测试页，用来确认打印链路是否正常。会报告打印后台返回的 Win32 返回码、打印机当前状态和队列里的作业状态。注意：会真的消耗纸张，只有用户明确要求时才用。',
    parameters: {
      printer: { type: 'string', required: true, description: '打印机名称，必须先在 printer_list 里确认。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const printer = String(args.printer || '').trim()
        if (!printer) throw new PrinterError('必须给出打印机名称。')
        // 名字写错时不能当成功：先确认这台机器真的存在，顺便把可用名单回给用户。
        // 列表本身读不到（WMI/注册表都不可用）时不拦，交给后面的步骤如实汇报。
        const known = await listPrinters(runOptions(exec)).catch(() => null)
        if (known && !known.printers.some((p) => p.name.toLowerCase() === printer.toLowerCase())) {
          const wanted = printer.toLowerCase()
          const near = known.printers.filter((p) => p.name.toLowerCase().includes(wanted))
          throw new PrinterError('找不到名为「' + printer + '」的打印机。'
            + (near.length ? '你是不是想用：' + near.map((p) => p.name).join('、') + '？' : '')
            + '这台机器上的打印机有：' + (known.printers.map((p) => p.name).join('、') || '（一台都没有）')
            + '。先用 printer_list 确认名称。')
        }
        const result = await printTestPage(printer, runOptions(exec))

        if (result.unverified) {
          return {
            ok: false,
            message: '已调用系统的测试页命令（退出码 ' + (result.detail || '?') + '），但当前进程读不到 WMI，'
              + '无法确认任务是否真的进了打印队列，这条结果不能当成功。\n'
              + '请用 printer_jobs 看队列，或直接看打印机面板。',
            data: result,
          }
        }

        const lines = []
        const accepted = String(result.returnValue) === '0'
        if (accepted) {
          lines.push('「' + result.printer + '」的测试页已提交给打印后台处理程序（Win32 返回码 0 = 后台接受）。')
        } else {
          lines.push('「' + result.printer + '」打印测试页失败，Win32 返回码：' + (result.returnValue || '未知') + '。')
        }
        if (result.state) lines.push('打印机当前状态：' + translateStatus(result.state) + '。')
        if (result.jobStatus) {
          lines.push('队列里的作业状态：' + result.jobStatus + '。')
        } else {
          lines.push('队列里没有查到这次作业（可能已经打完并离开队列）。')
        }
        lines.push('注意：后台接受任务不等于打印机真的出纸。如果纸没出来，请检查打印机面板和纸张，'
          + '再对照上面的状态（缺纸、卡纸、脱机、墨粉不足都会拦住它）。')
        return { ok: accepted, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'printer_set_default',
    description: '把指定打印机设为系统的默认打印机（影响之后所有不带打印机参数的打印）。',
    parameters: {
      printer: { type: 'string', required: true, description: '要设为默认的打印机名称，必须先在 printer_list 里确认。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const printer = String(args.printer || '').trim()
        if (!printer) throw new PrinterError('必须给出打印机名称。')
        const result = await setDefaultPrinter(printer, runOptions(exec))
        if (!result.verified) {
          return {
            ok: false,
            message: '没能把「' + printer + '」设成默认打印机' + (result.actual ? '，当前默认仍是「' + result.actual + '」' : '') + '。\n'
              + '常见原因：当前进程权限受限；或系统开启了「让 Windows 管理默认打印机」，默认打印机会跟着最后一次打印走。',
            data: result,
          }
        }
        return { ok: true, message: '已把「' + result.printer + '」设为系统默认打印机。', data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'printer_jobs',
    description: '查看或管理打印队列（基于 Win32_PrintJob，Windows 7 也能用）。action=list 列出排队的作业；action=cancel 取消一个作业；action=purge 清空某台打印机的全部作业（不可恢复）。空队列与读不到队列是两种不同结果，会分开说明。',
    parameters: {
      action: { type: 'string', description: 'list（默认）/ cancel / purge。' },
      printer: { type: 'string', description: '打印机名称；list 和 purge 不填表示所有打印机，cancel 建议填写以免误删同名作业。' },
      jobId: { type: 'string', description: 'action=cancel 时的作业号，从 list 的结果里取。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const action = oneOf(args.action, ['list', 'cancel', 'purge'], 'action') || 'list'
        const printer = args.printer ? String(args.printer) : ''
        if (action === 'list') {
          const result = await listJobs(printer, runOptions(exec))
          if (result.jobs.length === 0) {
            const lines = [printer ? '打印机「' + printer + '」的队列是空的。' : '所有打印机的队列都是空的。']
            lines.push(...warningLines(result.warnings))
            return { ok: true, message: lines.join('\n'), data: result }
          }
          const lines = ['队列里有 ' + result.jobs.length + ' 个作业：']
          for (const job of result.jobs) {
            lines.push('')
            lines.push('- 作业号 ' + job.id + '（' + job.printer + '）：' + (job.document || '(无文档名)'))
            const bits = []
            if (job.status) bits.push('状态：' + job.status)
            if (job.pages !== null) bits.push('页数：' + job.pages)
            if (job.owner) bits.push('所有者：' + job.owner)
            if (bits.length) lines.push('  ' + bits.join('，'))
          }
          lines.push(...warningLines(result.warnings))
          return { ok: true, message: lines.join('\n'), data: result }
        }
        if (action === 'cancel') {
          const jobId = String(args.jobId || '').trim()
          if (!jobId) throw new PrinterError('取消作业必须给出 jobId（先用 action=list 查看）。', 'noJobId')
          const result = await cancelJob(printer, jobId, runOptions(exec))
          const tail = result.readback === 'gone'
            ? '，已回读确认它从队列里消失'
            : result.status
              ? '，但回读时作业还在，状态：' + result.status
              : ''
          return { ok: true, message: '已请求取消作业 ' + result.jobId + (result.printer ? '（' + result.printer + '）' : '') + tail + '。', data: result }
        }
        const result = await purgeJobs(printer, runOptions(exec))
        const scope = result.printer ? '打印机「' + result.printer + '」' : '所有打印机'
        const lines = []
        if (!result.verified) {
          lines.push('已向' + scope + '的队列发出清空请求，但脚本没有给出回读结果，没法确认队列真的空了。请用 action=list 复查。')
        } else {
          lines.push('已清空' + scope + '的队列：请求取消 ' + result.requested + ' 个作业，回读确认删掉 ' + result.deleted + ' 个。')
          if (result.remaining > 0) {
            lines.push('注意：回读时队列里还有 ' + result.remaining + ' 个作业没删掉，可能是作业正在出纸、被驱动锁住或权限不足。可以用 action=list 看具体情况。')
          }
        }
        return { ok: result.verified && result.remaining === 0, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'scanner_list',
    description: '列出这台电脑上的扫描设备（WIA 设备：平板扫描仪、带扫描功能的喷墨/激光一体机、相机等），并逐个试开一次、标出此刻能不能用。同一个型号可能同时有 WIA 与 eSCL 两个条目，其中常有一个长期不在线。只读操作。',
    parameters: {},
    output: TEXT_OUTPUT,
    async execute(_args, exec) {
      try {
        const result = await listScanners(runOptions(exec))
        if (result.scanners.length === 0) {
          const lines = ['没有找到扫描设备。请确认扫描仪已开机、驱动已装，并且 Windows Image Acquisition (WIA) 服务在运行。']
          lines.push(...warningLines(result.warnings))
          return { ok: true, message: lines.join('\n'), data: result }
        }
        const lines = ['共 ' + result.scanners.length + ' 台扫描设备：']
        const usable = result.scanners.filter((s) => s.available !== false)
        result.scanners.forEach((s, index) => {
          lines.push('')
          lines.push((index + 1) + '. ' + s.name)
          if (s.typeName) lines.push('   类型：' + s.typeName + (s.type && s.typeName.indexOf(s.type) === -1 ? '（WIA 类型 ' + s.type + '）' : ''))
          if (s.id) lines.push('   DeviceID：' + s.id)
          if (s.available === true) lines.push('   当前状态：可用')
          else if (s.available === false) lines.push('   当前状态：现在打不开' + (s.error ? '（' + s.error + '）' : '') + '，扫描时跳过它')
        })
        lines.push('')
        if (usable.length < result.scanners.length) {
          lines.push('说明：同一个型号可能同时登记了 WIA 和 eSCL 两个条目，其中常有一个长期不在线；不指定 device 时插件会自动挑能打开的那台。')
          lines.push('')
        }
        lines.push('提示：扫描时 device 参数可以填名称、DeviceID 或上面的序号；想知道某台支不支持进纸器、有哪些分辨率，用 scanner_info。')
        lines.push(...warningLines(result.warnings))
        return { ok: true, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'scanner_info',
    description: '查询扫描设备的能力，不真的扫描：有没有进纸器/平板玻璃、支不支持双面、驱动报告的可用格式与分辨率、当前的进纸设置。不填 device 时列出所有设备。只读操作。',
    parameters: {
      device: { type: 'string', description: '扫描设备：名称、DeviceID 或 scanner_list 里的序号；不填看全部设备。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const result = await scannerInfo({ device: args.device ? String(args.device) : '' }, runOptions(exec))
        if (result.devices.length === 0) {
          const lines = ['没有找到扫描设备。请确认扫描仪已开机、驱动已装，并且 WIA 服务在运行。']
          lines.push(...warningLines(result.warnings))
          return { ok: true, message: lines.join('\n'), data: result }
        }
        const lines = ['共 ' + result.devices.length + ' 台扫描设备：']
        for (const d of result.devices) {
          lines.push('')
          lines.push('- ' + d.name)
          const src = []
          if (d.hasFeeder === true) src.push('进纸器：支持')
          else if (d.hasFeeder === false) src.push('进纸器：不支持')
          if (d.hasFlatbed === true) src.push('平板玻璃：支持')
          else if (d.hasFlatbed === false) src.push('平板玻璃：不支持')
          if (d.hasDuplex === true) src.push('双面：支持')
          if (src.length) lines.push('  ' + src.join('，'))
          if (d.supportedFormats.length) lines.push('  可用格式：' + d.supportedFormats.join('、'))
          else lines.push('  可用格式：驱动没有公开（扫描时插件会自动回退并转码）')
          if (d.supportedDpi.length) lines.push('  可用分辨率：' + d.supportedDpi.join('、') + ' dpi')
          if (d.currentSourceText) lines.push('  当前进纸设置：' + d.currentSourceText)
          if (d.note) lines.push('  备注：' + d.note)
        }
        lines.push('')
        lines.push('提示：想扫多页进纸器用 source=feed，要合成一个 PDF 用 format=pdf。')
        lines.push(...warningLines(result.warnings))
        return { ok: true, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'scanner_scan',
    description: '用扫描仪扫一页或多页：可以存成图片（jpg/png/tif/bmp），也可以直接合并成一个多页 PDF（format=pdf，页面以 JPEG 原样嵌入，不额外损失画质）。扫描会占用设备并可能耗时几十秒；单次最多 50 页。多页只有进纸器能做，平板玻璃一次一页。format=tif 时每页是单独一个文件（不合成多页 TIFF）。进纸器扫空是正常结束，已经扫到的页面都会保留；中途出错也会把已完成的那部分页面保留下来。',
    parameters: {
      output: { type: 'string', required: true, description: '输出路径：绝对路径，或相对当前会话工作目录。存图片时多页会从第 2 页起自动插入 -2、-3…；format=pdf 时这是 PDF 的路径（扩展名会自动纠正成 .pdf）。' },
      device: { type: 'string', description: '扫描设备：名称（可以只写前面的部分）、DeviceID 或 scanner_list 里的序号；不填时自动挑一台现在能打开的设备（跳过离线条目）。' },
      dpi: { type: 'integer', description: '分辨率，默认 300；可填 75~1200，但实际能到多少由设备决定，驱动不接受时会取最接近的可用值并如实汇报（想先知道就查 scanner_info）。低分辨率更快、文件更小；文档 200~300 够用。' },
      mode: { type: 'string', description: 'color（默认）/ gray（灰度）/ text（黑白文档，体积最小）。' },
      format: { type: 'string', description: 'jpg（默认）/ png / tif / bmp / pdf。pdf 会把所有页面合并成一个多页文件（页面以 JPEG 原样嵌入）；tif 每页各存一个文件。' },
      source: { type: 'string', description: 'auto（默认：有进纸器就先走进纸器，进纸器空就改用平板玻璃并说明）/ flat（平板玻璃，只扫一页）/ feed（进纸器，可连续多页）。' },
      crop: { type: 'string', description: 'off（默认）/ a4。进纸器按固定行程出图（常见 14 英寸），A4 原稿右侧/底部会多出一条空白；a4 会直接裁到 A4 尺寸。注意这是破坏性裁剪：裁掉的部分不会留在输出文件里，只在确实多出空白时才用。' },
      keepPages: { type: 'boolean', description: 'format=pdf 时是否保留每页的 jpg 分页文件，默认 false（只留 PDF）。' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      try {
        const output = String(args.output || '').trim()
        if (!output) throw new PrinterError('必须给出输出文件路径。')
        const mode = oneOf(args.mode, ['color', 'gray', 'text'], 'mode') || 'color'
        const format = oneOf(args.format, ['jpg', 'png', 'tif', 'bmp', 'pdf'], 'format') || 'jpg'
        const source = oneOf(args.source, ['auto', 'flat', 'feed'], 'source') || 'auto'
        const crop = oneOf(args.crop, ['off', 'a4'], 'crop') || 'off'
        let dpi = args.dpi === undefined || args.dpi === null ? 300 : Number(args.dpi)
        if (!Number.isFinite(dpi) || dpi <= 0) dpi = 300
        if (dpi > 1200) throw new PrinterError('分辨率上限是 1200 dpi；实际可用的分辨率以设备为准，可以用 scanner_info 查（这台机器常见的是 75/150/300）。')

        const full = resolvePath(exec, output)
        const result = await scan({
          device: args.device ? String(args.device) : '',
          output: full,
          dpi,
          mode,
          format,
          source,
          crop,
          keepPages: args.keepPages === true,
        }, runOptions(exec))

        const sourceText = result.source === 'feeder' ? '进纸器' : (result.source === 'flatbed' ? '平板' : '')
        const lines = []
        // 中途失败时不能写「扫描完成」：这一趟只完成了一部分，页数也不是全部的页数。
        const heading = result.failure ? '扫描未完成' : '扫描完成'
        const who = result.device ? result.device : '扫描仪'
        if (result.pdfPath) {
          lines.push(heading + '：' + who + (sourceText ? '（' + sourceText + '）' : '') + '，共 ' + result.pages.length + ' 页，已合并成一个 PDF：')
          lines.push('- ' + result.pdfPath + '（' + formatBytes(result.pdfBytes) + '）')
          const kept = result.pages.filter((page) => page.kept && page.path)
          if (kept.length) {
            lines.push('分页图片也保留了：')
            for (const page of kept) lines.push('- 第 ' + page.index + ' 页：' + page.path + '（' + formatBytes(page.bytes) + '）')
          } else {
            lines.push('（分页图片没有保留，需要的话下次传 keepPages=true。）')
          }
        } else {
          lines.push(heading + '：' + who + (sourceText ? '（' + sourceText + '）' : '') + '，共 ' + result.pages.length + ' 页。')
          for (const page of result.pages) {
            const bits = [formatBytes(page.bytes)]
            if (page.format) bits.push(page.format.toUpperCase())
            if (page.width && page.height) bits.push(page.width + '×' + page.height + ' 像素')
            if (page.dpi) bits.push(page.dpi + ' dpi')
            lines.push('- 第 ' + page.index + ' 页：' + page.path + '（' + bits.join('，') + '）')
          }
        }

        if (result.crop === 'a4') {
          lines.push('')
          lines.push('已按 A4 裁掉右侧/底部多出来的空白（crop=a4）。')
        }
        if (result.stoppedBy === 'paperEmpty' && result.source === 'feeder') {
          lines.push('')
          lines.push('进纸器没有更多纸了，扫描到此结束。这是正常结束，不是故障。')
        }
        if (result.stoppedBy === 'maxPages') {
          lines.push('')
          lines.push('已经扫到 50 页上限，剩下的纸没有继续扫。')
        }
        if (result.fallback === 'feederEmpty') {
          lines.push('')
          lines.push('注意：进纸器里没有纸，这一页是从平板玻璃上取的。')
        }
        if (result.requestedDpi && result.effectiveDpi && result.requestedDpi !== result.effectiveDpi) {
          lines.push('')
          let note = '注意：这台扫描仪不接受 ' + result.requestedDpi + ' dpi，实际按 ' + result.effectiveDpi + ' dpi 扫描。'
          if (result.supportedDpi.length) note += '它支持 ' + result.supportedDpi.join('、') + ' dpi。'
          lines.push(note)
        }
        if (result.pages.some((page) => page.converted)) {
          lines.push('')
          const native = result.nativeFormat ? '驱动实际给出的是 ' + result.nativeFormat.toUpperCase() + '，' : '驱动给出的原始格式与请求不同，'
          lines.push('注意：' + native + '已自动转成 ' + format + '。')
        }
        if (result.supportedFormats.length && format !== 'pdf' && !result.supportedFormats.includes(format)) {
          lines.push('')
          lines.push('提示：该设备报告的可用格式是 ' + result.supportedFormats.join('、') + '，请求的 ' + format + ' 可能不被驱动支持。')
        }
        if (result.failure) {
          lines.push('')
          const done = result.failure.completedPages
            ? '，出错之前扫到的 ' + result.failure.completedPages + ' 页都已经保留'
            : ''
          const stopText = {
            paperEmpty: '进纸器没有纸了',
            maxPages: '到了页数上限',
            error: '传输出错',
            noItem: '设备没有提供可扫描的条目',
          }
          const stop = result.failure.stoppedBy ? '，中止原因：' + (stopText[result.failure.stoppedBy] || result.failure.stoppedBy) : ''
          lines.push('注意：这一趟中途出错了（' + result.failure.message + '）' + stop + done + '。修好之后可以只扫剩下的页再合并。')
        }
        const notes = warningLines(result.warnings)
        if (notes.length) {
          lines.push('')
          lines.push(...notes)
        }

        return { ok: true, message: lines.join('\n'), data: result }
      } catch (error) {
        return fail(error)
      }
    },
  }))
}
