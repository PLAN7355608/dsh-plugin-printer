/**
 * 错误文案表：脚本只输出 ASCII，并把稳定的错误码放在 CODE 记录里，中文提示统一在这里生成。
 * 放在单独的模块里，是因为打印、队列、扫描三组封装的错误码共用一张表、一套翻译规则。
 */
import { PrinterError } from './ps.js'
/**
 * ps/*.ps1 只输出 ASCII，脚本通过 CODE 行给出稳定的错误码，这里按码查表翻成人话。
 * 三条贯穿全表的判断：
 *   1. 「读不到」和「没有」是两件事（队列、打印机列表、设备能力）；
 *   2. 「请求发出去了」和「真的生效了」是两件事（取消/清空作业会回读复核）；
 *   3. 拒绝的时候要告诉调用方下一步该怎么办，而不是只说不行。
 * 表里没有的码原样抛回，避免把未知错误吞掉。
 */
const ERROR_TEXT = {
  // 打印文件（print-file.ps1 / print-gdi.ps1）
  fileNotFound: (d) => '找不到要打印的文件：' + d,
  notAFile: (d) => '这个路径不是文件（可能是个文件夹）：' + d,
  unknownMode: () => '不认识的打印模式。可用 auto / gdi / queue / verb / file。',
  unknownDuplex: () => '不认识的双面方向。可用 off（单面）、long（长边翻页）、short（短边翻页）。',
  outputRequired: () => 'mode=file 需要同时给出输出文件路径（outputFile）。',
  noPrinter: () => '没有指定打印机，而且系统没有默认打印机。请显式传打印机名。',
  unsupportedSelection: () => 'queue / verb / file 这三条路径把文件整个交给系统，插件插不上手，所以设不了页码与逐份/逐页。请改用 mode=gdi（文本与图片文件）。',
  planOnlyUnsupported: () => 'planOnly 只能用 mode=gdi：只有 GDI 渲染后端能给出页面计划，其它路径无法规划。这次没有打印任何东西。',
  notText: () => '这个文件不是文本：开头是二进制格式的标识（PDF、Office 文档、压缩包或可执行文件），GDI 路径只渲染文本与图片。',
  fileUnreadable: (d) => '文件读不出来（通常是权限不够或文件被占用）：' + d,
  tooManyPages: () => '算出来的页数远超预期，熔断拒绝了这次打印；确认文件无误后可传 allowHugeJob 放行。',
  noOutput: () => '打印驱动没有写出目标文件：关联程序忽略了“打印到文件”的要求。',
  noPrintToVerb: () => '这个文件类型的关联程序没有注册 printto 打印动词（WPS Office 全系如此），没法把目标打印机交给它。显式传 allowDefaultSwitch 可临时切换系统默认打印机，或者先把文件转成文本/PDF 再打。',
  switchFailed: () => '临时切换系统默认打印机失败。',
  printFailed: () => '打印失败。',
  gdiUnavailable: () => '这台电脑上取不到 GDI+ 打印后端（System.Drawing 加载失败），没法用 gdi 模式。',
  badPageRange: () => '页码范围读不出来或超出范围。',
  noPrintablePages: () => '这个文件没有生成任何可打印的页面。',
  unknownPrinter: (d) => '目标打印机不在已安装的打印机里。' + d,
  capsUnreadable: (d) => '读不到这台打印机的能力与偏好：查询被系统拒绝，这不代表打印机没有能力。' + d,
  listUnreadable: (d) => '打印机列表的两个来源都读不到（注册表与 WMI 都失败），所以无法判断这台电脑有没有打印机。' + d,
  // 打印队列（printer-jobs.ps1）
  queueUnreadable: () => '读不到打印队列。这跟「队列是空的」不是一回事：空队列会正常返回 0 个作业；读不到说明打印后台拒绝了这次查询，通常是权限不够，或者 Print Spooler 服务没在运行。',
  jobNotFound: (d) => '队列里找不到这个作业：' + d,
  noJobId: () => '没有指定作业号。',
  deleteFailed: (d) => '打印后台拒绝删除作业：' + d,
  verifyFailed: (d) => '删除请求已经发出，但回读队列失败，无法确认作业是否真的消失：' + d,
  stillQueued: (d) => '删除请求已经发出，但作业仍留在队列里（打印后台可以让它挂在“正在删除”状态，直到打印机确认）。' + d,
  unknownAction: (d) => '不认识的操作：' + d,
  // 扫描（scan.ps1 / scanner-info.ps1 / list-scanners.ps1）—— 由 lib/scanners.js 使用
  scannerUnreadable: (d) => '读不到扫描设备：WIA 服务或设备管理器打不开，这不代表这台电脑没有扫描仪。' + d,
  noScanner: () => '这台电脑上没有可用的 WIA 扫描设备。',
  scannerNotFound: (d) => '找不到指定的扫描设备：' + d,
  scannerAmbiguous: (d) => '按这个前缀能匹配到多台扫描设备，请写出完整名称。' + d,
  noFeeder: () => '这台扫描仪没有进纸器。',
  noFlatbed: () => '这台扫描仪没有平板玻璃。',
  feederEmpty: () => '进纸器是空的：装纸后重试，或者改用平板玻璃扫描。',
  noPage: () => '扫描仪没有返回任何页面。',
  transferFailed: (d) => '扫描过程中设备传输失败，已经扫到的页面已按图片保存。' + d,
  saveFailed: (d) => '页面扫出来了，但写文件失败（磁盘空间、权限或路径不可写）。' + d,
  pdfFailed: (d) => '生成多页 PDF 失败，已经扫到的页面已按图片保存。' + d,
  pdfTooLarge: (d) => '扫描页累计超过单个 PDF 的体积上限，没有合并：' + d + ' 可以改用 format=jpg 逐页保存，或降低分辨率、分几次扫描。',
  deviceError: (d) => '扫描设备出错。' + d,
  unknownFormat: () => '不认识的输出格式。可用 jpg / png / tif / bmp / pdf。',
  unknownSource: () => '不认识的进纸来源。可用 auto / flat / feed。',
  unknownCrop: () => '不认识的裁切方式。可用 off / a4。',
}

/** 把脚本错误码翻成中文；没有登记过的码原样返回，避免吞掉新错误。 */
export function translateError(error) {
  if (!(error instanceof PrinterError)) return error
  const code = error && error.code ? String(error.code) : ''
  const entry = ERROR_TEXT[code]
  if (!entry) return error
  const detail = error && error.message ? String(error.message) : ''
  return new PrinterError(entry(detail), error.code)
}

export function translateQueueError(error) { return translateError(error) }
export function translatePrintError(error) { return translateError(error) }
