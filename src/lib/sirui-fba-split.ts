/**
 * 思锐FBA下单发票拆分 (Sirui FBA order invoice splitting)
 *
 * 业务背景：
 * 客户提供一张「下单发票」（一个 FBA 仓库集合的多行明细），内部报价表里记录了每个
 * FBA 码对应的「仓点」(如 FAT2) 与「渠道」(如 OA-拆送（OA联盟）)。本模块根据报价表，
 * 把一张下单发票按 FBA 码拆分为多张发票，每张文件名 = {FBA}-{仓点}.xlsx。
 *
 * 拆分后的每张发票需要：
 *  - 只保留该 FBA 码对应的明细行，并把行号从 R25 重新连续编号；
 *  - F9 收件人国家：USA → 美国（国别码转中文）；
 *  - F10 亚马逊仓库代码：填仓点（如 FAT2）；
 *  - B21 备注：填渠道（如 OA-拆送（OA联盟））；
 *  - K 列申报币种：美元 → USD；
 *  - 每行 W 列填 {FBA}-{仓点}；
 *  - 删除「服务渠道」「FBA地址」两个 sheet，输出只保留 1 个 sheet；
 *  - 保留 WPS DISPIMG 产品图片（cellimages.xml + media）。
 *
 * 技术要点（与代码库其它功能一致的 zip 手术方案）：
 *  - 用 ExcelJS 读取报价表与发票数据；
 *  - 用 JSZip 对源 xlsx 做字符串级手术，避免 ExcelJS 写回时丢失 DISPIMG 图片；
 *  - 新增单元格一律写入共享字符串 (t="s") 而非内联字符串 (inlineStr) —— 源发票所有中文
 *    都在 sharedStrings，易通系统已验证能读；inlineStr 是易通解析器可能不支持的类型。
 */

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readFileSync } from "fs";

// ============================================================
// Types
// ============================================================

interface QuoteEntry {
  warehouse: string; // 仓点，如 FAT2
  channel: string; // 渠道，如 OA-拆送（OA联盟）
}

interface InvoiceRow {
  sourceRow: number; // 源行号（>=25）
  boxId: string; // A 列：货箱编号
  currency: string; // K 列：申报币种（如 美元）
}

interface Group {
  fba: string;
  warehouse: string;
  channel: string;
  rows: InvoiceRow[];
}

export interface SiruiFileEntry {
  fba: string;
  warehouse: string;
  fileName: string;
  rowCount: number;
}

export interface SiruiSplitResult {
  buffer: Buffer; // 所有拆分文件的 ZIP
  zipFileName: string;
  files: SiruiFileEntry[];
  warnings: string[];
  total: number;
}

// ============================================================
// Constants
// ============================================================

const COUNTRY_CN: Record<string, string> = {
  US: "美国", USA: "美国",
  DE: "德国", DEU: "德国",
  GB: "英国", GBR: "英国",
  FR: "法国", FRA: "法国",
  IT: "意大利", ITA: "意大利",
  ES: "西班牙", ESP: "西班牙",
  JP: "日本", JPN: "日本",
  CA: "加拿大", CAN: "加拿大",
  AU: "澳大利亚", AUS: "澳大利亚",
  NL: "荷兰", NLD: "荷兰",
  MX: "墨西哥", MEX: "墨西哥",
  BE: "比利时", BEL: "比利时",
};

const CURRENCY_TO_ISO: Record<string, string> = {
  美元: "USD", 欧元: "EUR", 英镑: "GBP", 加币: "CAD", 加元: "CAD",
  人民币: "CNY", 日元: "JPY", 澳元: "AUD", 港币: "HKD", 墨西哥比索: "MXN",
};

const DATA_START_ROW = 25; // 明细区起始行号
const HEADER_LAST_ROW = 24; // 表头区结束行号
const ROW_EMU = 44 * 12700; // 明细行高 44pt = 558800 EMU（12700 EMU/pt）

// ============================================================
// Small helpers
// ============================================================

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function cellText(cell: ExcelJS.Cell): string {
  const v: unknown = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    const o = v as Record<string, unknown>;
    if ("result" in o) return o.result == null ? "" : String(o.result);
    if ("richText" in o) return (o.richText as { text?: string }[]).map((t) => t.text ?? "").join("");
    if (typeof o.text === "string") return o.text;
  }
  return String(v);
}

function toCnCountry(c: string): string {
  return COUNTRY_CN[c] ?? c;
}

function toIsoCurrency(c: string): string {
  const iso = CURRENCY_TO_ISO[c];
  if (iso) return iso;
  if (/^[A-Z]{3}$/.test(c)) return c; // 已是 ISO 代码
  return c;
}

// ============================================================
// 1. Parse 内部报价 → Map<FBA, {仓点, 渠道}>
// ============================================================

async function parseQuote(filePath: string): Promise<Map<string, QuoteEntry>> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.getWorksheet("内部");
  if (!ws) throw new Error("内部报价表里未找到「内部」工作表");

  const map = new Map<string, QuoteEntry>();
  ws.eachRow({ includeEmpty: false }, (row) => {
    const fba = cellText(row.getCell(4)); // D 列 = FBA 码
    if (!/^FBA[0-9A-Z]{6,}$/i.test(fba)) return; // 跳过表头与合计行
    const warehouse = cellText(row.getCell(3)); // C 列 = 仓点
    const channel = cellText(row.getCell(8)); // H 列 = 渠道
    if (!warehouse) return;
    map.set(fba, { warehouse, channel });
  });

  if (map.size === 0) throw new Error("「内部」工作表里未解析到 FBA 报价数据");
  return map;
}

// ============================================================
// 2. Parse 下单发票数据（ExcelJS，仅读业务字段）
// ============================================================

function matchFba(boxId: string, quoteMap: Map<string, QuoteEntry>): string | null {
  let best: string | null = null;
  for (const fba of quoteMap.keys()) {
    if (boxId.startsWith(fba)) {
      if (best === null || fba.length > best.length) best = fba;
    }
  }
  return best;
}

async function parseInvoice(
  filePath: string,
  quoteMap: Map<string, QuoteEntry>
): Promise<{ country: string; groups: Group[]; unmatched: InvoiceRow[] }> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];

  const country = cellText(ws.getCell("F9")); // 源发票收件人国家，如 USA

  const byFba = new Map<string, InvoiceRow[]>();
  const unmatched: InvoiceRow[] = [];

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < DATA_START_ROW) return;
    const boxId = cellText(row.getCell(1)); // A
    const currency = cellText(row.getCell(11)); // K
    if (!boxId) return; // 空行（如合并续行无箱号）
    const fba = matchFba(boxId, quoteMap);
    const rec: InvoiceRow = { sourceRow: rowNumber, boxId, currency };
    if (!fba) {
      unmatched.push(rec);
      return;
    }
    if (!byFba.has(fba)) byFba.set(fba, []);
    byFba.get(fba)!.push(rec);
  });

  const groups: Group[] = [];
  for (const [fba, rows] of byFba) {
    const q = quoteMap.get(fba)!;
    rows.sort((a, b) => a.sourceRow - b.sourceRow);
    groups.push({ fba, warehouse: q.warehouse, channel: q.channel, rows });
  }
  groups.sort((a, b) => a.fba.localeCompare(b.fba));

  return { country, groups, unmatched };
}

// ============================================================
// Shared strings editor（追加共享字符串，保持 t="s"）
// ============================================================

class SharedStringsEditor {
  private xml: string;
  private nextIndex: number;

  constructor(xml: string) {
    this.xml = xml;
    this.nextIndex = (xml.match(/<si>/g) || []).length;
  }

  append(text: string): number {
    const idx = this.nextIndex;
    const esc = escapeXml(text);
    this.xml = this.xml
      .replace(/ uniqueCount="(\d+)"/, (_m, n) => ` uniqueCount="${Number(n) + 1}"`)
      .replace(/ count="(\d+)"/, (_m, n) => ` count="${Number(n) + 1}"`)
      .replace(/<\/sst>/, `<si><t>${esc}</t></si></sst>`);
    this.nextIndex += 1;
    return idx;
  }

  getXml(): string {
    return this.xml;
  }
}

// ============================================================
// Cell manipulation（字符串级，作用于 XML 片段）
// ============================================================

/** 把指定单元格设置为共享字符串值（兼容空单元格与已有值单元格）。 */
/** 统计某列数据行（>=DATA_START_ROW）出现最多的样式索引，用于对齐参考输出的统一样式。 */
function dominantStyle(sheet1Xml: string, col: string): string | null {
  const counts = new Map<string, number>();
  const re = new RegExp(`<c r="${col}(\\d+)" s="(\\d+)"`, "g");
  let m: RegExpExecArray | null;
  while ((m = re.exec(sheet1Xml))) {
    const r = Number(m[1]);
    if (r < DATA_START_ROW) continue;
    counts.set(m[2], (counts.get(m[2]) ?? 0) + 1);
  }
  let best: string | null = null;
  let bestN = -1;
  for (const [s, n] of counts) {
    if (n > bestN) {
      bestN = n;
      best = s;
    }
  }
  return best;
}

function setCellSharedStr(xml: string, ref: string, idx: number, forceStyle?: string): string {
  // 分别匹配自闭合 <c r="REF" .../> 与成对 <c r="REF" ...>...</c>，避免非贪婪正则吞掉后续单元格 / 丢样式
  const build = (attrs: string) => {
    let cleaned = attrs.replace(/\s+t="[^"]*"/, "");
    if (forceStyle) {
      cleaned = cleaned.replace(/\s+s="[^"]*"/g, "");
      cleaned += ` s="${forceStyle}"`;
    }
    cleaned = cleaned.trim();
    return `<c r="${ref}"${cleaned ? " " + cleaned : ""} t="s"><v>${idx}</v></c>`;
  };

  const selfClose = new RegExp(`<c r="${ref}"([^>]*?)/>`);
  const m1 = selfClose.exec(xml);
  if (m1) {
    return xml.slice(0, m1.index) + build(" " + m1[1]) + xml.slice(m1.index + m1[0].length);
  }

  const paired = new RegExp(`<c r="${ref}"([^>]*?)>[\\s\\S]*?<\\/c>`);
  const m2 = paired.exec(xml);
  if (m2) {
    return xml.slice(0, m2.index) + build(" " + m2[1]) + xml.slice(m2.index + m2[0].length);
  }
  return xml;
}

/** 展开 shared formula（<f t="shared" ...>）为独立公式：跨组拆分后 master 可能落在别的组被删，成员会读不到公式。 */
function expandSharedFormulas(rowXml: string): string {
  // 1) 成员 <f t="shared" si="N"/><v>=DISPIMG(...)</v> → 用 <v> 缓存值生成独立公式
  rowXml = rowXml.replace(
    /<f t="shared"[^>]*\/><v>([^<]*)<\/v>/g,
    (_m, vtext: string) => `<f>_xlfn.${vtext.replace(/^=/, "")}</f><v>${vtext}</v>`
  );
  // 2) master <f t="shared" ref="..." si="N">TEXT</f> → <f>TEXT</f>
  rowXml = rowXml.replace(
    /<f t="shared"[^>]*>\s*([\s\S]*?)\s*<\/f>/g,
    (_m, text: string) => `<f>${text}</f>`
  );
  return rowXml;
}

/** 在行尾追加一个共享字符串单元格（用于 W 列）。 */
function appendSharedStrCell(
  rowXml: string,
  col: string,
  rowNum: number,
  style: string | null,
  idx: number
): string {
  const styleAttr = style ? ` s="${style}"` : "";
  const cell = `<c r="${col}${rowNum}"${styleAttr} t="s"><v>${idx}</v></c>`;
  return rowXml.replace(/<\/row>$/, cell + "</row>");
}

/** 行重编号：更新 <row r="N">、所有单元格 r="COLN"、spans 含 W 列(23)。 */
function renumberRow(rowXml: string, newRow: number): string {
  let out = rowXml.replace(/<row r="\d+"/, `<row r="${newRow}"`);
  out = out.replace(/r="([A-Z]+)\d+"/g, (_m, col: string) => `r="${col}${newRow}"`);
  out = out.replace(/spans="1:\d+"/, `spans="1:23"`);
  return out;
}

// ============================================================
// 3. Build single output file（zip 手术）
// ============================================================

interface RawParts {
  sharedStrings: string;
  sheet1: string;
  cellimages: string;
  cellimagesRels: string;
  workbook: string;
  workbookRels: string;
  contentTypes: string;
  media: Map<string, Buffer>; // path -> bytes
  staticFiles: Map<string, Buffer>; // 原样保留的二进制/文本文件
}

async function readRawParts(zip: JSZip): Promise<RawParts> {
  const text = (p: string) => zip.file(p)!.async("string");
  const buf = (p: string) => zip.file(p)!.async("nodebuffer");

  const sharedStrings = await text("xl/sharedStrings.xml");
  const sheet1 = await text("xl/worksheets/sheet1.xml");
  const cellimages = await text("xl/cellimages.xml");
  const cellimagesRels = await text("xl/_rels/cellimages.xml.rels");
  const workbook = await text("xl/workbook.xml");
  const workbookRels = await text("xl/_rels/workbook.xml.rels");
  const contentTypes = await text("[Content_Types].xml");

  const media = new Map<string, Buffer>();
  const mediaPaths = Object.keys(zip.files).filter(
    (p) => p.startsWith("xl/media/") && !p.endsWith("/")
  );
  for (const p of mediaPaths) media.set(p, await buf(p));

  const staticPaths = [
    "_rels/.rels",
    "docProps/app.xml",
    "docProps/core.xml",
    "docProps/custom.xml",
    "xl/styles.xml",
    "xl/theme/theme1.xml",
    "xl/comments1.xml",
    "xl/customProperty1.bin",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/drawings/vmlDrawing1.vml",
    "xl/worksheets/_rels/sheet1.xml.rels",
  ];
  const staticFiles = new Map<string, Buffer>();
  for (const p of staticPaths) staticFiles.set(p, await buf(p));

  return {
    sharedStrings,
    sheet1,
    cellimages,
    cellimagesRels,
    workbook,
    workbookRels,
    contentTypes,
    media,
    staticFiles,
  };
}

interface CellImageInfo {
  name: string;
  rEmbed: string;
  offY: number;
  rawXml: string;
}

function parseCellImages(xml: string): CellImageInfo[] {
  const items = xml.match(/<etc:cellImage>[\s\S]*?<\/etc:cellImage>/g) || [];
  return items.map((raw) => {
    const name = /<xdr:cNvPr[^>]*name="([^"]+)"/.exec(raw)?.[1] ?? "";
    const rEmbed = /r:embed="([^"]+)"/.exec(raw)?.[1] ?? "";
    const offM = /<a:off[^>]*y="(\d+)"/.exec(raw);
    return { name, rEmbed, offY: offM ? parseInt(offM[1], 10) : 0, rawXml: raw };
  });
}

function remapCellImageY(rawXml: string, newY: number): string {
  return rawXml.replace(/(<a:off[^>]*y=")(\d+)(")/, (_m, p1: string, _p2: string, p3: string) => {
    return `${p1}${newY}${p3}`;
  });
}

async function buildOutput(
  group: Group,
  countryCn: string,
  currencyIso: string,
  raw: RawParts
): Promise<Buffer> {
  const out = new JSZip();

  // ---- sharedStrings：追加新值 ----
  const sse = new SharedStringsEditor(raw.sharedStrings);
  const idxCountry = sse.append(countryCn);
  const idxCurrency = sse.append(currencyIso);
  const idxWarehouse = sse.append(group.warehouse);
  const idxChannel = sse.append(group.channel);
  const idxKey = sse.append(`${group.fba}-${group.warehouse}`);

  // ---- sheet1：过滤 / 重编号 / 改单元格 ----
  const sdMatch = /<sheetData>([\s\S]*?)<\/sheetData>/.exec(raw.sheet1);
  if (!sdMatch) throw new Error("源发票缺少 sheetData");
  const allRows = sdMatch[1].match(/<row\b[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\/>/g) || [];

  const sourceRowSet = new Set(group.rows.map((r) => r.sourceRow));
  const rowMap = new Map<number, number>();
  group.rows.forEach((r, i) => rowMap.set(r.sourceRow, DATA_START_ROW + i));

  // K 列币种统一用「无填充」主样式（源发票部分行的币种带黄色高亮，参考输出统一无填充）
  const kStyle = dominantStyle(raw.sheet1, "K");

  const headerRows: string[] = [];
  const dataRows: string[] = [];
  for (const rowXml of allRows) {
    const rM = /<row r="(\d+)"/.exec(rowXml);
    if (!rM) continue;
    const r = Number(rM[1]);
    if (r < DATA_START_ROW) {
      headerRows.push(rowXml);
    } else if (sourceRowSet.has(r)) {
      const newRow = rowMap.get(r)!;
      const rowStyle = /<row\b[^>]*\bs="(\d+)"/.exec(rowXml)?.[1] ?? null;
      let newRowXml = renumberRow(rowXml, newRow);
      // V 列 DISPIMG 公式：展开 shared formula（master 可能已被过滤到别的组）
      newRowXml = expandSharedFormulas(newRowXml);
      // K 列：币种 -> ISO（统一无填充样式）
      newRowXml = setCellSharedStr(newRowXml, `K${newRow}`, idxCurrency, kStyle ?? undefined);
      // W 列：{FBA}-{仓点}
      newRowXml = appendSharedStrCell(newRowXml, "W", newRow, rowStyle, idxKey);
      dataRows.push(newRowXml);
    }
  }

  const newSheetData =
    headerRows.join("") + dataRows.join("");

  const lastRow = HEADER_LAST_ROW + group.rows.length;

  // dimension
  let pre = raw.sheet1.slice(0, sdMatch.index) + "<sheetData>" + newSheetData + "</sheetData>";
  let post = raw.sheet1.slice(sdMatch.index + sdMatch[0].length);
  pre = pre.replace(/<dimension ref="A1:W\d+"\/>/, `<dimension ref="A1:W${lastRow}"/>`);
  // W 列宽对齐参考输出（25.21pt），并标记 customWidth
  pre = pre.replace(/<col min="23" max="23"[^>]*?\/>/, (m) =>
    m
      .replace(/\s+width="[^"]*"/, "")
      .replace(/\s+customWidth="\d*"/, "")
      .replace(/\/>$/, ' width="25.2090909090909" customWidth="1"/>')
  );

  // mergeCells 过滤
  post = filterMergeCells(post, rowMap);
  // dataValidations 过滤
  post = filterDataValidations(post);

  let sheet1 = pre + post;

  // 表头单元格：F9 国家 / F10 仓库代码 / B21 备注
  sheet1 = setCellSharedStr(sheet1, "F9", idxCountry);
  sheet1 = setCellSharedStr(sheet1, "F10", idxWarehouse);
  sheet1 = setCellSharedStr(sheet1, "B21", idxChannel);

  // ---- cellimages：过滤 + 去重 + 重映射 y 偏移 ----
  const images = parseCellImages(raw.cellimages);
  const imgByName = new Map(images.map((im) => [im.name, im]));

  // 每张图片在整表源行里的首次出现行（锚点行）：同一产品多箱/跨组共用同一张图时按锚点定位
  const imageAnchorRow = new Map<string, number>();
  for (const rowXml of allRows) {
    const rM = /<row r="(\d+)"/.exec(rowXml);
    if (!rM) continue;
    const r = Number(rM[1]);
    if (r < DATA_START_ROW) continue;
    const disp = /DISPIMG\(&quot;(ID_[A-Za-z0-9]+)&quot;/.exec(rowXml);
    if (!disp) continue;
    if (!imageAnchorRow.has(disp[1])) imageAnchorRow.set(disp[1], r);
  }

  // 去重：同一图片 ID 只保留一次，锚到本组内首次出现的行（跨组借图则锚到本组首行）
  const seenImg = new Set<string>();
  const keptImages: { info: CellImageInfo; newY: number }[] = [];
  for (const rowXml of allRows) {
    const rM = /<row r="(\d+)"/.exec(rowXml);
    if (!rM) continue;
    const r = Number(rM[1]);
    if (!sourceRowSet.has(r)) continue;
    const disp = /DISPIMG\(&quot;(ID_[A-Za-z0-9]+)&quot;/.exec(rowXml);
    if (!disp) continue;
    const imgId = disp[1];
    if (seenImg.has(imgId)) continue;
    const info = imgByName.get(imgId);
    if (!info) continue;
    seenImg.add(imgId);
    const anchorRow = imageAnchorRow.get(imgId) ?? r;
    // 锚点行在本组内 → 映射到其新行；跨组借图 → 锚到本组首次出现的行
    const targetNewRow = rowMap.has(anchorRow) ? rowMap.get(anchorRow)! : rowMap.get(r)!;
    const newY = info.offY + (targetNewRow - anchorRow) * ROW_EMU;
    keptImages.push({ info, newY });
  }

  const keptCellImageXml = keptImages
    .map((k) => remapCellImageY(k.info.rawXml, k.newY))
    .join("");

  // 重建 cellimages.xml：保留命名空间头部 + 过滤后的 cellImage
  const nsHead = raw.cellimages.match(/<etc:cellImages[^>]*>/)?.[0] ?? "";
  const cellimagesXml = `${nsHead}${keptCellImageXml}</etc:cellImages>`;

  // cellimages.xml.rels：只保留被引用到的 rId（保持原 rId，不改编号）
  const keptRIds = new Set(keptImages.map((k) => k.info.rEmbed));
  const relItems = raw.cellimagesRels.match(/<Relationship[^>]*\/>/g) || [];
  const keptRelItems = relItems.filter((rel) => {
    const idM = /Id="([^"]+)"/.exec(rel);
    return idM && keptRIds.has(idM[1]);
  });
  const cellimagesRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    keptRelItems.join("") +
    `</Relationships>`;

  // 需要保留的媒体文件：logo(image1.png) + 被引用媒体
  const keptMediaPaths = new Set<string>(["xl/media/image1.png"]);
  for (const rel of keptRelItems) {
    const tM = /Target="([^"]+)"/.exec(rel);
    if (!tM) continue;
    const target = tM[1];
    if (target === "NULL") continue;
    keptMediaPaths.add(`xl/media/${target.replace(/^media\//, "")}`);
  }

  // ---- workbook / rels / contentTypes ----
  const workbookXml = raw.workbook
    .replace(/<sheet [^>]*r:id="rId2"[^>]*\/>/, "")
    .replace(/<sheet [^>]*r:id="rId3"[^>]*\/>/, "")
    .replace(/<definedNames>[\s\S]*?<\/definedNames>/, "");

  let workbookRelsXml = removeRelationship(raw.workbookRels, "worksheets/sheet2.xml");
  workbookRelsXml = removeRelationship(workbookRelsXml, "worksheets/sheet3.xml");
  workbookRelsXml = removeRelationship(workbookRelsXml, "../customXml/item1.xml");
  workbookRelsXml = removeRelationship(workbookRelsXml, "../customXml/item2.xml");

  let contentTypesXml = raw.contentTypes;
  for (const p of [
    "/customXml/itemProps1.xml",
    "/customXml/itemProps2.xml",
    "/xl/customProperty2.bin",
    "/xl/customProperty3.bin",
    "/xl/worksheets/sheet2.xml",
    "/xl/worksheets/sheet3.xml",
  ]) {
    contentTypesXml = contentTypesXml.replace(new RegExp(`<Override PartName="${p}"[^>]*/>`, "g"), "");
  }

  // ---- 组装 zip ----
  out.file("[Content_Types].xml", contentTypesXml);
  out.file("xl/workbook.xml", workbookXml);
  out.file("xl/_rels/workbook.xml.rels", workbookRelsXml);
  out.file("xl/worksheets/sheet1.xml", sheet1);
  out.file("xl/sharedStrings.xml", sse.getXml());
  out.file("xl/cellimages.xml", cellimagesXml);
  out.file("xl/_rels/cellimages.xml.rels", cellimagesRelsXml);

  for (const [p, b] of raw.staticFiles) out.file(p, b);
  for (const [p, b] of raw.media) if (keptMediaPaths.has(p)) out.file(p, b);

  return Buffer.from(await out.generateAsync({ type: "nodebuffer" }));
}

function removeRelationship(relsXml: string, target: string): string {
  return relsXml.replace(
    new RegExp(`<Relationship [^>]*Target="${target}"[^>]*/>`, "g"),
    ""
  );
}

function filterMergeCells(xml: string, rowMap: Map<number, number>): string {
  const block = /<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/.exec(xml);
  if (!block) return xml;
  const items = block[1].match(/<mergeCell[^>]*\/>/g) || [];
  const kept: string[] = [];
  for (const item of items) {
    const m = /ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/.exec(item);
    if (!m) {
      kept.push(item);
      continue;
    }
    const c1 = m[1];
    const n1 = Number(m[2]);
    const c2 = m[3];
    const n2 = Number(m[4]);
    if (n1 < DATA_START_ROW && n2 < DATA_START_ROW) {
      kept.push(item); // 表头合并
      continue;
    }
    const nn1 = rowMap.get(n1);
    const nn2 = rowMap.get(n2);
    if (nn1 !== undefined && nn2 !== undefined) {
      kept.push(`<mergeCell ref="${c1}${nn1}:${c2}${nn2}"/>`);
    }
    // 否则丢弃（明细合并两端不在本组）
  }
  const newBlock = kept.length ? `<mergeCells count="${kept.length}">${kept.join("")}</mergeCells>` : "";
  return xml.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, newBlock);
}

function filterDataValidations(xml: string): string {
  const block = /<dataValidations[^>]*>([\s\S]*?)<\/dataValidations>/.exec(xml);
  if (!block) return xml;
  const items = block[1].match(/<dataValidation[\s\S]*?<\/dataValidation>/g) || [];
  const kept = items.filter((item) => !/<formula1>[^<]*![^<]*<\/formula1>/.test(item));
  const newBlock = kept.length
    ? `<dataValidations count="${kept.length}">${kept.join("")}</dataValidations>`
    : "";
  return xml.replace(/<dataValidations[^>]*>[\s\S]*?<\/dataValidations>/, newBlock);
}

// ============================================================
// Main entry
// ============================================================

export async function splitInvoice(
  quotePath: string,
  invoicePath: string,
  invoiceName: string
): Promise<SiruiSplitResult> {
  const quoteMap = await parseQuote(quotePath);
  const { country, groups, unmatched } = await parseInvoice(invoicePath, quoteMap);

  if (groups.length === 0) throw new Error("发票里没有匹配到任何 FBA 明细行");

  const warnings: string[] = [];
  if (unmatched.length > 0) {
    warnings.push(
      `${unmatched.length} 行明细未匹配到报价表 FBA 码（已跳过）：` +
        unmatched.slice(0, 5).map((r) => r.boxId).join("、") +
        (unmatched.length > 5 ? "…" : "")
    );
  }

  const sourceZip = await JSZip.loadAsync(readFileSync(invoicePath));
  const raw = await readRawParts(sourceZip);

  const countryCn = toCnCountry(country);

  const files: SiruiFileEntry[] = [];
  const buffers: { fileName: string; buffer: Buffer }[] = [];

  for (const group of groups) {
    const currencyIso = toIsoCurrency(group.rows[0].currency);
    const fileName = `${group.fba}-${group.warehouse}.xlsx`;
    const buffer = await buildOutput(group, countryCn, currencyIso, raw);
    buffers.push({ fileName, buffer });
    files.push({ fba: group.fba, warehouse: group.warehouse, fileName, rowCount: group.rows.length });
  }

  const outZip = new JSZip();
  for (const b of buffers) outZip.file(b.fileName, b.buffer);
  const zipBuffer = Buffer.from(await outZip.generateAsync({ type: "nodebuffer" }));

  const baseName = (invoiceName || "思锐FBA下单发票").replace(/\.xlsx$/i, "");
  return {
    buffer: zipBuffer,
    zipFileName: `${baseName}-拆分结果.zip`,
    files,
    warnings,
    total: files.length,
  };
}
