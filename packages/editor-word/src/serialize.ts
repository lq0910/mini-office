import JSZip from 'jszip';
import {
  type DocBlock,
  type DocImage,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocStyleDefinition,
  type DocTable,
  type DocTableCell,
  type FieldMarker,
  type PageSetup,
  type ParagraphStyle,
  type RunStyle,
  type TableBorderSide,
  type TableBorders,
  type HeaderFooter,
} from '@mini-office/document-model';
import {
  escapeXml,
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  parseXmlDom,
  pxToTwips,
  resolvePartPath,
  xmlDeclaration,
  type OoxmlPackage,
} from '@mini-office/ooxml';
import { extensionOf, uid } from '@mini-office/shared';
import { buildWordMailMergeXml, fileUriTarget, MAIL_MERGE_DATA_REL, MAIL_MERGE_ODSO_REL, MAIL_MERGE_SOURCE_REL, type WordMailMergeInfo } from './mail-merge-ooxml';
import { documentSections } from './sections';
// 功能：保存时把文档用到的样式定义写入 styles.xml，Word/WPS 才能显示样式名与标题层级；时间：2026-10-29；作者：lq。
import { collectUsedStyleDefinitions, resolveStyleRun } from './styles';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
/** 功能：VML 命名空间声明——v:rect 浮层需要，缺省会让 Word 判定部件非法；时间：2026-10-30。 */
const VML_NS_DECL = ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"';
/** 功能：页眉页脚部件额外声明 w14（复选框控件需要），正文部件已在根元素声明；时间：2026-10-30。 */
const CONTROL_NS_DECL = ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

interface RelationshipInfo {
  id: string;
  type: string;
  target: string;
  /** 外部数据源关系使用 External，否则省略（内部关系默认 Internal） */
  targetMode?: 'External' | 'Internal';
}

interface SerializeState {
  zip: JSZip;
  sourceExists: boolean;
  documentRels: RelationshipInfo[];
  relsChanged: boolean;
  contentTypesXml?: string;
  contentTypesChanged: boolean;
  addedImages: Map<string, string>;
  imageCounter: number;
  /** 功能：图形对象 ID 必须全文档唯一，独立计数避免复用关系时产生重复 docPr id；时间：2026-10-29；作者：lq。 */
  drawingCounter: number;
  /** 功能：书签 id 计数器——写回书签时保证 id 唯一；时间：2026-10-30。 */
  bookmarkCounter: number;
  hasNumbering: boolean;
  sectPrXml?: string;
}

const IMAGE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
};

/**
 * 功能：把 #rgb / ARGB 归一化为 ST_HexColor 要求的 6 位十六进制，非法值返回 undefined；
 * Word 对 w:color、w:shd/@w:fill 的取值做严格校验，三位简写会让文档被判为非法；时间：2026-10-29；作者：lq。
 */
function hexColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const raw = value.replace('#', '').trim();
  if (/^[0-9a-fA-F]{6}$/.test(raw)) return raw.toUpperCase();
  if (/^[0-9a-fA-F]{8}$/.test(raw)) return raw.slice(2).toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(raw)) return raw.split('').map((char) => char + char).join('').toUpperCase();
  return undefined;
}

function attr(value: string | number | undefined, name: string): string {
  if (value === undefined || value === '') return '';
  return ` ${name}="${escapeXml(String(value))}"`;
}

function boolTag(tag: string, value: boolean | undefined, valueAttribute = true): string {
  if (value === undefined) return '';
  if (!value) return valueAttribute ? `<w:${tag} w:val="0"/>` : '';
  return valueAttribute ? `<w:${tag}/>` : `<w:${tag}/>`;
}

/**
 * 功能：按 ECMA-376 CT_RPr 的固定顺序输出字符属性，避免 Word 因元素顺序不符模式而拒绝打开；时间：2026-10-29；作者：lq。
 * 顺序：rStyle → rFonts → b → i → caps → smallCaps → strike → color → spacing → sz/szCs → u → shd → vertAlign。
 */
function runProperties(style: RunStyle): string {
  const parts: string[] = [];
  if (style.fontFamily || style.eastAsiaFont) {
    // 功能：中西文字体分槽写回（w:ascii / w:eastAsia），保存后 Word/WPS 显示不变；时间：2026-10-30。
    const latin = escapeXml(style.fontFamily ?? style.eastAsiaFont!);
    const eastAsia = escapeXml(style.eastAsiaFont ?? style.fontFamily!);
    parts.push(
      `<w:rFonts w:ascii="${latin}" w:hAnsi="${latin}" w:eastAsia="${eastAsia}" w:cs="${latin}"/>`,
    );
  }
  if (style.bold !== undefined) parts.push(`<w:b${style.bold ? '' : ' w:val="0"'}/>`);
  if (style.italic !== undefined) parts.push(`<w:i${style.italic ? '' : ' w:val="0"'}/>`);
  if (style.allCaps) parts.push('<w:caps/>');
  if (style.smallCaps) parts.push('<w:smallCaps/>');
  if (style.strike !== undefined) parts.push(`<w:strike${style.strike ? '' : ' w:val="0"'}/>`);
  const color = hexColor(style.color);
  if (color) parts.push(`<w:color w:val="${color}"/>`);
  if (style.letterSpacingPx) {
    parts.push(`<w:spacing w:val="${Math.round(style.letterSpacingPx * 15)}"/>`);
  }
  if (style.fontSizePt) {
    const halfPoints = Math.max(2, Math.round(style.fontSizePt * 2));
    parts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  if (style.underline !== undefined) {
    parts.push(`<w:u w:val="${style.underline ? 'single' : 'none'}"/>`);
  }
  const highlight = hexColor(style.highlight);
  if (highlight) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${highlight}"/>`);
  }
  if (style.verticalAlign === 'super') parts.push('<w:vertAlign w:val="superscript"/>');
  if (style.verticalAlign === 'sub') parts.push('<w:vertAlign w:val="subscript"/>');
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

/**
 * 功能：按 ECMA-376 CT_PPrBase 的固定顺序输出段落属性，避免 Word 因元素顺序不符模式而拒绝打开；时间：2026-10-29；作者：lq。
 * 顺序：pStyle → keepNext → pageBreakBefore → numPr → pBdr → shd → spacing → ind → jc → outlineLvl。
 */
function paragraphProperties(
  paragraph: DocParagraph,
  state: SerializeState | undefined,
  includeNumbering = true,
): string {
  const style = paragraph.style;
  const parts: string[] = [];
  if (paragraph.styleId) parts.push(`<w:pStyle w:val="${escapeXml(paragraph.styleId)}"/>`);
  if (style.keepNext) parts.push('<w:keepNext/>');
  if (style.pageBreakBefore) parts.push('<w:pageBreakBefore/>');
  if (includeNumbering && paragraph.list && state?.hasNumbering && paragraph.list.numId) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${paragraph.list.level}"/><w:numId w:val="${escapeXml(paragraph.list.numId)}"/></w:numPr>`,
    );
  }
  // 功能：保存段落边框，确保横线和签名线往返 DOCX 后仍存在；时间：2026-10-27；作者：lq。
  if (style.borders) {
    const content = ['top', 'left', 'bottom', 'right'].map((side) => borderXml(side, style.borders?.[side as keyof TableBorders])).join('');
    if (content) parts.push(`<w:pBdr>${content}</w:pBdr>`);
  }
  const background = hexColor(style.backgroundColor);
  if (background) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`);
  }
  if (style.tabStops?.length) {
    // 功能：写出制表位前导符（w:leader），Word/WPS 打开后目录仍显示「标题……页码」点线；时间：2026-10-30。
    const tabXml = style.tabStops
      .map(
        (t) =>
          `<w:tab w:val="${t.kind}" w:pos="${Math.round(t.posPx * 15)}"${
            t.leader && t.leader !== 'none' ? ` w:leader="${t.leader}"` : ''
          }/>`,
      )
      .join('');
    parts.push(`<w:tabs>${tabXml}</w:tabs>`);
  }
  const spacingParts: string[] = [];
  if (style.spaceBeforePx !== undefined) spacingParts.push(`w:before="${Math.round(style.spaceBeforePx * 15)}"`);
  if (style.spaceAfterPx !== undefined) spacingParts.push(`w:after="${Math.round(style.spaceAfterPx * 15)}"`);
  if (style.lineHeightExactPx !== undefined) {
    spacingParts.push(`w:line="${Math.round(style.lineHeightExactPx * 15)}"`);
    spacingParts.push('w:lineRule="exact"');
  } else if (style.lineHeightAtLeastPx !== undefined) {
    // 功能：最小值行距写回 atLeast，Word 打开后仍可继续调整；时间：2026-10-29；作者：lq。
    spacingParts.push(`w:line="${Math.round(style.lineHeightAtLeastPx * 15)}"`);
    spacingParts.push('w:lineRule="atLeast"');
  } else if (style.lineHeight !== undefined) {
    spacingParts.push(`w:line="${Math.round(style.lineHeight * 240)}"`);
    spacingParts.push('w:lineRule="auto"');
  }
  if (spacingParts.length) parts.push(`<w:spacing ${spacingParts.join(' ')}/>`);
  const indentParts: string[] = [];
  if (style.indentLeftPx) indentParts.push(`w:left="${Math.round(style.indentLeftPx * 15)}"`);
  if (style.indentRightPx) indentParts.push(`w:right="${Math.round(style.indentRightPx * 15)}"`);
  if (style.firstLineIndentPx !== undefined) {
    if (style.firstLineIndentPx >= 0) {
      indentParts.push(`w:firstLine="${Math.round(style.firstLineIndentPx * 15)}"`);
    } else {
      indentParts.push(`w:hanging="${Math.round(Math.abs(style.firstLineIndentPx) * 15)}"`);
    }
  }
  if (indentParts.length) parts.push(`<w:ind ${indentParts.join(' ')}/>`);
  if (style.alignment) {
    const map: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
    parts.push(`<w:jc w:val="${map[style.alignment] ?? 'left'}"/>`);
  }
  if (style.outlineLevel !== undefined) parts.push(`<w:outlineLvl w:val="${style.outlineLevel}"/>`);
  // 功能：段落关闭「对齐到网格」时写回 w:snapToGrid=0，Word 打开后行高不再按网格取整；时间：2026-10-30。
  if (style.snapToGrid === false) parts.push('<w:snapToGrid w:val="0"/>');
  return parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : '';
}

function borderXml(tag: string, border: TableBorderSide | undefined): string {
  if (!border) return '';
  if (border.style === 'none' || border.widthPx === 0) return `<w:${tag} w:val="nil"/>`;
  const top = Math.max(2, Math.round(border.widthPx * 6));
  const color = hexColor(border.color) ?? '000000';
  return `<w:${tag} w:val="${border.style === 'double' ? 'double' : border.style}" w:sz="${top}" w:space="0" w:color="${color}"/>`;
}

function tableBordersXml(borders: TableBorders | undefined): string {
  if (!borders) return '';
  const content = [
    borderXml('top', borders.top),
    borderXml('left', borders.left),
    borderXml('bottom', borders.bottom),
    borderXml('right', borders.right),
    borderXml('insideH', borders.insideH),
    borderXml('insideV', borders.insideV),
  ].join('');
  return content ? `<w:tblBorders>${content}</w:tblBorders>` : '';
}

function cellBordersXml(borders: TableBorders | undefined): string {
  if (!borders) return '';
  const content = [
    borderXml('top', borders.top),
    borderXml('left', borders.left),
    borderXml('bottom', borders.bottom),
    borderXml('right', borders.right),
  ].join('');
  return content ? `<w:tcBorders>${content}</w:tcBorders>` : '';
}

function dataUrlToBytes(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(dataUrl);
  if (!match) return null;
  const mime = match[1] || 'application/octet-stream';
  const isBase64 = !!match[2];
  try {
    if (isBase64) {
      const binary = atob(match[3]);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
      return { bytes, mime };
    }
    const decoded = decodeURIComponent(match[3]);
    const bytes = new TextEncoder().encode(decoded);
    return { bytes, mime };
  } catch {
    return null;
  }
}

async function blobUrlToBytes(url: string): Promise<{ bytes: Uint8Array; mime: string } | null> {
  try {
    const response = await fetch(url);
    const blob = await response.blob();
    const buffer = await blob.arrayBuffer();
    return { bytes: new Uint8Array(buffer), mime: blob.type || 'image/png' };
  } catch {
    return null;
  }
}

function existingRels(xml: string | undefined): RelationshipInfo[] {
  if (!xml) return [];
  try {
    const doc = parseXmlDom(xml);
    return Array.from(doc.getElementsByTagName('Relationship')).map((element) => ({
      id: getAttrLocal(element, 'Id') ?? '',
      type: getAttrLocal(element, 'Type') ?? '',
      target: getAttrLocal(element, 'Target') ?? '',
      targetMode: getAttrLocal(element, 'TargetMode') === 'External' ? 'External' : undefined,
    }));
  } catch {
    return [];
  }
}

async function ensureImageRelationship(
  state: SerializeState,
  image: DocImage,
): Promise<string | undefined> {
  const cacheKey = image.id ?? image.relId ?? image.part ?? image.dataUrl?.slice(0, 80) ?? image.src.slice(0, 120);
  const cached = state.addedImages.get(cacheKey);
  if (cached) return cached;

  if (image.relId) {
    const existing = state.documentRels.find((rel) => rel.id === image.relId);
    if (existing) {
      state.addedImages.set(cacheKey, existing.id);
      return existing.id;
    }
  }

  if (image.part) {
    const target = resolvePartPath('word/document.xml', image.part) ?? image.part;
    const match = state.documentRels.find((rel) => {
      const resolved = resolvePartPath('word/document.xml', rel.target) ?? rel.target;
      return resolved === target;
    });
    if (match) {
      state.addedImages.set(cacheKey, match.id);
      return match.id;
    }
  }

  let payload: { bytes: Uint8Array; mime: string } | null = null;
  if (image.dataUrl) payload = dataUrlToBytes(image.dataUrl);
  else if (image.src.startsWith('blob:') || image.src.startsWith('http')) {
    payload = await blobUrlToBytes(image.src);
  }
  if (!payload) return undefined;

  const ext =
    IMAGE_MIME_EXT[payload.mime] ??
    (extensionOf(image.part ?? '').replace('.', '') || 'png');
  state.imageCounter += 1;
  const fileName = `mini-office-${state.imageCounter}.${ext}`;
  const mediaPath = `word/media/${fileName}`;
  state.zip.file(mediaPath, payload.bytes);
  const relId = `rIdMo${state.imageCounter}`;
  state.documentRels.push({
    id: relId,
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    target: `media/${fileName}`,
  });
  state.relsChanged = true;
  state.contentTypesChanged = true;
  state.addedImages.set(cacheKey, relId);
  return relId;
}

function drawingXml(image: DocImage, relId: string, drawingId: number): string {
  const widthEmu = Math.round(image.widthPx * 9525);
  const heightEmu = Math.round(image.heightPx * 9525);
  const name = escapeXml(image.name ?? `图片 ${drawingId}`);
  const descr = escapeXml(image.alt ?? image.name ?? '');
  const geometry = `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`;
  const frame = `<wp:docPr id="${drawingId}" name="${name}" descr="${descr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${A_NS}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${PIC_NS}"><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="${name}" descr="${descr}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const wrap = image.wrap ?? 'inline';
  // 功能：浮动图的环绕方式必须排在 wp:docPr 之前（CT_Anchor 顺序），否则 Word 会拒绝打开含浮动图的文档；时间：2026-10-29；作者：lq。
  const drawing = wrap === 'inline'
    ? `<wp:inline distT="0" distB="0" distL="0" distR="0">${geometry}${frame}</wp:inline>`
    : `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>${geometry}${wrap === 'top-bottom' ? '<wp:wrapTopAndBottom/>' : `<wp:wrapSquare wrapText="${wrap === 'square-right' ? 'right' : 'left'}"/>`}${frame}</wp:anchor>`;
  return `<w:drawing>${drawing}</w:drawing>`;
}

/**
 * 功能：把内容控件写回标准 w:sdt——复选框用 w14:checkbox 记录勾选，下拉/组合框写候选，日期写格式；
 * 子元素顺序遵循 CT_SdtPr（alias/tag/lock 在前，控件类型在后），否则 Word 会判定部件非法；时间：2026-10-30。
 */
function sdtPrXml(control: NonNullable<DocRun['control']>): string {
  const parts: string[] = [];
  if (control.alias) parts.push(`<w:alias w:val="${escapeXml(control.alias)}"/>`);
  if (control.tag) parts.push(`<w:tag w:val="${escapeXml(control.tag)}"/>`);
  if (control.locked) parts.push('<w:lock w:val="sdtContentLocked"/>');
  if (control.kind === 'checkbox') {
    parts.push(`<w14:checkbox><w14:checked w14:val="${control.checked ? 1 : 0}"/></w14:checkbox>`);
  } else if (control.kind === 'dropdown' || control.kind === 'comboBox') {
    const tag = control.kind === 'dropdown' ? 'w:dropDownList' : 'w:comboBox';
    const items = (control.choices ?? [])
      .map((choice) => `<w:listItem w:displayText="${escapeXml(choice)}" w:value="${escapeXml(choice)}"/>`)
      .join('');
    parts.push(`<${tag}>${items}</${tag}>`);
  } else if (control.kind === 'date') {
    parts.push(`<w:date><w:dateFormat w:val="${escapeXml(control.dateFormat ?? 'yyyy年M月d日')}"/><w:lid w:val="zh-CN"/></w:date>`);
  } else {
    parts.push('<w:text/>');
  }
  return `<w:sdtPr>${parts.join('')}</w:sdtPr>`;
}

async function runXml(run: DocRun, state: SerializeState): Promise<string> {
  // 功能：文本框写回为 VML v:shape + v:textbox，Word/WPS 打开仍是文本框；时间：2026-10-31。
  if (run.textbox) {
    const box = run.textbox;
    const style = [
      'position:absolute',
      `margin-left:${(box.left * 0.75).toFixed(2)}pt`,
      `margin-top:${(box.top * 0.75).toFixed(2)}pt`,
      `width:${(box.width * 0.75).toFixed(2)}pt`,
      `height:${(box.height * 0.75).toFixed(2)}pt`,
      box.xRelative === 'page' ? 'mso-position-horizontal-relative:page' : 'mso-position-horizontal-relative:text',
      'mso-position-vertical-relative:paragraph',
    ].join(';');
    let body = '';
    for (const block of box.blocks) body += block.type === 'paragraph' ? await paragraphXml(block, state) : '';
    return `<w:r><w:pict><v:shape id="_x0000_s${Math.floor(Math.random() * 100000)}" type="#_x0000_t202" style="${style}" filled="t" stroked="t"><v:textbox><w:txbxContent>${body}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
  }
  // 功能：VML 浮层（复选框等）按原样式写回，保证另存后在 Word/WPS 中位置与描边不变；时间：2026-10-30。
  if (run.image?.vmlRect) {
    const rect = run.image.vmlRect;
    const stroke = rect ? (run.image.strokeWidthPx ?? 1) : 1;
    const color = run.image.strokeColor ?? '#000000';
    const style = [
      'position:absolute',
      `margin-left:${(rect.left * 0.75).toFixed(2)}pt`,
      `margin-top:${(rect.top * 0.75).toFixed(2)}pt`,
      `width:${(rect.width * 0.75).toFixed(2)}pt`,
      `height:${(rect.height * 0.75).toFixed(2)}pt`,
      `mso-position-horizontal-relative:${rect.xRelative === 'page' ? 'page' : 'text'}`,
      'mso-position-vertical-relative:paragraph',
    ].join(';');
    const strokeweight = `${(stroke * 0.75).toFixed(2)}pt`;
    // 功能：勾选状态写回 v:textbox 里的 ✓（Word 里显示为框内打勾），重新打开可再读回；时间：2026-10-30。
    const textbox = rect.checked
      ? '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>✓</w:t></w:r></w:p></w:txbxContent></v:textbox>'
      : '';
    const openTag = `<v:rect style="${style}" filled="false" stroked="true" strokecolor="${color}" strokeweight="${strokeweight}"`;
    const rectXml = textbox ? `${openTag}>${textbox}</v:rect>` : `${openTag}/>`;
    return `<w:r><w:pict>${rectXml}</w:pict></w:r>`;
  }
  if (run.image) {
    const relId = await ensureImageRelationship(state, run.image);
    if (!relId) return '';
    const properties = runProperties(run.style);
    state.drawingCounter += 1;
    return `<w:r>${properties}${drawingXml(run.image, relId, state.drawingCounter)}</w:r>`;
  }
  const properties = runProperties(run.style);
  let content = '';
  if (run.tab) content += '<w:tab/>';
  // 功能：删除修订的文本写回为 w:delText；时间：2026-10-31。
  // 功能：分栏符写回 w:br w:type="column"，多栏节另存后栏序不变；时间：2026-10-30。
  if (run.columnBreak) content += '<w:br w:type="column"/>';
  // 功能：段内分页符写回 w:br w:type="page"，保证 docx 往返不丢分页位置；时间：2026-10-30；作者：lq。
  if (run.pageBreak) content += '<w:br w:type="page"/>';
  else if (run.lineBreak) content += '<w:br/>';
  if (run.text) content += run.revision?.type === 'delete' ? `<w:delText xml:space="preserve">${escapeXml(run.text)}</w:delText>` : `<w:t xml:space="preserve">${escapeXml(run.text)}</w:t>`;
  if (!content && !properties) content = '<w:t xml:space="preserve"></w:t>';
  const core = `<w:r>${properties}${content}</w:r>`;
  if (run.revision) {
    // 功能：修订写回 w:ins/w:del（作者与时间保留，Word 打开仍是待审修订）；时间：2026-10-31。
    const attrs = `${run.revision.author ? ` w:author="${escapeXml(run.revision.author)}"` : ''}${run.revision.date ? ` w:date="${escapeXml(run.revision.date)}"` : ''}`;
    if (run.revision.type === 'insert') return `<w:ins${attrs}>${core}</w:ins>`;
    if (run.revision.type === 'delete') return `<w:del${attrs}>${core}</w:del>`;
  }
  // 功能：模板字段保存为原生 MERGEFIELD，字段名安全转义；格式串仅由外层 instr 统一转义，避免 &amp;quot; 双重转义；时间：2026-10-24；作者：lq。
  if (run.mergeField) {
    const name = run.mergeField.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const f = run.mergeFormat ?? {};
    let instr = ` MERGEFIELD "${name}"`;
    if (f.num) instr += ` \\# "${f.num}"`;
    if (f.date) instr += ` \\@ "${f.date}"`;
    if (f.case) instr += ` \\* ${f.case}`;
    if (f.before) instr += ` \\b "${f.before}"`;
    if (f.after) instr += ` \\f "${f.after}"`;
    instr += ` \\* MERGEFORMAT `;
    return `<w:fldSimple w:instr="${escapeXml(instr)}">${core}</w:fldSimple>`;
  }
  if (run.mergeNext) {
    return `<w:fldSimple w:instr=" NEXT \* MERGEFORMAT ">${core}</w:fldSimple>`;
  }
  if (run.mergeNextIf) {
    return `<w:fldSimple w:instr="${escapeXml(` NEXTIF ${run.mergeNextIf} \* MERGEFORMAT `)}">${core}</w:fldSimple>`;
  }
  if (run.mergeSkipIf) {
    return `<w:fldSimple w:instr="${escapeXml(` SKIPIF ${run.mergeSkipIf} \* MERGEFORMAT `)}">${core}</w:fldSimple>`;
  }
  if (run.mergeRecord) {
    return `<w:fldSimple w:instr=" MERGEREC \* MERGEFORMAT ">${core}</w:fldSimple>`;
  }
  if (run.mergeSeq) {
    return `<w:fldSimple w:instr=" MERGESEQ \* MERGEFORMAT ">${core}</w:fldSimple>`;
  }
  if (run.field === 'PAGE' || run.field === 'NUMPAGES' || run.field === 'SECTIONPAGES') {
    const instruction = ` ${run.field} \\* MERGEFORMAT `;
    return `<w:fldSimple w:instr="${instruction}">${core}</w:fldSimple>`;
  }
  // 功能：内容控件包住该 run 的内容，另存后 Word 里仍是可勾选/可选控件；复选框用 ☑/☐ 字符呈现勾选态；时间：2026-10-30。
  if (run.control) {
    const inner = run.control.kind === 'checkbox'
      ? `<w:r>${properties}<w:t>${run.control.checked ? '☑' : '☐'}</w:t></w:r>`
      : core;
    return `<w:sdt>${sdtPrXml(run.control)}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  }
  return core;
}

/**
 * 功能：复杂域标记写回为 fldChar/instrText——保住 TOC 等域结构，Word 打开后仍可"更新目录"；
 * 时间：2026-10-30。
 */
function fieldMarkerXml(marker: FieldMarker): string {
  if (marker.role === 'begin') {
    const instr = marker.instr?.trim();
    const instruction = instr
      ? `<w:r><w:instrText xml:space="preserve">${escapeXml(instr)}</w:instrText></w:r>`
      : '';
    return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>${instruction}`;
  }
  if (marker.role === 'separate') return `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  return `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

async function paragraphXml(
  paragraph: DocParagraph,
  state: SerializeState,
  options: { isInTable?: boolean; forceNumberingAsText?: boolean } = {},
): Promise<string> {
  const includeNumbering = !options.forceNumberingAsText;
  const pPr = paragraphProperties(paragraph, state, includeNumbering);
  let prefix = '';
  if (
    paragraph.list &&
    (!state.hasNumbering || !paragraph.list.numId || options.forceNumberingAsText)
  ) {
    prefix = `<w:r><w:t xml:space="preserve">${escapeXml(paragraph.list.marker)} </w:t></w:r>`;
  }
  const runs = await Promise.all(paragraph.runs.map((run) => runXml(run, state)));
  // 功能：段首分栏符写回为前缀 run（不插入 runs 数组，避免打乱域标记的 run 序号）；时间：2026-10-30。
  if (paragraph.style.columnBreakBefore) prefix = `<w:r><w:br w:type="column"/></w:r>${prefix}`;
  // 功能：写回段落书签（目录项的锚点目标），保存后 Word 里目录才能真正跳转；时间：2026-10-30。
  const bookmarkNames = [...new Set(paragraph.bookmarks ?? [])];
  const bookmarkIds: number[] = [];
  let bookmarkStartXml = '';
  for (const name of bookmarkNames) {
    state.bookmarkCounter += 1;
    bookmarkIds.push(state.bookmarkCounter);
    bookmarkStartXml += `<w:bookmarkStart w:id="${state.bookmarkCounter}" w:name="${escapeXml(name)}"/>`;
  }
  const bookmarkEndXml = bookmarkIds.map((id) => `<w:bookmarkEnd w:id="${id}"/>`).join('');
  // 功能：域标记按段内位置插回 run 之间，保证域结构与条目顺序一致；时间：2026-10-30。
  const markers = (paragraph.fieldMarkers ?? [])
    .filter((marker) => !marker.skip)
    .slice()
    .sort((a, b) => a.afterRun - b.afterRun);
  const parts: string[] = [];
  let groupAnchor: string | undefined;
  let group: string[] = [];
  const flushGroup = () => {
    if (!group.length) return;
    const xml = group.join('');
    // 功能：目录项等内部书签链接写回 w:hyperlink，保存后 Word 仍可点击跳转；时间：2026-10-30。
    parts.push(
      groupAnchor
        ? `<w:hyperlink w:anchor="${escapeXml(groupAnchor)}" w:history="1">${xml}</w:hyperlink>`
        : xml,
    );
    group = [];
  };
  // 功能：批注范围标记——首个携带 id 的 run 前开范围，末个 run 后关范围并挂引用；时间：2026-10-31。
  const openComments = new Set<string>();
  const lastRunWith = new Map<string, number>();
  paragraph.runs.forEach((run, index) => {
    for (const id of run.commentIds ?? []) lastRunWith.set(id, index);
  });
  for (let index = 0; index <= runs.length; index += 1) {
    for (const marker of markers) {
      if (marker.afterRun !== index) continue;
      flushGroup();
      parts.push(fieldMarkerXml(marker));
    }
    if (index === runs.length) break;
    const run = paragraph.runs[index];
    for (const id of run.commentIds ?? []) {
      if (!openComments.has(id)) {
        flushGroup();
        parts.push(`<w:commentRangeStart w:id="${escapeXml(id)}"/>`);
        openComments.add(id);
      }
    }
    for (const id of [...openComments]) {
      if ((run.commentIds ?? []).includes(id)) continue;
      flushGroup();
      parts.push(`<w:commentRangeEnd w:id="${escapeXml(id)}"/><w:r><w:commentReference w:id="${escapeXml(id)}"/></w:r>`);
      openComments.delete(id);
    }
    if (group.length && run.anchor !== groupAnchor) flushGroup();
    groupAnchor = run.anchor;
    group.push(runs[index]);
  }
  for (const id of openComments) {
    flushGroup();
    parts.push(`<w:commentRangeEnd w:id="${escapeXml(id)}"/><w:r><w:commentReference w:id="${escapeXml(id)}"/></w:r>`);
  }
  flushGroup();
  return `<w:p>${pPr}${prefix}${bookmarkStartXml}${parts.join('')}${bookmarkEndXml}</w:p>`;
}

async function cellXml(
  cell: DocTableCell,
  state: SerializeState,
  columnWidthPx: number | undefined,
  continuation = false,
): Promise<string> {
  const props: string[] = [];
  // 功能：以网格列宽优先保存合并区域总宽度，避免旧 tcW 覆盖尺寸调整；时间：2026-10-19；作者：lq。
  const width = columnWidthPx ?? cell.widthPx;
  if (width) props.push(`<w:tcW w:w="${Math.round(width * 15)}" w:type="dxa"/>`);
  if (cell.colSpan > 1) props.push(`<w:gridSpan w:val="${cell.colSpan}"/>`);
  if (continuation) props.push('<w:vMerge/>');
  else if (cell.rowSpan > 1) props.push('<w:vMerge w:val="restart"/>');
  // 功能：按 CT_TcPr 顺序输出（tcW → gridSpan → vMerge → tcBorders → shd → vAlign），避免 Word 判定部件非法；时间：2026-10-29；作者：lq。
  if (cell.borders) {
    const tcBorders = cellBordersXml(cell.borders);
    if (tcBorders) props.push(tcBorders);
  }
  const background = hexColor(cell.backgroundColor);
  if (background) {
    props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`);
  }
  if (cell.verticalAlign) {
    const map = { top: 'top', middle: 'center', bottom: 'bottom' } as const;
    props.push(`<w:vAlign w:val="${map[cell.verticalAlign]}"/>`);
  }
  const content = continuation
    ? ''
    : (await Promise.all(cell.blocks.map((block) => blockXml(block, state)))).join('');
  return `<w:tc><w:tcPr>${props.join('')}</w:tcPr>${content || '<w:p/>'}</w:tc>`;
}

async function tableXml(table: DocTable, state: SerializeState): Promise<string> {
  const props: string[] = [];
  if (table.styleId) props.push(`<w:tblStyle w:val="${escapeXml(table.styleId)}"/>`);
  if (table.widthPct) {
    props.push(`<w:tblW w:w="${Math.round(table.widthPct * 50)}" w:type="pct"/>`);
  } else {
    props.push('<w:tblW w:w="0" w:type="auto"/>');
  }
  if (table.alignment) {
    const map = { left: 'left', center: 'center', right: 'right', justify: 'left' } as const;
    props.push(`<w:jc w:val="${map[table.alignment]}"/>`);
  }
  const borders = tableBordersXml(table.borders);
  if (borders) props.push(borders);
  // 功能：固定网格布局确保 Office 按用户设置的列宽打开表格；时间：2026-10-19；作者：lq。
  // 功能：tblLayout 必须位于 tblBorders 之后（CT_TblPr 顺序），否则 Word 会判定表格部件非法；时间：2026-10-29；作者：lq。
  props.push('<w:tblLayout w:type="fixed"/>');
  const grid = table.columnWidthsPx
    .map((width) => `<w:gridCol w:w="${Math.round(width * 15)}"/>`)
    .join('');
  const columnCount = Math.max(
    table.columnWidthsPx.length,
    ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + cell.colSpan, 0)),
  );
  const active: Array<{ start: number; span: number; remaining: number }> = [];
  const rows: string[] = [];
  for (const row of table.rows) {
    const trProps: string[] = [];
    if (row.heightPx) trProps.push(`<w:trHeight w:val="${Math.round(row.heightPx * 15)}" w:hRule="atLeast"/>`);
    if (row.isHeader) trProps.push('<w:tblHeader/>');
    const cellsXml: string[] = [];
    let column = 0;
    let cellIndex = 0;
    while (column < columnCount || cellIndex < row.cells.length) {
      const spanning = active.find((item) => item.start === column && item.remaining > 0);
      if (spanning) {
        cellsXml.push(
          await cellXml(
            {
              id: `continue-${column}`,
              blocks: [],
              colSpan: spanning.span,
              rowSpan: 1,
            },
            state,
            table.columnWidthsPx.slice(column, column + spanning.span).reduce((sum, width) => sum + width, 0),
            true,
          ),
        );
        spanning.remaining -= 1;
        column += spanning.span;
        continue;
      }
      const cell = row.cells[cellIndex];
      if (!cell) break;
      cellsXml.push(await cellXml(cell, state, table.columnWidthsPx.slice(column, column + cell.colSpan).reduce((sum, width) => sum + width, 0)));
      if (cell.rowSpan > 1) {
        active.push({ start: column, span: cell.colSpan, remaining: cell.rowSpan - 1 });
      }
      column += cell.colSpan;
      cellIndex += 1;
    }
    for (const item of active) {
      if (item.remaining <= 0) item.remaining = 0;
    }
    rows.push(`<w:tr>${trProps.length ? `<w:trPr>${trProps.join('')}</w:trPr>` : ''}${cellsXml.join('')}</w:tr>`);
  }
  return `<w:tbl><w:tblPr>${props.join('')}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>`;
}

async function blockXml(block: DocBlock, state: SerializeState): Promise<string> {
  if (block.type === 'paragraph') return paragraphXml(block, state);
  if (block.type === 'table') return tableXml(block, state);
  // 功能：分页标记只占最小行高，保持长表分页与预览一致；时间：2026-10-24；作者：lq。
  return '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>';
}

/**
 * 功能：ECMA-376 对 w:sectPr 子元素规定了固定顺序，插入时必须按序落位，否则 Word 会判定部件非法；时间：2026-10-29；作者：lq。
 */
const SECT_PR_ORDER = [
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
];

/** 功能：把节点插入到模式允许的位置，重复元素（页眉页脚引用）排在同类之后；时间：2026-10-29；作者：lq。 */
function insertInSequence(parent: Element, child: Element, order: string[]): void {
  const target = order.indexOf(child.localName ?? '');
  let reference: Element | null = null;
  for (const existing of Array.from(parent.children)) {
    if (order.indexOf(existing.localName ?? '') > target) {
      reference = existing;
      break;
    }
  }
  parent.insertBefore(child, reference);
}

/** 功能：生成 w:cols——不等宽栏写出每栏 w:w/w:space 并标记 equalWidth=0；时间：2026-10-30。 */
function columnsXml(page: PageSetup): string {
  if (!page.columns || page.columns <= 1) return '';
  const widths = page.columnWidthsPx;
  if (!widths?.length || widths.length !== page.columns) {
    return `<w:cols w:num="${page.columns}"/>`;
  }
  const cols = widths
    .map((width, index) => {
      const space = page.columnSpacePx?.[index];
      return `<w:col w:w="${Math.round(pxToTwips(width))}"${space !== undefined ? ` w:space="${Math.round(pxToTwips(space))}"` : ''}/>`;
    })
    .join('');
  return `<w:cols w:num="${page.columns}" w:equalWidth="0">${cols}</w:cols>`;
}

function buildSectPr(page: PageSetup, state: SerializeState): string {
  const width = Math.round(pxToTwips(page.widthPx));
  const height = Math.round(pxToTwips(page.heightPx));
  const margin = page.margins;
  const pgSz = `<w:pgSz w:w="${width}" w:h="${height}"${page.orientation === 'landscape' ? ' w:orient="landscape"' : ''}/>`;
  const pgMar = `<w:pgMar w:top="${Math.round(pxToTwips(margin.top))}" w:right="${Math.round(pxToTwips(margin.right))}" w:bottom="${Math.round(pxToTwips(margin.bottom))}" w:left="${Math.round(pxToTwips(margin.left))}" w:header="${Math.round(pxToTwips(margin.header))}" w:footer="${Math.round(pxToTwips(margin.footer))}" w:gutter="${Math.round(pxToTwips(margin.gutter ?? 0))}"/>`;
  if (state.sectPrXml) {
    try {
      const doc = parseXmlDom(
        `<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${state.sectPrXml}</w:root>`,
      );
      const sectPr = firstDescendant(doc, 'w:sectPr');
      if (sectPr) {
        const size = firstChild(sectPr, 'w:pgSz') ?? doc.createElementNS(W_NS, 'w:pgSz');
        size.setAttribute('w:w', String(width));
        size.setAttribute('w:h', String(height));
        if (page.orientation === 'landscape') size.setAttribute('w:orient', 'landscape');
        else size.removeAttribute('w:orient');
        if (!size.parentNode) insertInSequence(sectPr, size, SECT_PR_ORDER);
        const margins = firstChild(sectPr, 'w:pgMar') ?? doc.createElementNS(W_NS, 'w:pgMar');
        margins.setAttribute('w:top', String(Math.round(pxToTwips(margin.top))));
        margins.setAttribute('w:right', String(Math.round(pxToTwips(margin.right))));
        margins.setAttribute('w:bottom', String(Math.round(pxToTwips(margin.bottom))));
        margins.setAttribute('w:left', String(Math.round(pxToTwips(margin.left))));
        margins.setAttribute('w:header', String(Math.round(pxToTwips(margin.header))));
        margins.setAttribute('w:footer', String(Math.round(pxToTwips(margin.footer))));
        margins.setAttribute('w:gutter', String(Math.round(pxToTwips(margin.gutter ?? 0))));
        if (!margins.parentNode) insertInSequence(sectPr, margins, SECT_PR_ORDER);
        // 功能：写回起始页码并保留其他页码格式属性；时间：2026-10-24；作者：lq。
        if (page.pageNumberStart !== undefined) {
          const numbering = firstChild(sectPr, 'w:pgNumType') ?? doc.createElementNS(W_NS, 'w:pgNumType');
          numbering.setAttribute('w:start', String(page.pageNumberStart));
          if (!numbering.parentNode) insertInSequence(sectPr, numbering, SECT_PR_ORDER);
        }
        // 功能：把栏数/栏宽、分节类型与网格行距同步进原有 sectPr，避免改版后写回丢失；时间：2026-10-30。
        const typeElement = firstChild(sectPr, 'w:type');
        if (page.type) {
          const typeNode = typeElement ?? doc.createElementNS(W_NS, 'w:type');
          typeNode.setAttribute('w:val', page.type);
          if (!typeNode.parentNode) insertInSequence(sectPr, typeNode, SECT_PR_ORDER);
        }
        const existingCols = firstChild(sectPr, 'w:cols');
        const freshCols = page.columns && page.columns > 1
          ? firstDescendant(parseXmlDom(`<w:root xmlns:w="${W_NS}">${columnsXml(page)}</w:root>`), 'w:cols')
          : null;
        if (existingCols && freshCols) {
          const parent = existingCols.parentNode!;
          parent.replaceChild(doc.importNode(freshCols, true), existingCols);
        } else if (existingCols) {
          existingCols.parentNode?.removeChild(existingCols);
        } else if (freshCols) {
          insertInSequence(sectPr, doc.importNode(freshCols, true), SECT_PR_ORDER);
        }
        if (page.linePitchPx !== undefined) {
          const grid = firstChild(sectPr, 'w:docGrid') ?? doc.createElementNS(W_NS, 'w:docGrid');
          grid.setAttribute('w:linePitch', String(Math.round(pxToTwips(page.linePitchPx))));
          if (!grid.parentNode) insertInSequence(sectPr, grid, SECT_PR_ORDER);
        }
        // 功能：首页不同写回 w:titlePg；时间：2026-10-31。
        if (page.titlePage && !firstChild(sectPr, 'w:titlePg')) {
          insertInSequence(sectPr, doc.createElementNS(W_NS, 'w:titlePg'), SECT_PR_ORDER);
        }
        return new XMLSerializer().serializeToString(sectPr);
      }
    } catch {
      // Fall back to a generated section.
    }
  }
  // 功能：写回分节类型、栏宽间距与文档网格，保证不等宽多栏与网格行距在 Word/WPS 中保持原样；时间：2026-10-30。
  const sectType = `<w:type w:val="${page.type ?? 'nextPage'}"/>`;
  const titlePgXml = page.titlePage ? '<w:titlePg/>' : '';
  const colsXml = columnsXml(page);
  const docGridXml = page.linePitchPx !== undefined
    ? `<w:docGrid w:type="lines" w:linePitch="${Math.round(pxToTwips(page.linePitchPx))}"/>`
    : '';
  return `<w:sectPr>${sectType}${pgSz}${pgMar}${titlePgXml}${page.pageNumberStart !== undefined ? `<w:pgNumType w:start="${page.pageNumberStart}"/>` : ''}${colsXml}${docGridXml}</w:sectPr>`;
}

/** 功能：保存新增/编辑后的默认页眉页脚，写入独立部件、关系和内容类型；时间：2026-10-24；作者：lq。 */
async function writeHeaderFooters(model: DocModel, state: SerializeState, suffix = ''): Promise<void> {
  const sectionDoc = parseXmlDom(`<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${state.sectPrXml ?? buildSectPr(model.page, state)}</w:root>`);
  const section = firstDescendant(sectionDoc, 'w:sectPr')!;
  // 功能：支持首页/奇偶页页眉页脚变体——独立部件文件与对应 type 的引用；时间：2026-10-31。
  const variantEntries: Array<{ kind: 'header' | 'footer'; variant: 'default' | 'first' | 'even'; value?: HeaderFooter }> = [
    { kind: 'header', variant: 'default', value: model.header },
    { kind: 'footer', variant: 'default', value: model.footer },
    { kind: 'header', variant: 'first', value: model.headerFirst },
    { kind: 'header', variant: 'even', value: model.headerEven },
    { kind: 'footer', variant: 'first', value: model.footerFirst },
    { kind: 'footer', variant: 'even', value: model.footerEven },
  ];
  for (const { kind, variant, value } of variantEntries) {
    if (!value || (state.sourceExists && !value.edited)) continue;
    // 每个部件拥有独立关系作用域；新建文本内容不复用原图片关系；时间：2026-10-24；作者：lq。
    const variantSuffix = variant === 'default' ? '' : `-${variant}`;
    const target = `mini-${kind}${variantSuffix}${suffix}.xml`;
    const part = `word/${target}`;
    const sourceRelPath = value.sourcePart ? `word/_rels/${value.sourcePart.split('/').at(-1)}.rels` : '';
    const partRels = sourceRelPath ? await state.zip.file(sourceRelPath)?.async('text') : undefined;
    const partState: SerializeState = { ...state, documentRels: existingRels(partRels), relsChanged: false, addedImages: new Map() };
    let body = '';
    for (const block of value.blocks) body += await blockXml(block, partState);
    state.imageCounter = partState.imageCounter;
    state.drawingCounter = partState.drawingCounter;
    const root = kind === 'header' ? 'hdr' : 'ftr';
    state.zip.file(part, `${xmlDeclaration()}<w:${root} xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"${VML_NS_DECL}${CONTROL_NS_DECL}>${body || '<w:p/>'}</w:${root}>`);
    if (partState.documentRels.length) state.zip.file(`word/_rels/${target}.rels`, documentRelsXml(partState.documentRels));
    let rel = state.documentRels.find((item) => item.target === target && item.type.endsWith(`/${kind}`));
    if (!rel) {
      rel = { id: uid(`rId${kind}`), type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, target };
      state.documentRels.push(rel);
    }
    state.relsChanged = true;
    const references = Array.from(section.children).filter((element) => element.tagName === `w:${kind}Reference` && (element.getAttribute('w:type') ?? 'default') === variant);
    references.forEach((element) => element.remove());
    const reference = sectionDoc.createElementNS(W_NS, `w:${kind}Reference`);
    reference.setAttribute('w:type', variant);
    reference.setAttributeNS(R_NS, 'r:id', rel.id);
    insertInSequence(section, reference, SECT_PR_ORDER);
    const ct = parseXmlDom(state.contentTypesXml ?? contentTypesXml(false, false, false));
    const types = ct.documentElement;
    if (!Array.from(types.children).some((element) => element.getAttribute('PartName') === `/${part}`)) {
      const override = ct.createElementNS(types.namespaceURI, 'Override');
      override.setAttribute('PartName', `/${part}`);
      override.setAttribute('ContentType', `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
      types.appendChild(override);
    }
    state.contentTypesXml = updateContentTypes(new XMLSerializer().serializeToString(ct), new Set(partState.documentRels.filter((item) => item.type.endsWith('/image')).map((item) => extensionOf(item.target).replace('.', ''))));
    state.contentTypesChanged = true;
  }
  state.sectPrXml = new XMLSerializer().serializeToString(section);
}

function buildDocumentXml(body: string, page: PageSetup, state: SerializeState): string {
  return `${xmlDeclaration()}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}"${VML_NS_DECL} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" mc:Ignorable="w14 wp14"><w:body>${body}${buildSectPr(page, state)}</w:body></w:document>`;
}

function defaultStylesXml(doc: DocModel): string {
  const font = escapeXml(doc.defaultFont);
  const halfPoints = Math.max(2, Math.round(doc.defaultFontSizePt * 2));
  return `${xmlDeclaration()}<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
}

/**
 * 功能：把一条样式定义序列化为 CT_Style，元素顺序严格按 ECMA-376 排列；时间：2026-10-29；作者：lq。
 *
 * 说明：
 * 1. 内置样式写入 Word 认识的英文名（heading 1 / Title…），Word 会按内置样式处理，
 *    标题层级、导航窗格与目录才能正常工作；
 * 2. 只输出段落与字符属性，不写 numbering 引用——列表标记由正文文本承担，
 *    避免与 Word 内置列表样式叠加出双重编号。
 */
function styleDefinitionXml(definition: DocStyleDefinition, doc: DocModel): string {
  const parts: string[] = [];
  parts.push(`<w:name w:val="${escapeXml(definition.builtinName ?? definition.name)}"/>`);
  const basedOn = definition.basedOn ?? (definition.id === 'Normal' ? undefined : 'Normal');
  if (basedOn && basedOn !== definition.id) parts.push(`<w:basedOn w:val="${escapeXml(basedOn)}"/>`);
  if (definition.id !== 'Normal') parts.push('<w:next w:val="Normal"/>');
  parts.push('<w:qFormat/>');
  parts.push(
    paragraphProperties(
      { type: 'paragraph', id: `style-${definition.id}`, runs: [], style: definition.paragraph },
      undefined,
      false,
    ),
  );
  parts.push(runProperties(resolveStyleRun(definition, doc)));
  const custom = definition.origin === 'custom' ? ' w:customStyle="1"' : '';
  return `<w:style w:type="paragraph"${custom} w:styleId="${escapeXml(definition.id)}">${parts.join('')}</w:style>`;
}

/**
 * 功能：把文档用到的样式合并进 styles.xml，已存在的同 ID 样式保持原样；时间：2026-10-29；作者：lq。
 * 采用「在 </w:styles> 前追加」而不是 DOM 重新序列化，避免丢失命名空间前缀与 mc:Ignorable 等扩展属性。
 */
export function mergeStyleDefinitions(
  xml: string,
  definitions: ReadonlyArray<DocStyleDefinition>,
  doc: DocModel,
): string {
  if (definitions.length === 0) return xml;
  const existing = new Set<string>();
  try {
    const dom = parseXmlDom(xml);
    for (const node of Array.from(dom.getElementsByTagName('w:style'))) {
      const id = getAttrLocal(node, 'styleId');
      if (id) existing.add(id);
    }
  } catch {
    return xml;
  }
  const additions = definitions
    .filter((definition) => !existing.has(definition.id))
    .map((definition) => styleDefinitionXml(definition, doc))
    .join('');
  if (!additions) return xml;
  const index = xml.lastIndexOf('</w:styles>');
  if (index === -1) return xml;
  return `${xml.slice(0, index)}${additions}${xml.slice(index)}`;
}

function corePropertiesXml(doc: DocModel): string {
  const now = new Date().toISOString();
  const title = escapeXml(doc.meta.title ?? doc.meta.author ?? 'Mini Office 文档');
  const author = escapeXml(doc.meta.author ?? 'Mini Office');
  return `${xmlDeclaration()}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${title}</dc:title><dc:creator>${author}</dc:creator><cp:lastModifiedBy>Mini Office</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${doc.meta.created ?? now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${doc.meta.modified ?? now}</dcterms:modified></cp:coreProperties>`;
}

function appPropertiesXml(): string {
  return `${xmlDeclaration()}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Mini Office</Application><AppVersion>0.1</AppVersion></Properties>`;
}

function contentTypesXml(hasJpeg: boolean, hasPng: boolean, hasGif: boolean): string {
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    hasPng ? '<Default Extension="png" ContentType="image/png"/>' : '',
    hasJpeg ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : '',
    hasGif ? '<Default Extension="gif" ContentType="image/gif"/>' : '',
  ].join('');
  return `${xmlDeclaration()}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
}

function packageRelsXml(): string {
  return `${xmlDeclaration()}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

function documentRelsXml(rels: RelationshipInfo[]): string {
  const items = rels
    .map(
      (rel) =>
        `<Relationship Id="${escapeXml(rel.id)}" Type="${escapeXml(rel.type)}" Target="${escapeXml(rel.target)}"${rel.targetMode === 'External' ? ' TargetMode="External"' : ''}/>`,
    )
    .join('');
  return `${xmlDeclaration()}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}

function settingsXml(): string {
  return `${xmlDeclaration()}<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="420"/><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
}

function updateContentTypes(xml: string, extensions: Set<string>): string {
  const typeMap: Record<string, string> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    jpg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
  };
  let next = xml;
  const appended: string[] = [];
  for (const extension of extensions) {
    const lower = extension.toLowerCase();
    if (new RegExp(`<Default[^>]*Extension="${lower}"`, 'i').test(next)) continue;
    const type = typeMap[lower];
    if (!type) continue;
    appended.push(`<Default Extension="${lower}" ContentType="${type}"/>`);
  }
  if (appended.length === 0) return xml;
  if (next.includes('</Types>')) next = next.replace('</Types>', `${appended.join('')}</Types>`);
  return next;
}

function collectSectPrXml(documentXml: string | undefined): string | undefined {
  if (!documentXml) return undefined;
  try {
    const doc = parseXmlDom(documentXml);
    const sections = doc.getElementsByTagName('w:sectPr');
    const section = sections[sections.length - 1];
    return section ? new XMLSerializer().serializeToString(section) : undefined;
  } catch {
    return undefined;
  }
}

function decimalToHexByte(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0');
}

function ensurePngPlaceholder(): string {
  // 1×1 transparent PNG for pathological empty image references.
  return `data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==`;
}

export interface SerializeDocxOptions {
  sourceBytes?: Uint8Array;
  name: string;
  /** 功能：批量报告按节合并，生成结果移除数据源连接；时间：2026-10-24；作者：lq。 */
  sections?: DocModel[];
  materialized?: boolean;
  /** 邮件合并面板持久化状态（JSON），普通保存时写入 settings.xml 的文档变量，成品输出时移除。 */
  mailMergeState?: string;
  /** 模板保存时写给 Word / WPS 的邮件合并数据源与筛选，成品输出时移除。 */
  mailMerge?: WordMailMergeInfo;
}

export async function serializeDocx(
  doc: DocModel,
  options: SerializeDocxOptions,
): Promise<Uint8Array> {
  const zip = options.sourceBytes
    ? await JSZip.loadAsync(options.sourceBytes)
    : new JSZip();
  const sourceExists = !!options.sourceBytes;
  const documentXml = sourceExists ? await zip.file('word/document.xml')?.async('text') : undefined;
  const relsXml = sourceExists ? await zip.file('word/_rels/document.xml.rels')?.async('text') : undefined;
  const ctXml = sourceExists ? await zip.file('[Content_Types].xml')?.async('text') : undefined;

  const state: SerializeState = {
    zip,
    sourceExists,
    documentRels: existingRels(relsXml),
    relsChanged: !sourceExists,
    contentTypesXml: ctXml,
    contentTypesChanged: !sourceExists,
    addedImages: new Map(),
    imageCounter: 0,
    drawingCounter: 0,
    bookmarkCounter: 0,
    hasNumbering: sourceExists && zip.file('word/numbering.xml') !== null,
    sectPrXml: collectSectPrXml(documentXml),
  };

  // 功能：每条记录对应独立节，分别设置页眉页脚及起始页码；时间：2026-10-24；作者：lq。
  const sections = options.sections ?? documentSections(doc);
  const originalSection = state.sectPrXml;
  let body = '';
  for (let index = 0; index < sections.length; index++) {
    const current = sections[index];
    state.sectPrXml = originalSection;
    const sectionDoc = sections.length > 1 ? { ...current, header: current.header ? { ...current.header, edited: true } : undefined, footer: current.footer ? { ...current.footer, edited: true } : undefined } : current;
    await writeHeaderFooters(sectionDoc, state, options.sections || sections.length > 1 ? `-${index + 1}` : '');
    let content = '';
    const blocks = [...current.blocks];
    while (blocks.at(-1)?.type === 'pageBreak') blocks.pop();
    for (const block of blocks) content += await blockXml(block, state);
    if (index < sections.length - 1) {
      const sectionDom = parseXmlDom(`<w:root xmlns:w="${W_NS}" xmlns:r="${R_NS}">${buildSectPr(current.page, state)}</w:root>`);
      const section = firstDescendant(sectionDom, 'w:sectPr')!;
      const type = firstChild(section, 'w:type') ?? sectionDom.createElementNS(W_NS, 'w:type');
      // 功能：保留原节的分隔类型，continuous 节写回后仍在同一页接续；时间：2026-10-30。
      type.setAttribute('w:val', current.page.type ?? 'nextPage');
      if (!type.parentNode) insertInSequence(section, type, SECT_PR_ORDER);
      const sectionXml = new XMLSerializer().serializeToString(section);
      // 在正文段落内附加节属性；末块为表格时使用零高度节标记，避免额外空白页。
      if (blocks.at(-1)?.type === 'paragraph') {
        const lastStart = content.lastIndexOf('<w:p>');
        const last = content.slice(lastStart);
        const withSection = last.includes('</w:pPr>') ? last.replace('</w:pPr>', `${sectionXml}</w:pPr>`) : last.replace('<w:p>', `<w:p><w:pPr>${sectionXml}</w:pPr>`);
        content = content.slice(0, lastStart) + withSection;
      } else content += `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>${sectionXml}</w:pPr></w:p>`;
    }
    body += content;
  }
  const documentContent = buildDocumentXml(body, sections.at(-1)?.page ?? doc.page, state);
  zip.file('word/document.xml', documentContent);

  // 功能：批注部件（word/comments.xml）——含内容类型与关系；时间：2026-10-31。
  if (doc.comments?.length) {
    let commentsXml = '';
    for (const comment of doc.comments) {
      let commentBody = '';
      const blocks = comment.blocks?.length
        ? comment.blocks
        : [{ type: 'paragraph' as const, id: uid('p'), style: {}, runs: [{ text: comment.text, style: {} }] }];
      for (const block of blocks) commentBody += block.type === 'paragraph' ? await paragraphXml(block, state) : '';
      commentsXml += `<w:comment w:id="${escapeXml(comment.id)}" w:author="${escapeXml(comment.author || '用户')}"${comment.date ? ` w:date="${escapeXml(comment.date)}"` : ''}${comment.resolved ? ' w:resolved="1"' : ''}>${commentBody || '<w:p/>'}</w:comment>`;
    }
    const commentsPart = 'word/comments.xml';
    state.zip.file(commentsPart, `${xmlDeclaration()}<w:comments xmlns:w="${W_NS}">${commentsXml}</w:comments>`);
    if (!state.documentRels.some((item) => item.target === 'comments.xml')) {
      state.documentRels.push({ id: uid('rIdComments'), type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', target: 'comments.xml' });
      state.relsChanged = true;
    }
    const ct = parseXmlDom(state.contentTypesXml ?? contentTypesXml(false, false, false));
    const types = ct.documentElement;
    if (!Array.from(types.children).some((element) => element.getAttribute('PartName') === '/word/comments.xml')) {
      const override = ct.createElementNS(types.namespaceURI, 'Override');
      override.setAttribute('PartName', '/word/comments.xml');
      override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
      types.appendChild(override);
    }
    state.contentTypesXml = updateContentTypes(new XMLSerializer().serializeToString(ct), new Set());
    state.contentTypesChanged = true;
  }

  if (!sourceExists) {
    zip.file('[Content_Types].xml', contentTypesXml(false, false, false));
    zip.file('_rels/.rels', packageRelsXml());
    zip.file('word/settings.xml', settingsXml());
    zip.file('docProps/core.xml', corePropertiesXml(doc));
    zip.file('docProps/app.xml', appPropertiesXml());
    if (!state.documentRels.some((rel) => rel.type.endsWith('/styles'))) {
      state.documentRels.push({
        id: 'rIdStyles',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
        target: 'styles.xml',
      });
    }
    if (!state.documentRels.some((rel) => rel.type.endsWith('/settings'))) {
      state.documentRels.push({
        id: 'rIdSettings',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
        target: 'settings.xml',
      });
    }
    state.relsChanged = true;
  }

  // 功能：把文档实际用到的样式写入 styles.xml（新文档先建默认表，旧文档在原表上追加）；时间：2026-10-29；作者：lq。
  const currentStylesXml = sourceExists ? await zip.file('word/styles.xml')?.async('text') : undefined;
  // 源包没有 styles.xml 时保持原状：凭空补一个未在 [Content_Types].xml 与 rels 中声明的部件风险更高。
  if (!sourceExists || currentStylesXml) {
    zip.file(
      'word/styles.xml',
      mergeStyleDefinitions(currentStylesXml ?? defaultStylesXml(doc), collectUsedStyleDefinitions(doc), doc),
    );
  }

  // 功能：成品不再连接外部合并数据源，打开时刷新分页域；时间：2026-10-24；作者：lq。
  if (options.materialized || options.sections) {
    const settings = parseXmlDom(await zip.file('word/settings.xml')?.async('text') ?? settingsXml());
    Array.from(settings.getElementsByTagName('w:mailMerge')).forEach((element) => element.remove());
    const update = firstChild(settings.documentElement, 'w:updateFields') ?? settings.createElementNS(W_NS, 'w:updateFields');
    update.setAttribute('w:val', 'true');
    if (!update.parentNode) settings.documentElement.appendChild(update);
    zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
  }

  // 功能：模板保存时写入面板状态与 Word / WPS 兼容的邮件合并信息；成品输出已在前一个分支移除；时间：2026-10-29；作者：lq。
  else if (options.mailMergeState !== undefined || options.mailMerge) {
    const settings = parseXmlDom(await zip.file('word/settings.xml')?.async('text') ?? settingsXml());
    if (options.mailMergeState !== undefined) {
      // 普通保存模板：把邮件合并面板状态写入文档变量，下次打开自动恢复
      const docVars = firstChild(settings.documentElement, 'w:docVars') ?? settings.createElementNS(W_NS, 'w:docVars');
      if (!docVars.parentNode) settings.documentElement.appendChild(docVars);
      let target: Element | undefined;
      for (const element of Array.from(docVars.getElementsByTagName('w:docVar'))) {
        if (element.getAttribute('w:name') === 'mailMergePanelState') target = element;
      }
      if (options.mailMergeState) {
        if (!target) {
          target = settings.createElementNS(W_NS, 'w:docVar');
          target.setAttribute('w:name', 'mailMergePanelState');
          docVars.appendChild(target);
        }
        target.setAttribute('w:val', options.mailMergeState);
      } else if (target) {
        target.remove();
      }
    }
    if (options.mailMerge) {
      // 功能：写出标准 w:mailMerge 与外部数据源关系，让 Word / WPS 能把本模板识别为邮件合并主文档；时间：2026-10-29；作者：lq。
      Array.from(settings.getElementsByTagName('w:mailMerge')).forEach((element) => element.remove());
      // 清除 Word 缓存的收件人名单（recipientData），否则 Word 会沿用旧名单而不执行我们的筛选
      Array.from(settings.getElementsByTagName('w:recipientData')).forEach((element) => element.remove());
      const wrapper = parseXmlDom(`<w:wrapper xmlns:w="${W_NS}" xmlns:r="${R_NS}">${buildWordMailMergeXml(options.mailMerge)}</w:wrapper>`);
      const node = wrapper.documentElement.firstElementChild;
      // w:mailMerge 必须是 w:settings 的靠前子元素（CT_Settings 中排在 defaultTabStop 之前）
      if (node) settings.documentElement.insertBefore(node, settings.documentElement.firstChild);
      // r:id 需要关系命名空间前缀，缺失时补在 w:settings 根上
      if (!settings.documentElement.hasAttributeNS(XMLNS_NS, 'r')) settings.documentElement.setAttributeNS(XMLNS_NS, 'xmlns:r', R_NS);
      const target = fileUriTarget(options.mailMerge.dataSourcePath);
      const rels = existingRels(await zip.file('word/_rels/settings.xml.rels')?.async('text')).filter((rel) => rel.type !== MAIL_MERGE_SOURCE_REL && !rel.type.endsWith('/recipientData'));
      rels.push({ id: MAIL_MERGE_DATA_REL, type: MAIL_MERGE_SOURCE_REL, target, targetMode: 'External' });
      rels.push({ id: MAIL_MERGE_ODSO_REL, type: MAIL_MERGE_SOURCE_REL, target, targetMode: 'External' });
      zip.file('word/_rels/settings.xml.rels', documentRelsXml(rels));
      zip.remove('word/recipientData.xml');
    }
    zip.file('word/settings.xml', new XMLSerializer().serializeToString(settings));
  }

  if (state.relsChanged || !zip.file('word/_rels/document.xml.rels')) {
    zip.file('word/_rels/document.xml.rels', documentRelsXml(state.documentRels));
  }
  if (state.contentTypesChanged) {
    const extensions = new Set<string>();
    for (const rel of state.documentRels) {
      if (!rel.type.endsWith('/image')) continue;
      const extension = extensionOf(rel.target).replace('.', '').toLowerCase();
      if (extension) extensions.add(extension);
    }
    const base = state.contentTypesXml ?? contentTypesXml(false, false, false);
    zip.file('[Content_Types].xml', updateContentTypes(base, extensions));
  }

  if (!sourceExists) {
    zip.file(
      '[Content_Types].xml',
      updateContentTypes(
        await zip.file('[Content_Types].xml')!.async('text'),
        new Set(['png', 'jpeg', 'gif']),
      ),
    );
  }

  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

export { ensurePngPlaceholder, decimalToHexByte };
