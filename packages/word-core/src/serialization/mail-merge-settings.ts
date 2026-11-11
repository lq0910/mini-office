/** 功能：模板面板配置按原生快照保存，未触及配置时保持导入邮件合并关系与原始字节；时间：2026-11-11；作者：lq。 */
import { loadDocxZip } from '../upstream/packages/docx-engine/src/zip-load';
import { applyXmlEdits, wordXmlAttr, wordXmlEscape, wordXmlNode, xmlSpans } from './xml-spans';
export interface WordMergeSettings { panelState: string; source?: { xml: string; target: string } | null; }
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** 功能：配置与标准数据源关系独立于正文；仅更新指定变量和邮件合并设置；时间：2026-11-11；作者：lq。 */
export async function patchWordMergeSettings(bytes: Uint8Array, settings: WordMergeSettings): Promise<Uint8Array> {
  const zip = await loadDocxZip(bytes);
  let xml = await zip.file('word/settings.xml')?.async('string') ?? `<w:settings xmlns:w="${W}" xmlns:r="${R}"></w:settings>`;
  let nodes = xmlSpans(xml); let root = nodes.find(node => !node.parent)!;
  if (root.end === root.openEnd) xml = xml.slice(0, root.start) + xml.slice(root.start, root.end).replace(/\/\s*>$/, `></${root.name}>`) + xml.slice(root.end);
  nodes = xmlSpans(xml); root = nodes.find(node => !node.parent)!;
  const existing = nodes.filter(node => wordXmlNode(node, 'docVar') && wordXmlAttr(node, 'name') === 'mailMergePanelState');
  const value = `<w:docVar w:name="mailMergePanelState" w:val="${wordXmlEscape(settings.panelState)}"/>`;
  if (existing.length) xml = applyXmlEdits(xml, existing.map((node, index) => ({ start: node.start, end: node.end, text: index === 0 ? value : '' })));
  else {
    const vars = nodes.find(node => wordXmlNode(node, 'docVars'));
    if (vars && vars.end > vars.openEnd) xml = xml.slice(0, vars.closeStart) + value + xml.slice(vars.closeStart);
    else if (vars) xml = applyXmlEdits(xml, [{ start: vars.start, end: vars.end, text: `<w:docVars>${value}</w:docVars>` }]);
    else {
      const after = root.children.find(node => ['rsids', 'mathPr', 'uiCompat97To2003', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'].includes(node.local));
      const at = after?.start ?? root.closeStart; xml = xml.slice(0, at) + `<w:docVars>${value}</w:docVars>` + xml.slice(at);
    }
  }
  if (settings.source !== undefined) {
    nodes = xmlSpans(xml); const old = nodes.filter(node => wordXmlNode(node, 'mailMerge') || wordXmlNode(node, 'recipientData'));
    xml = applyXmlEdits(xml, old.filter(node => !old.some(parent => parent !== node && parent.start <= node.start && parent.end >= node.end)).map(node => ({ start: node.start, end: node.end, text: '' })));
    let rels = await zip.file('word/_rels/settings.xml.rels')?.async('string') ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const relationships = xmlSpans(rels); const obsolete = relationships.filter(node => node.local === 'Relationship' && /\/(mailMergeSource|recipientData)$/.test(node.attrs.Type));
    rels = applyXmlEdits(rels, obsolete.map(node => ({ start: node.start, end: node.end, text: '' })));
    if (settings.source) {
      nodes = xmlSpans(xml); root = nodes.find(node => !node.parent)!;
      const before = new Set(['writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries', 'displayBackgroundShape', 'printPostScriptOverText', 'printFractionalCharacterWidth', 'printFormsData', 'embedTrueTypeFonts', 'embedSystemFonts', 'saveSubsetFonts', 'saveFormsData', 'mirrorMargins', 'alignBordersAndEdges', 'bordersDoNotSurroundHeader', 'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors', 'hideGrammaticalErrors', 'activeWritingStyle', 'proofState', 'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter', 'stylePaneSortMethod', 'documentType']);
      const at = root.children.find(node => !before.has(node.local))?.start ?? root.closeStart;
      const used = new Set(relationships.filter(node => !obsolete.includes(node)).map(node => node.attrs.Id));
      let i = 0; const allocate = () => { let id: string; do { id = `rIdMoMergeSource${++i}`; } while (used.has(id)); used.add(id); return id; };
      const dataId = allocate(), odsoId = allocate();
      const merge = settings.source.xml.replace(/r:id="rIdMailMergeSource"/g, `r:id="${dataId}"`).replace(/r:id="rIdMailMergeOdso"/g, `r:id="${odsoId}"`);
      xml = xml.slice(0, at) + merge + xml.slice(at);
      const relRoot = xmlSpans(rels).find(node => !node.parent)!;
      if (relRoot.openEnd === relRoot.end) rels = rels.replace(/\/\s*>\s*$/, `></${relRoot.name}>`);
      const end = xmlSpans(rels).find(node => !node.parent)!.closeStart;
      rels = rels.slice(0, end) + [dataId, odsoId].map(id => `<Relationship Id="${id}" Type="${R}/mailMergeSource" Target="${wordXmlEscape(settings.source!.target)}" TargetMode="External"/>`).join('') + rels.slice(end);
    }
    zip.file('word/_rels/settings.xml.rels', rels);
  }
  root = xmlSpans(xml).find(node => !node.parent)!;
  if (!root.attrs['xmlns:r']) xml = xml.slice(0, root.openEnd - 1) + ` xmlns:r="${R}"` + xml.slice(root.openEnd - 1);
  if (!root.attrs['xmlns:w']) { root = xmlSpans(xml).find(node => !node.parent)!; xml = xml.slice(0, root.openEnd - 1) + ` xmlns:w="${W}"` + xml.slice(root.openEnd - 1); }
  zip.file('word/settings.xml', xml);
  // 功能：原模板可能没有 settings 部件，新增时补齐包类型及主关系；时间：2026-11-11；作者：lq。
  let rels = await zip.file('word/_rels/document.xml.rels')?.async('string') ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  if (!xmlSpans(rels).some(node => node.local === 'Relationship' && /\/settings$/.test(node.attrs.Type))) {
    const ids = new Set(xmlSpans(rels).map(node => node.attrs.Id)); let id = 'rIdMoMergeSettings', n = 0; while (ids.has(id)) id = `rIdMoMergeSettings${++n}`;
    const root = xmlSpans(rels).find(node => !node.parent)!;
    if (root.openEnd === root.end) rels = rels.replace(/\/\s*>\s*$/, `></${root.name}>`);
    const at = xmlSpans(rels).find(node => !node.parent)!.closeStart; rels = rels.slice(0, at) + `<Relationship Id="${id}" Type="${R}/settings" Target="settings.xml"/>` + rels.slice(at); zip.file('word/_rels/document.xml.rels', rels);
  }
  let types = await zip.file('[Content_Types].xml')!.async('string');
  if (!xmlSpans(types).some(node => node.local === 'Override' && node.attrs.PartName === '/word/settings.xml')) {
    const at = xmlSpans(types).find(node => !node.parent)!.closeStart; types = types.slice(0, at) + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' + types.slice(at); zip.file('[Content_Types].xml', types);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
