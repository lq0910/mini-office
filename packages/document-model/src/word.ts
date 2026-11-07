export type WordAlignment = 'left' | 'center' | 'right' | 'justify';

export interface RunStyle {
  fontFamily?: string;
  /** 功能：源字体表明确声明的备用字体，仅用于显示和度量，不替换原字体名；时间：2026-11-05；作者：lq。 */
  fontFallback?: string;
  /** 功能：绑定备用名对应的原字体，工具栏改用其他字体后不继承旧替代规则；时间：2026-11-05；作者：lq。 */
  fontFallbackFor?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  verticalAlign?: 'super' | 'sub' | 'baseline';
  letterSpacingPx?: number;
  allCaps?: boolean;
  smallCaps?: boolean;
}

/**
 * 功能：记录 wp:anchor 的锚定与环绕参数，让浮动图片按 Word 语义定位而不是简单 float；
 * 时间：2026-10-31；作者：lq。
 *
 * Word 的浮动图片是「锚点段落 + 相对基准偏移 + 环绕距离」三部分决定位置的，
 * 只保存 wrap 方向会让图片挤占后续段落的行盒，造成正文莫名换行。
 */
export interface ImageAnchor {
  /** 水平基准：column（栏）/ margin（页边距）/ page（页面）/ character（字符）。 */
  hRelative?: string;
  /** 水平偏移，px；负值表示向左越界。 */
  hOffsetPx?: number;
  /** 水平对齐（与偏移二选一）。 */
  hAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  /** 垂直基准：paragraph（段落）/ line（行）/ margin / page。 */
  vRelative?: string;
  /** 垂直偏移，px。 */
  vOffsetPx?: number;
  vAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
  /** 环绕类型，none 表示浮于文字上/下方不占位。 */
  wrapType?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** 环绕侧：both / left / right / largest。 */
  wrapText?: string;
  /** 衬于文字下方。 */
  behindDoc?: boolean;
  /** 四周环绕距离（EMU 转 px）。 */
  distTopPx?: number;
  distBottomPx?: number;
  distLeftPx?: number;
  distRightPx?: number;
}

export interface DocImage {
  id?: string;
  src: string;
  widthPx: number;
  heightPx: number;
  alt?: string;
  /** Existing OOXML relationship + media part, preserved on save. */
  relId?: string;
  part?: string;
  name?: string;
  /** Base64 data URL for newly inserted images. */
  dataUrl?: string;
  /** 功能：保存图片的文字环绕方式，支持嵌入、左右环绕和上下型；时间：2026-10-24；作者：lq。 */
  wrap?: 'inline' | 'square-left' | 'square-right' | 'top-bottom';
  /** 功能：浮动图片的锚定参数，保存时原样写回；时间：2026-10-31；作者：lq。 */
  anchor?: ImageAnchor;
  /** 功能：保留图片效果外扩的占位尺寸（px），用于内嵌行排版及 OOXML 保存；时间：2026-11-04；作者：lq。 */
  effectExtent?: { left: number; top: number; right: number; bottom: number };
}

export interface MergeFieldFormat {
  num?: string;        // \# 格式，如 "#,##0.00"
  date?: string;       // \@ 格式，如 "yyyy年M月d日"
  case?: 'upper' | 'lower' | 'caps' | 'firstcap'; // \* 开关
  before?: string;     // \b 非空前缀
  after?: string;      // \f 非空后缀
}

export interface DocRun {
  text?: string;
  style: RunStyle;
  image?: DocImage;
  /** Explicit line break attached to this run. */
  lineBreak?: boolean;
  tab?: boolean;
  /** Field code such as PAGE / NUMPAGES, resolved by the renderer. */
  field?: string;
  /** 功能：结构化邮件合并字段名称，与普通文字独立保存；时间：2026-10-24；作者：lq。 */
  mergeField?: string;
  mergeFormat?: MergeFieldFormat;
  /** Word 邮件合并控制域（解析后占位文本，合并时消耗） */
  mergeNext?: boolean;         // NEXT
  mergeNextIf?: string;        // NEXTIF 条件表达式
  mergeSkipIf?: string;        // SKIPIF 条件表达式
  mergeRecord?: boolean;       // MERGEREC
  mergeSeq?: boolean;          // MERGESEQ
}

/**
 * 功能：Word 制表位（w:tabs/w:tab），标尺与正文 Tab 定位共用同一份数据；时间：2026-10-30；作者：lq。
 *
 * `posPx` 相对正文栏左沿，不叠加段落左缩进，与 OOXML 的 `w:pos` 语义一致，
 * 因此解析后可直接与浏览器实测位置比较，不需要再做一次坐标换算。
 * 功能：依据公开制表位探针明确缩进后的坐标基准；时间：2026-11-03；作者：lq。
 */
export type TabAlignment = 'left' | 'center' | 'right' | 'decimal';

export type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'heavy';

export interface TabStop {
  posPx: number;
  alignment: TabAlignment;
  leader?: TabLeader;
}

export interface ParagraphStyle {
  alignment?: WordAlignment;
  indentLeftPx?: number;
  indentRightPx?: number;
  firstLineIndentPx?: number;
  spaceBeforePx?: number;
  spaceAfterPx?: number;
  /** 功能：同段落样式相邻时合并段前段后距（w:contextualSpacing）；时间：2026-11-05；作者：lq。 */
  contextualSpacing?: boolean;
  /** Multiplier, e.g. 1.5. */
  lineHeight?: number;
  /** Exact line height in CSS px (fixed line spacing). */
  lineHeightExactPx?: number;
  /** 功能：最小值行距（w:lineRule="atLeast"），低于该高度时按内容撑开；时间：2026-10-29；作者：lq。 */
  lineHeightAtLeastPx?: number;
  keepNext?: boolean;
  /** 功能：整段不拆页（w:keepLines），用于条款、签名栏等不允许跨页的段落；时间：2026-11-03；作者：lq。 */
  keepLines?: boolean;
  /** 功能：孤行控制（w:widowControl），默认开启；显式 false 表示允许页首页尾出现单行；时间：2026-11-03；作者：lq。 */
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  /**
   * 功能：段落网格吸附开关（w:snapToGrid）；缺省 true 表示吸附到文档网格行距；
   * 显式 false 时段落按字体度量定行高，不使用 docGrid 的 linePitch；时间：2026-11-03；作者：lq。
   */
  snapToGrid?: boolean;
  outlineLevel?: number;
  backgroundColor?: string;
  /** 功能：保留段落上下左右边框，用于兼容 Word 横线与签名栏；时间：2026-10-27；作者：lq。 */
  borders?: TableBorders;
  /** 功能：段落自定义制表位，为空时按文档默认制表位（defaultTabStopPx）等距推进；时间：2026-10-30；作者：lq。 */
  tabStops?: TabStop[];
}

export interface ListInfo {
  numId?: string;
  level: number;
  marker: string;
  ordered: boolean;
}

/**
 * 功能：文档级样式定义，供「样式库」批量套用并在保存时写入 styles.xml；时间：2026-10-29；作者：lq。
 *
 * 设计要点：
 * 1. `id` 直接作为 OOXML 的 `w:styleId`，内置预设沿用 Word 的内置 ID（Heading1/Title…），
 *    这样打开 Word 文档时其标题段落能被识别，反向保存后 Word 也仍视其为内置样式。
 * 2. 只保存「样式自身」的格式，不保存继承结果；应用时再与文档默认字体合并，
 *    避免文档换默认字体后样式仍停留在旧字体上。
 * 3. `origin` 仅用于 UI 分组与删除权限，不写入文件。
 */
export interface DocStyleDefinition {
  id: string;
  name: string;
  /** OOXML 内置样式名（如 heading 1），写入 styles.xml 后 Word 会按内置样式处理。 */
  builtinName?: string;
  paragraph: ParagraphStyle;
  run: RunStyle;
  basedOn?: string;
  /** 列表型样式：应用时同时写入段落列表标记。 */
  list?: { ordered: boolean; level: number };
  origin?: 'preset' | 'custom' | 'imported';
}

export interface DocParagraph {
  type: 'paragraph';
  id: string;
  runs: DocRun[];
  style: ParagraphStyle;
  /** 功能：独立保留段落结束标记的字符格式，不将其加粗、字号等传播到正文；时间：2026-11-05；作者：lq。 */
  markStyle?: RunStyle;
  list?: ListInfo;
  styleId?: string;
  /** Source paragraph index for stable round-trips. */
  sourceIndex?: number;
  /** 功能：保留段尾分节信息，让批量报告重新打开后仍独立编号；时间：2026-10-24；作者：lq。 */
  sectionEnd?: DocSection;
  /**
   * 功能：标记该段落是目录域（TOC），渲染时按文档标题实时生成条目；时间：2026-10-30；作者：lq。
   *
   * 目录用「一个段落承载全部条目」而不是新增块类型，原因是分页、测量、撤销、查找替换等
   * 现有流程都以段落为最小单位，这样改造面最小且不会漏掉某个分支。
   */
  toc?: { levels: number };
}

export interface TableBorderSide {
  style: 'none' | 'single' | 'double' | 'dashed' | 'dotted';
  widthPx: number;
  color: string;
}

export interface TableBorders {
  top?: TableBorderSide;
  right?: TableBorderSide;
  bottom?: TableBorderSide;
  left?: TableBorderSide;
  insideH?: TableBorderSide;
  insideV?: TableBorderSide;
}

export interface DocTableCell {
  id: string;
  blocks: DocBlock[];
  colSpan: number;
  rowSpan: number;
  widthPx?: number;
  backgroundColor?: string;
  borders?: TableBorders;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** 功能：单元格自定义内边距（w:tcMar），缺省沿用表格级 tblCellMar；时间：2026-11-03；作者：lq。 */
  marginPx?: TableCellMargin;
}

/**
 * 功能：表格与单元格的单元格内边距；时间：2026-11-03；作者：lq。
 *
 * Word 默认上下为 0、左右 108 twips（≈5.4px）。tblCellMar 定义表格级默认值，
 * tcMar 定义单个单元格的覆盖值。
 */
export interface TableCellMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface DocTableRow {
  id: string;
  cells: DocTableCell[];
  heightPx?: number;
  /**
   * 功能：行高规则（w:trHeight@w:hRule）；缺省为 atLeast，与 Word 一致；时间：2026-11-03；作者：lq。
   *
   * atLeast = 最小高度，内容更高时行会被撑开；exact = 固定高度，内容溢出被裁剪。
   */
  heightRule?: 'atLeast' | 'exact';
  isHeader?: boolean;
}

export interface DocTable {
  type: 'table';
  id: string;
  rows: DocTableRow[];
  columnWidthsPx: number[];
  widthPct?: number;
  alignment?: WordAlignment;
  borders?: TableBorders;
  styleId?: string;
  /**
   * 功能：表格列宽布局模式（w:tblLayout@w:type）；时间：2026-11-03；作者：lq。
   *
   * fixed = 严格使用 tblGrid 指定的列宽；autofit = 按正文可用宽度自适应收敛。
   * OOXML 中缺省为 autofit，但仅在 tblW 为 auto 时才真正生效。
   */
  layout?: 'fixed' | 'autofit';
  /** 功能：表格左缩进（w:tblInd）；时间：2026-11-03；作者：lq。 */
  indentPx?: number;
  /** 功能：表格级单元格内边距（w:tblCellMar）；时间：2026-11-03；作者：lq。 */
  cellMarginPx?: TableCellMargin;
}

export interface PageBreakBlock {
  type: 'pageBreak';
  id: string;
}

export type DocBlock = DocParagraph | DocTable | PageBreakBlock;

export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter?: number;
}

export interface PageSetup {
  widthPx: number;
  heightPx: number;
  orientation: 'portrait' | 'landscape';
  sizeName: string;
  margins: PageMargins;
  columns?: number;
  /** 功能：页码起始值，独立于物理页序；时间：2026-10-24；作者：lq。 */
  pageNumberStart?: number;
  /** 功能：文档网格（w:docGrid）；中文公文按网格逐行对齐，缺省表示未启用；时间：2026-11-02；作者：lq。 */
  docGrid?: DocGrid;
}

/**
 * 功能：文档网格设置；时间：2026-11-02；作者：lq。
 *
 * Word 在启用网格时用 `linePitch` 作为「单倍行距」的行高，多倍行距按倍数叠加，
 * 因此排版不能只看字体度量，否则公文类文档每页行数会对不上。
 */
export interface DocGrid {
  type: 'none' | 'lines' | 'linesAndChars' | 'snapToChars';
  /** 功能：行距只对 lines/linesAndChars 生效；snapToChars 可无此属性；时间：2026-11-05；作者：lq。 */
  linePitchPx?: number;
  /** 功能：字符间距增量保留零值及负值；当前不据此承诺字符网格排版完整支持；时间：2026-11-05；作者：lq。 */
  charSpacePx?: number;
}

/** 功能：页眉页脚中只读显示的 DrawingML 文本框，保留原绘图 XML 供保存回写；时间：2026-11-05；作者：lq。 */
export interface HeaderFooterTextBox {
  id: string;
  blocks: DocBlock[];
  widthPx: number;
  heightPx: number;
  anchor?: ImageAnchor;
  autoFit: boolean;
  /** 功能：自动适应且不折行时按内容宽度收缩，允许折行时仍保留声明的框宽；时间：2026-11-05；作者：lq。 */
  noWrap: boolean;
  insetsPx: { left: number; top: number; right: number; bottom: number };
  sourceXml: string;
}

export interface HeaderFooter {
  blocks: DocBlock[];
  distancePx: number;
  pageNumbers?: boolean;
  /** 功能：仅重写经用户编辑的页眉页脚，保留导入文档的其他变体；时间：2026-10-24；作者：lq。 */
  edited?: boolean;
  /** 功能：保留页眉页脚部件路径以正确复用图片关系；时间：2026-10-24；作者：lq。 */
  sourcePart?: string;
  /** 功能：锚定文本框不挤占普通页眉页脚段落，页码域随所在页面动态显示；时间：2026-11-05；作者：lq。 */
  textBoxes?: HeaderFooterTextBox[];
}

/** 功能：描述一个报告节的页面与页眉页脚；时间：2026-10-24；作者：lq。 */
export interface DocSection { page: PageSetup; header?: HeaderFooter; footer?: HeaderFooter }

/** 功能：邮件合并筛选条件，面板筛选与 Word 兼容输出共用同一结构；时间：2026-10-29；作者：lq。 */
export interface DocMergeCondition {
  field: string;
  op: 'contains' | 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte' | 'empty' | 'notEmpty' | 'notContains';
  value: string;
  /** 与上一条件的组合方式，仅第 2 条起生效；对齐 Word 查询选项的「且 / 或」。 */
  conjunction?: 'and' | 'or';
}

export interface DocMeta {
  title?: string;
  author?: string;
  subject?: string;
  created?: string;
  modified?: string;
  lastModifiedBy?: string;
  /** 邮件合并面板持久化的筛选/排序/输出状态（JSON 字符串），保存在 settings.xml 的文档变量中。 */
  mailMergeState?: string;
  /** 写给 Word / WPS 的邮件合并数据源与筛选，保存模板时转成标准 w:mailMerge，不参与 mini-office 自身恢复。 */
  mailMergeTemplate?: { dataSourcePath: string; sheetName: string; fields: string[]; conditions: DocMergeCondition[] };
}

export interface DocModel {
  version: 1;
  page: PageSetup;
  blocks: DocBlock[];
  header?: HeaderFooter;
  footer?: HeaderFooter;
  defaultFont: string;
  defaultFontSizePt: number;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  /** Map relationship id → media info for images referenced by runs. */
  images: Record<string, { src: string; part?: string }>;
  meta: DocMeta;
  hasTrackedChanges?: boolean;
  hasComments?: boolean;
  hasContentControls?: boolean;
  stats?: { words: number; characters: number };
  /** 功能：页面渲染时的节总页数，不写入文档包；时间：2026-10-24；作者：lq。 */
  sectionPageCount?: number;
  /** 功能：本文档自定义/导入的样式，与内置预设合并后供样式库使用；时间：2026-10-29；作者：lq。 */
  styles?: DocStyleDefinition[];
  /** 功能：文档默认制表位间距（settings.xml 的 w:defaultTabStop），未设置时按 0.5 英寸处理；时间：2026-10-30；作者：lq。 */
  defaultTabStopPx?: number;
  /** 功能：保留导入 XML 与块基线，用于仅回写用户修改的正文；时间：2026-11-03；作者：lq。 */
  source?: WordSourceSnapshot;
  /** 功能：区分导入定位失败与正常新建文档，定位失败时禁止有损生成覆盖；时间：2026-11-07；作者：lq。 */
  sourceMappingFailed?: boolean;
}

/** 功能：记录源正文区间及编辑基线，不参与渲染或导出内容；时间：2026-11-03；作者：lq。 */
export interface WordSourceSnapshot {
  xml: string;
  bodyStart: number;
  bodyEnd: number;
  signature: string;
  page: string;
  styles: string;
  groups: { start: number; end: number; ids: string[]; baseline: string }[];
}

export const A4_PAGE: PageSetup = {
  widthPx: 793.7,
  heightPx: 1122.52,
  orientation: 'portrait',
  sizeName: 'A4',
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
};

export function createBlankDoc(): DocModel {
  return {
    version: 1,
    page: { ...A4_PAGE, margins: { ...A4_PAGE.margins } },
    blocks: [
      {
        type: 'paragraph',
        id: 'p-1',
        runs: [{ text: '', style: { fontSizePt: 10.5 } }],
        style: { alignment: 'left', lineHeight: 1.6 },
      },
    ],
    defaultFont: '宋体',
    defaultFontSizePt: 10.5,
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    images: {},
    meta: { title: '未命名文档' },
    stats: { words: 0, characters: 0 },
  };
}
