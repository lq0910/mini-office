export type WordAlignment = 'left' | 'center' | 'right' | 'justify';

/** 功能：编号层级定义（numbering.xml 的 w:lvl），编辑期重编号使用；时间：2026-10-31。 */
export interface NumberingLevel {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
  indentLeftPx?: number;
  firstLineIndentPx?: number;
}

export interface RunStyle {
  fontFamily?: string;
  /** 功能：中文字体（Word w:rFonts 的 w:eastAsia 槽），与西文 fontFamily 分开记录，避免中文样式被西文覆盖；时间：2026-10-30。 */
  eastAsiaFont?: string;
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
  /** 功能：保存图片的文字环绕方式，支持嵌入、左右环绕、上下型和无环绕（VML rect 浮层）；时间：2026-10-24；作者：lq。 */
  wrap?: 'inline' | 'square-left' | 'square-right' | 'top-bottom' | 'none';
  /**
   * 功能：VML 浮层定位（v:rect 带 position:absolute 时存在），CSS px；
   * left 按 xRelative 解释：page=页面左缘，text=内容区左缘；top 为相对锚定段落顶缘的偏移（可负）；时间：2026-10-30。
   */
  vmlRect?: { left: number; top: number; width: number; height: number; xRelative?: 'page' | 'text'; checked?: boolean };
  /** 功能：VML rect 的描边颜色与粗细（px），用于渲染方框边框而非占位字符；时间：2026-10-30。 */
  strokeColor?: string;
  strokeWidthPx?: number;
}

export interface MergeFieldFormat {
  num?: string;        // \# 格式，如 "#,##0.00"
  date?: string;       // \@ 格式，如 "yyyy年M月d日"
  case?: 'upper' | 'lower' | 'caps' | 'firstcap'; // \* 开关
  before?: string;     // \b 非空前缀
  after?: string;      // \f 非空后缀
}

/**
 * 功能：Word 批注（w:comment）——面板展示与正文范围高亮共用；时间：2026-10-31。
 */
export interface DocComment {
  /** Word w:comment 的 w:id（数字字符串）。 */
  id: string;
  author: string;
  date?: string;
  resolved?: boolean;
  /** 批注正文（纯文本聚合，面板显示用）。 */
  text: string;
  /** 批注内容段落（保存写回 comments.xml；无则由 text 生成）。 */
  blocks?: DocBlock[];
}

/**
 * 功能：修订标记（w:ins/w:del）——run 级保留插入与删除痕迹；时间：2026-10-31。
 */
export interface RunRevision {
  type: 'insert' | 'delete';
  author?: string;
  date?: string;
}

/**
 * 功能：Word 内容控件（w:sdt）的语义——复选框可勾选、下拉/组合框可选项、日期选择器可按格式填值；
 * 解析后保留控件类型与当前值，渲染成可交互控件并可原样写回；时间：2026-10-30。
 */
export interface RunControl {
  kind: 'checkbox' | 'dropdown' | 'comboBox' | 'date' | 'text';
  /** 复选框当前是否勾选（对应 w14:checked / w:checked）。 */
  checked?: boolean;
  /** 下拉/组合框的候选值（w:listItem 的 displayText）。 */
  choices?: string[];
  /** 日期格式（w:dateFormat，如 yyyy年M月d日）。 */
  dateFormat?: string;
  /** 控件别名与标记，写回时保留（w:alias / w:tag）。 */
  alias?: string;
  tag?: string;
  /** 尚未填写（内容为空、显示占位提示）。 */
  placeholder?: boolean;
  /** 是否锁定内容不可编辑（w:lock）。 */
  locked?: boolean;
}

export interface DocRun {
  text?: string;
  style: RunStyle;
  image?: DocImage;
  /** Explicit line break attached to this run. */
  lineBreak?: boolean;
  /** 功能：段内分页符（w:br w:type="page"），该位置必须强制断页；时间：2026-10-30；作者：lq。 */
  pageBreak?: boolean;
  /** 功能：段内分栏符（w:br w:type="column"），多栏节中在此换到下一栏；时间：2026-10-30。 */
  columnBreak?: boolean;
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
  /** 功能：内部书签链接（指向 DocParagraph.bookmarks 中的书签名），用于目录项点击跳转；时间：2026-10-30。 */
  anchor?: string;
  /** 功能：外部超链接地址（由 w:hyperlink 的 r:id 经文档关系解析得到）；时间：2026-10-30。 */
  href?: string;
  /** 功能：该 run 是内容控件（w:sdt）的呈现内容；时间：2026-10-30。 */
  control?: RunControl;
  /** 功能：run 所在的批注范围（w:commentRangeStart/End 之间）；时间：2026-10-31。 */
  commentIds?: string[];
  /** 功能：修订标记（w:ins/w:del）；时间：2026-10-31。 */
  revision?: RunRevision;
  /** 功能：文本框（只读渲染）——锚定段落定位的绝对定位文字容器；时间：2026-10-31。 */
  textbox?: {
    left: number;
    top: number;
    width: number;
    height: number;
    /** 水平参照系：page=页面左缘，text=内容区左缘。 */
    xRelative?: 'page' | 'text';
    blocks: DocBlock[];
  };
}

export interface WordTabStop {
  /** 距内容盒左缘（左边距）的偏移，CSS px。 */
  posPx: number;
  /** 制表位对齐方式。 */
  kind: 'left' | 'center' | 'right' | 'decimal';
  /** 功能：制表符前导符（Word 目录靠它在标题与页码间填充点/横线）；时间：2026-10-30。 */
  leader?: 'none' | 'dot' | 'hyphen' | 'middleDot' | 'underline';
}

export interface ParagraphStyle {
  alignment?: WordAlignment;
  indentLeftPx?: number;
  indentRightPx?: number;
  firstLineIndentPx?: number;
  /** 制表位（Word 标尺可设）。 */
  tabStops?: WordTabStop[];
  spaceBeforePx?: number;
  spaceAfterPx?: number;
  /** Multiplier, e.g. 1.5. */
  lineHeight?: number;
  /** Exact line height in CSS px (fixed line spacing). */
  lineHeightExactPx?: number;
  /** 功能：最小值行距（w:lineRule="atLeast"），低于该高度时按内容撑开；时间：2026-10-29；作者：lq。 */
  lineHeightAtLeastPx?: number;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
  outlineLevel?: number;
  backgroundColor?: string;
  /** 功能：保留段落上下左右边框，用于兼容 Word 横线与签名栏；时间：2026-10-27；作者：lq。 */
  borders?: TableBorders;
  /** 功能：段落是否对齐到文档网格（w:snapToGrid，缺省对齐；仅显式为 false 时存储）；时间：2026-10-30。 */
  snapToGrid?: boolean;
  /** 功能：孤寡控制（w:widowControl，Word 默认开启；仅显式为 false 时存储）——跨页断点两侧至少保留两行；时间：2026-10-31。 */
  widowControl?: boolean;
  /** 功能：段首分栏符（w:br w:type="column"），多栏节中本段另起一栏；时间：2026-10-30。 */
  columnBreakBefore?: boolean;
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
  list?: ListInfo;
  styleId?: string;
  /** Source paragraph index for stable round-trips. */
  sourceIndex?: number;
  /** 功能：保留段尾分节信息，让批量报告重新打开后仍独立编号；时间：2026-10-24；作者：lq。 */
  sectionEnd?: DocSection;
  /** 功能：段落上标记的书签名集合（来自 w:bookmarkStart），供目录项锚点跳转定位；时间：2026-10-30。 */
  bookmarks?: string[];
  /** 功能：段落内复杂域的域字符标记（begin/separate/end），保存时还原 TOC 等域结构；时间：2026-10-30。 */
  fieldMarkers?: FieldMarker[];
}

/** 功能：复杂域的域字符标记，用于 docx 往返时保留 TOC/PAGEREF 等域结构；时间：2026-10-30。 */
export interface FieldMarker {
  /** 功能：域字符类型；时间：2026-10-30。 */
  role: 'begin' | 'separate' | 'end';
  /** 功能：域指令文本（如 TOC \o "1-3" \h \z \u），仅 begin 使用；时间：2026-10-30。 */
  instr?: string;
  /** 功能：标记插入位置——位于本段第 N 个 run 之后；时间：2026-10-30。 */
  afterRun: number;
  /** 功能：该域已由 fldSimple 表达（MERGEFIELD/PAGE 等），无需再写复杂域结构；时间：2026-10-30。 */
  skip?: boolean;
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
}

export interface DocTableRow {
  id: string;
  cells: DocTableCell[];
  heightPx?: number;
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
  /** 每栏宽度（CSS px），w:cols 下每个 w:col 的 w:w；时间：2026-10-30。 */
  columnWidthsPx?: number[];
  /** 每栏间距（CSS px），w:cols 下每个 w:col 的 w:space；时间：2026-10-30。 */
  columnSpacePx?: number[];
  /** 是否等宽栏，w:cols w:equalWidth；时间：2026-10-30。 */
  columnEqualWidth?: boolean;
  /** 节分隔类型：'nextPage' 强制新页，'continuous' 同页继续；时间：2026-10-30。 */
  type?: 'nextPage' | 'continuous';
  /** 文档网格行距（w:docGrid w:linePitch），CSS px，段落行高按此网格对齐；时间：2026-10-30。 */
  linePitchPx?: number;
  /** 功能：首页页眉页脚不同（w:titlePg）；时间：2026-10-31。 */
  titlePage?: boolean;
  /** 功能：页码起始值，独立于物理页序；时间：2026-10-24；作者：lq。 */
  pageNumberStart?: number;
}

export interface HeaderFooter {
  blocks: DocBlock[];
  distancePx: number;
  pageNumbers?: boolean;
  /** 功能：仅重写经用户编辑的页眉页脚，保留导入文档的其他变体；时间：2026-10-24；作者：lq。 */
  edited?: boolean;
  /** 功能：保留页眉页脚部件路径以正确复用图片关系；时间：2026-10-24；作者：lq。 */
  sourcePart?: string;
}

/** 功能：描述一个报告节的页面与页眉页脚（含首页/奇偶页变体）；时间：2026-10-31。 */
export interface DocSection {
  page: PageSetup;
  header?: HeaderFooter;
  footer?: HeaderFooter;
  headerFirst?: HeaderFooter;
  headerEven?: HeaderFooter;
  footerFirst?: HeaderFooter;
  footerEven?: HeaderFooter;
}

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
  /** 功能：编号定义（numbering.xml 的 numId → 层级），供编辑期列表重编号使用；时间：2026-10-31。 */
  numberingLevels?: Record<string, NumberingLevel[]>;
  /** 功能：文档批注（word/comments.xml），正文 run 以 commentIds 关联；时间：2026-10-31。 */
  comments?: DocComment[];
  /** 功能：默认制表位间距（settings.xml 的 w:defaultTabStop，CSS px）；无显式制表位时 tab 按它的整数倍跳转；时间：2026-10-30。 */
  defaultTabStopPx?: number;
  /** 功能：奇偶页页眉页脚不同（settings.xml 的 w:evenAndOddHeaders）；时间：2026-10-31。 */
  evenAndOddHeaders?: boolean;
  /** 功能：首页/奇偶页页眉页脚变体（由节展开合并进文档对象供渲染读取）；时间：2026-10-31。 */
  headerFirst?: HeaderFooter;
  headerEven?: HeaderFooter;
  footerFirst?: HeaderFooter;
  footerEven?: HeaderFooter;
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
