import { DEFAULT_ACCENT_HEX, type AccentId, type AppSettings } from '@mini-office/shared';

// 功能：文档纸张色换算（深色模式提亮文字、压暗底纹、框线跟随主题）供 Word 与 Excel 共用；时间：2026-10-31；作者：lq。
export * from './colors';
export * from './use-paper-theme';

export type ThemeMode = 'light' | 'dark' | 'system';
export type ResolvedTheme = 'light' | 'dark';

export const THEME_STORAGE_KEY = 'mini-office.theme';

export const UI_FONT_STACK =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", system-ui, sans-serif';

export const DOC_FONT_STACK =
  '"PingFang SC", "Songti SC", "SimSun", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", serif';

/** Common Chinese-friendly font choices offered by editors. */
export const FONT_CHOICES = [
  '默认',
  '宋体',
  '仿宋',
  '黑体',
  '楷体',
  '微软雅黑',
  '等线',
  '思源宋体',
  '思源黑体',
  'PingFang SC',
  'Songti SC',
  'Helvetica Neue',
  'Arial',
  'Times New Roman',
  'Courier New',
] as const;

/**
 * 功能：Word 常见西文字体 → 随包度量兼容字体的映射；时间：2026-10-31；作者：lq。
 *
 * Calibri/Cambria 在未安装 Office 的机器上不存在，Carlito/Caladea 与它们字符宽度一致，
 * 用它兜底才能让换行位置、页数与 Word 打开时接近。
 */
const LATIN_FONT_STACKS: Record<string, string> = {
  calibri: '"Calibri", "Carlito", "Segoe UI", system-ui, sans-serif',
  carlito: '"Carlito", "Calibri", system-ui, sans-serif',
  cambria: '"Cambria", "Caladea", "Georgia", serif',
  caladea: '"Caladea", "Cambria", "Georgia", serif',
  'times new roman': '"Times New Roman", "Liberation Serif", "Songti SC", serif',
  arial: '"Arial", "Helvetica Neue", "Liberation Sans", sans-serif',
  helvetica: '"Helvetica Neue", "Arial", sans-serif',
  'helvetica neue': '"Helvetica Neue", "Arial", sans-serif',
  'courier new': '"Courier New", "Courier", monospace',
  consolas: '"Consolas", "Menlo", monospace',
  'segoe ui': '"Segoe UI", system-ui, sans-serif',
  verdana: '"Verdana", "Geneva", sans-serif',
  georgia: '"Georgia", "Caladea", serif',
  garamond: '"Garamond", "Caladea", serif',
  'trebuchet ms': '"Trebuchet MS", "Helvetica Neue", sans-serif',
  'comic sans ms': '"Comic Sans MS", "Chalkboard SE", cursive',
};

/** 功能：候选字体清单，用于探测本机已安装字体；时间：2026-10-31；作者：lq。 */
export const FONT_PROBES: readonly string[] = [
  '宋体',
  '新宋体',
  '仿宋',
  '楷体',
  '黑体',
  '微软雅黑',
  '等线',
  '幼圆',
  '隶书',
  '华文宋体',
  '华文中宋',
  '华文仿宋',
  '华文楷体',
  '华文细黑',
  '思源宋体',
  '思源黑体',
  'PingFang SC',
  'Songti SC',
  'Heiti SC',
  'Kaiti SC',
  'STSong',
  'STFangsong',
  'STKaiti',
  'STHeiti',
  'Calibri',
  'Cambria',
  'Arial',
  'Times New Roman',
  'Courier New',
  'Consolas',
  'Georgia',
  'Verdana',
  'Trebuchet MS',
  'Segoe UI',
  'Helvetica Neue',
  'Caladea',
  'Carlito',
];

/**
 * 功能：用 document.fonts.check 探测本机真正可用的字体；时间：2026-10-31；作者：lq。
 *
 * 浏览器不能直接枚举系统字体，但可以逐个探测候选族是否可渲染。
 * 探测失败时退回内置清单，保证下拉永远有选项。
 */
export function detectInstalledFonts(): string[] {
  if (typeof document === 'undefined' || typeof document.fonts?.check !== 'function') return [...FONT_CHOICES];
  const available: string[] = [];
  for (const name of FONT_PROBES) {
    try {
      if (document.fonts.check(`16px "${name}"`)) available.push(name);
    } catch {
      // 非法字体名直接跳过
    }
  }
  return available.length ? available : [...FONT_CHOICES];
}

/**
 * 功能：字体下拉的最终选项——本机可用字体 + 当前文档正在用的字体；时间：2026-10-31；作者：lq。
 *
 * `installed` 传入平台适配器枚举到的真实字体清单（桌面端可拿到全部系统字体），
 * 拿不到时才退回到逐个探测；两者都为空时使用内置清单，保证下拉永远有选项。
 */
export function availableFontChoices(current?: string, installed?: string[]): string[] {
  const detected = installed && installed.length > 0 ? installed : detectInstalledFonts();
  const list = ['默认', ...detected.filter((name) => name && name !== '默认')];
  if (current && !list.includes(current)) list.push(current);
  return list;
}

/**
 * 功能：CJK 字体按字形风格分成几类回退栈；时间：2026-10-31；作者：lq。
 *
 * 之前所有未知中文字体都回退到 DOC_FONT_STACK（以 PingFang SC 打头，属于黑体类），
 * 于是宋体、仿宋、楷体文档在没装对应字体的机器上会整体变成黑体，字形与原件差别很大。
 * 按风格分类后，至少能保证「衬线还是无衬线」「仿宋/楷体的书写感」这两层是对的。
 */
const CJK_STACKS = {
  song: '"Songti SC", "SimSun", "Noto Serif CJK SC", "Source Han Serif SC", serif',
  fangsong: '"FangSong", "STFangsong", "FangSong_GB2312", "Songti SC", serif',
  kai: '"Kaiti SC", "KaiTi", "STKaiti", "Kaiti_GB2312", serif',
  li: '"Baoli SC", "STLiti", "Kaiti SC", serif',
  hei: '"PingFang SC", "Heiti SC", "SimHei", "Noto Sans CJK SC", "Microsoft YaHei", sans-serif',
  round: '"Yuanti SC", "YouYuan", "PingFang SC", "Noto Sans CJK SC", sans-serif',
  mono: '"SF Mono", "Menlo", "Consolas", "Courier New", monospace',
} as const;

/** 功能：常见中文字体名 → 回退栈；覆盖 Office 自带、华文、方正、GB2312 等别名；时间：2026-10-31；作者：lq。 */
const CJK_FONT_STACKS: Record<string, string> = {
  宋体: CJK_STACKS.song,
  新宋体: CJK_STACKS.song,
  '宋体-简': CJK_STACKS.song,
  '宋体-繁': CJK_STACKS.song,
  华文宋体: '"STSong", "Songti SC", "SimSun", serif',
  华文中宋: '"STZhongsong", "Songti SC", "SimSun", serif',
  方正小标宋简体: '"STZhongsong", "Songti SC", "SimSun", serif',
  方正书宋简体: CJK_STACKS.song,
  仿宋: CJK_STACKS.fangsong,
  '仿宋_GB2312': CJK_STACKS.fangsong,
  华文仿宋: '"STFangsong", "FangSong", serif',
  方正仿宋简体: CJK_STACKS.fangsong,
  楷体: CJK_STACKS.kai,
  '楷体_GB2312': CJK_STACKS.kai,
  华文楷体: '"STKaiti", "Kaiti SC", serif',
  方正楷体简体: CJK_STACKS.kai,
  隶书: CJK_STACKS.li,
  华文隶书: '"STLiti", "Baoli SC", serif',
  黑体: CJK_STACKS.hei,
  华文细黑: '"STHeiti", "Heiti SC", "PingFang SC", sans-serif',
  微软雅黑: '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif',
  '微软雅黑 Light': '"Microsoft YaHei", "PingFang SC", sans-serif',
  等线: '"DengXian", "PingFang SC", sans-serif',
  '等线 Light': '"DengXian", "PingFang SC", sans-serif',
  幼圆: CJK_STACKS.round,
  思源宋体: '"Source Han Serif SC", "Noto Serif CJK SC", "Songti SC", serif',
  思源黑体: '"Source Han Sans SC", "Noto Sans CJK SC", "PingFang SC", sans-serif',
};

/** 功能：按字体名里的字形关键字推断回退栈，覆盖表里没列出的厂商字体；时间：2026-10-31；作者：lq。 */
function heuristicCjkStack(name: string): string {
  if (/仿宋|fangsong/i.test(name)) return CJK_STACKS.fangsong;
  if (/楷|kai/i.test(name)) return CJK_STACKS.kai;
  if (/隶|lishu|liti/i.test(name)) return CJK_STACKS.li;
  if (/圆|yuanti|youyuan/i.test(name)) return CJK_STACKS.round;
  if (/宋|song|明朝|mincho/i.test(name)) return CJK_STACKS.song;
  if (/黑|hei|yahei|雅黑|sans|gothic/i.test(name)) return CJK_STACKS.hei;
  if (/mono|code|consol|courier/i.test(name)) return CJK_STACKS.mono;
  return DOC_FONT_STACK;
}

/**
 * 功能：把文档里的字体名换成浏览器可用的字体栈；时间：2026-10-31；作者：lq。
 *
 * 顺序：西文度量兼容映射 → 中文字体表 → 字形关键字推断。
 * 未命中的字体仍把它自己放在栈首，本机装了就直接用，没装才走回退。
 */
export function fontStackFor(name?: string): string {
  if (!name || name === '默认') return DOC_FONT_STACK;
  const key = name.trim();
  if (!key) return DOC_FONT_STACK;
  const latin = LATIN_FONT_STACKS[key.toLowerCase()];
  if (latin) return latin;
  const mapped = CJK_FONT_STACKS[key];
  // 功能：本机装了同名中文字体时优先使用它（如方正小标宋简体），只有缺字时才走回退栈；时间：2026-11-01；作者：lq。
  if (mapped) return mapped.startsWith(`"${key}"`) ? mapped : `"${key}", ${mapped}`;
  return `"${key}", ${heuristicCjkStack(key)}`;
}

/* -------------------------------------------------------------------------- */
/* Accent palettes                                                            */
/* -------------------------------------------------------------------------- */

export interface AccentShades {
  /** Primary accent used for filled controls and links. */
  accent: string;
  /** Slightly deeper / brighter accent for hover states. */
  hover: string;
  /** Quiet tint used for selected rows, chips and soft buttons. */
  soft: string;
  /** Text/icon color that stays readable on top of `accent`. */
  ink: string;
}

export interface AccentPreset {
  id: Exclude<AccentId, 'custom'>;
  label: string;
  description: string;
  light: AccentShades;
  dark: AccentShades;
}

/**
 * Five calm, low-glare palettes. Each one ships a hand-tuned light and dark
 * variant so the UI never has to guess contrast from a single seed color.
 *
 * `mist` is the default: a deep, slightly desaturated teal that reads as
 * "quiet stationery" rather than a generic Office blue.
 */
export const ACCENT_PRESETS: AccentPreset[] = [
  {
    id: 'mist',
    label: '雾青',
    description: '沉静的松石绿，长时间阅读不刺眼',
    light: { accent: '#2E7D72', hover: '#25665D', soft: '#E6F2EF', ink: '#FFFFFF' },
    dark: { accent: '#6CC9BB', hover: '#8ADBD0', soft: '#1B3A36', ink: '#10221F' },
  },
  {
    id: 'ocean',
    label: '远山蓝',
    description: '更克制的靛蓝，保留熟悉的 Office 感',
    light: { accent: '#2F63E0', hover: '#2552C2', soft: '#E9EFFF', ink: '#FFFFFF' },
    dark: { accent: '#8FB2FF', hover: '#ADC7FF', soft: '#21304F', ink: '#111B2E' },
  },
  {
    id: 'iris',
    label: '紫藤',
    description: '柔和的紫罗兰，适合创意与演示场景',
    light: { accent: '#6B5BD6', hover: '#5949BD', soft: '#F0ECFC', ink: '#FFFFFF' },
    dark: { accent: '#B7A8FF', hover: '#CDC2FF', soft: '#2C2946', ink: '#171428' },
  },
  {
    id: 'pine',
    label: '松针',
    description: '自然深绿，安静而专注',
    light: { accent: '#2F7D4F', hover: '#256640', soft: '#E9F4EE', ink: '#FFFFFF' },
    dark: { accent: '#7BD3A0', hover: '#98E0B7', soft: '#1B3728', ink: '#0F2118' },
  },
  {
    id: 'clay',
    label: '陶土',
    description: '温暖的赤陶色，给文档一点手作感',
    light: { accent: '#B25F4A', hover: '#974C3A', soft: '#F9EDE9', ink: '#FFFFFF' },
    dark: { accent: '#F0A892', hover: '#F6C0AE', soft: '#3B2822', ink: '#2A1812' },
  },
];

export const DEFAULT_ACCENT: AccentId = 'mist';

export function normalizeAccentId(value: string | undefined | null): AccentId {
  if (value === 'custom') return 'custom';
  return ACCENT_PRESETS.some((preset) => preset.id === value) ? (value as AccentId) : DEFAULT_ACCENT;
}

/** Normalize `#abc` / `#aabbcc` to uppercase `#AABBCC`; returns null when invalid. */
export function normalizeAccentHex(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const input = value.trim();
  if (!/^#?[0-9a-f]{3}([0-9a-f]{3})?$/i.test(input)) return null;
  const hex = input.replace(/^#/, '');
  const expanded =
    hex.length === 3
      ? hex
          .split('')
          .map((char) => `${char}${char}`)
          .join('')
      : hex;
  return `#${expanded.toUpperCase()}`;
}

function hexToRgb(hex: string): [number, number, number] {
  const value = normalizeAccentHex(hex) ?? DEFAULT_ACCENT_HEX;
  return [
    Number.parseInt(value.slice(1, 3), 16),
    Number.parseInt(value.slice(3, 5), 16),
    Number.parseInt(value.slice(5, 7), 16),
  ];
}

function channelToHex(channel: number): string {
  return Math.round(Math.max(0, Math.min(255, channel)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();
}

/** Mix `amount` (0–1) of `to` into `from`. */
export function mixHex(from: string, to: string, amount: number): string {
  const [r1, g1, b1] = hexToRgb(from);
  const [r2, g2, b2] = hexToRgb(to);
  const ratio = Math.max(0, Math.min(1, amount));
  return `#${channelToHex(r1 + (r2 - r1) * ratio)}${channelToHex(g1 + (g2 - g1) * ratio)}${channelToHex(
    b1 + (b2 - b1) * ratio,
  )}`;
}

function relativeLuminance(hex: string): number {
  const [r, g, b] = hexToRgb(hex).map((channel) => {
    const value = channel / 255;
    return value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function contrastRatio(a: string, b: string): number {
  const [light, dark] = [relativeLuminance(a), relativeLuminance(b)].sort((x, y) => y - x);
  return (light + 0.05) / (dark + 0.05);
}

/** Pick the ink with the better contrast against the accent. */
export function readableInk(accent: string): string {
  return contrastRatio(accent, '#FFFFFF') >= contrastRatio(accent, '#1D2939') ? '#FFFFFF' : '#1D2939';
}

function customShades(seed: string | undefined, theme: ResolvedTheme): AccentShades {
  const base = normalizeAccentHex(seed) ?? DEFAULT_ACCENT_HEX;
  if (theme === 'light') {
    return {
      accent: base,
      hover: mixHex(base, '#000000', 0.16),
      soft: mixHex(base, '#FFFFFF', 0.9),
      ink: readableInk(base),
    };
  }
  const accent = mixHex(base, '#FFFFFF', 0.34);
  return {
    accent,
    hover: mixHex(base, '#FFFFFF', 0.46),
    soft: mixHex(base, '#262C36', 0.8),
    ink: readableInk(accent),
  };
}

export function resolveAccentPalette(
  accent: AccentId | string | undefined | null,
  customAccent: string | undefined | null,
  theme: ResolvedTheme,
): AccentShades {
  const id = normalizeAccentId(accent);
  if (id === 'custom') return customShades(customAccent ?? undefined, theme);
  const preset = ACCENT_PRESETS.find((item) => item.id === id) ?? ACCENT_PRESETS[0];
  return preset[theme];
}

export function accentPresetById(accent: AccentId | string | undefined | null): AccentPreset | undefined {
  const id = normalizeAccentId(accent);
  return ACCENT_PRESETS.find((preset) => preset.id === id);
}

/** 功能：外观配置；护眼模式为可选项，缺省视为关闭，便于旧调用点无需改动；时间：2026-10-30；作者：lq。 */
export type AppearanceConfig = Pick<AppSettings, 'theme' | 'accent' | 'customAccent'> & {
  eyeCare?: boolean;
};

let lastAccent: { accent: AccentId; customAccent: string } = {
  accent: DEFAULT_ACCENT,
  customAccent: DEFAULT_ACCENT_HEX,
};

/**
 * Apply an accent palette to the document root. All values are CSS custom
 * properties, so the rest of the UI stays theme-agnostic.
 */
export function applyAccent(
  accent: AccentId | string | undefined,
  customAccent: string | undefined,
  resolved: ResolvedTheme,
): AccentShades {
  const id = normalizeAccentId(accent);
  const shades = resolveAccentPalette(id, customAccent, resolved);
  lastAccent = {
    accent: id,
    customAccent: normalizeAccentHex(customAccent) ?? DEFAULT_ACCENT_HEX,
  };
  if (typeof document === 'undefined') return shades;
  const root = document.documentElement;
  root.dataset.accent = id;
  root.style.setProperty('--mo-accent', shades.accent);
  root.style.setProperty('--mo-accent-hover', shades.hover);
  root.style.setProperty('--mo-accent-soft', shades.soft);
  root.style.setProperty('--mo-accent-ink', shades.ink);
  return shades;
}

export function resolveTheme(mode: ThemeMode): ResolvedTheme {
  if (mode === 'dark') return 'dark';
  if (mode === 'light') return 'light';
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

export function applyTheme(mode: ThemeMode): ResolvedTheme {
  if (typeof document === 'undefined') return 'light';
  const resolved = resolveTheme(mode);
  document.documentElement.dataset.theme = resolved;
  document.documentElement.style.colorScheme = resolved;
  // Keep the accent in sync when the resolved theme changes (system mode).
  applyAccent(lastAccent.accent, lastAccent.customAccent, resolved);
  return resolved;
}

/** Apply theme + accent together; this is the entry point used by the app. */
export function applyAppearance(config: AppearanceConfig): ResolvedTheme {
  lastAccent = {
    accent: normalizeAccentId(config.accent),
    customAccent: normalizeAccentHex(config.customAccent) ?? DEFAULT_ACCENT_HEX,
  };
  // 功能：护眼模式只切换 data-eyecare，配色全部由 tokens.css 承担，方便主题插件叠加；时间：2026-10-30；作者：lq。
  if (typeof document !== 'undefined') {
    document.documentElement.dataset.eyecare = config.eyeCare ? 'on' : 'off';
  }
  return applyTheme(config.theme);
}

let mediaCleanup: (() => void) | null = null;

export function watchSystemTheme(mode: ThemeMode): () => void {
  if (typeof window === 'undefined') return () => {};
  mediaCleanup?.();
  const media = window.matchMedia('(prefers-color-scheme: dark)');
  const onChange = () => {
    if (mode === 'system') applyTheme('system');
  };
  media.addEventListener('change', onChange);
  mediaCleanup = () => media.removeEventListener('change', onChange);
  return mediaCleanup;
}

export const MOTION = {
  fast: 'var(--mo-motion-fast)',
  base: 'var(--mo-motion-base)',
} as const;
