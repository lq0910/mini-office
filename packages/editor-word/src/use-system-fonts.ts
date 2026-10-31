import { useEffect, useState } from 'react';
import { getPlatformAdapter } from '@mini-office/platform-adapter';
import { FONT_CHOICES } from '@mini-office/design-system';

/**
 * 功能：系统字体列表——桌面端走 Tauri 的 list_fonts（扫描系统字体注册表/目录），
 * Web 端退回 canvas 度量探测；模块级缓存，整个会话只取一次；时间：2026-10-30；作者：hou0。
 */
let cachedFonts: string[] | null = null;

export function useSystemFonts(): string[] {
  const [fonts, setFonts] = useState<string[]>(() => cachedFonts ?? []);
  useEffect(() => {
    if (cachedFonts) return;
    let cancelled = false;
    getPlatformAdapter()
      .listFonts()
      .then((list) => {
        if (cancelled) return;
        cachedFonts = list.map((font) => font.family).filter((family) => !!family);
        setFonts(cachedFonts);
      })
      .catch(() => {
        /* 探测失败时保留内置字体列表 */
      });
    return () => {
      cancelled = true;
    };
  }, []);
  return fonts;
}

/**
 * 功能：字体下拉框选项合并——顺序为「默认 / 当前字体 / 文档默认 / 系统字体 / 内置列表」，
 * 大小写不敏感去重；当前字体即使不在任何列表里（导入文档带入了系统里的字体）也会出现在
 * 下拉框中，避免选择框显示空白；时间：2026-10-30；作者：hou0。
 */
export function mergeFontChoices(systemFonts: string[], current?: string, docDefault?: string): string[] {
  const result: string[] = ['默认'];
  const seen = new Set(['默认']);
  const push = (name: string | undefined) => {
    if (!name || name === '默认') return;
    const key = name.trim().toLowerCase();
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push(name.trim());
  };
  push(current);
  push(docDefault);
  for (const font of systemFonts) push(font);
  for (const font of FONT_CHOICES) push(font);
  return result;
}
