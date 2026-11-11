/**
 * 功能：段落设置对话框，按 Word / WPS「段落」对话框的信息结构组织缩进与间距；时间：2026-10-29；作者：lq。
 *
 * 为什么单独做对话框：工具栏只能放高频的少量取值，而固定行距的磅值、悬挂缩进、
 * 段前段后间距都需要精确输入。这里一次提交整份补丁，编辑器只产生一条撤销记录。
 */
import { useEffect, useRef, useState } from 'react';
import type { ParagraphStyle, WordAlignment } from '@mini-office/document-model';
import { Button, Dialog, Field, Input, Select, Switch } from '@mini-office/ui';
import {
  charWidthPx,
  formatPt,
  ptToPx,
  pxToChars,
  pxToPt,
  readLineSpacing,
  readSpecialIndent,
  writeLineSpacing,
  writeSpecialIndent,
  type SpecialIndentMode,
} from './typography';

type SpacingOption = '1' | '1.15' | '1.5' | '2' | '2.5' | '3' | 'exact' | 'atLeast' | 'multiple';

const ALIGNMENTS: Array<{ value: WordAlignment; label: string }> = [
  { value: 'left', label: '左对齐' },
  { value: 'center', label: '居中' },
  { value: 'right', label: '右对齐' },
  { value: 'justify', label: '两端对齐' },
];

const SPECIAL_INDENTS: Array<{ value: SpecialIndentMode; label: string }> = [
  { value: 'none', label: '（无）' },
  { value: 'firstLine', label: '首行缩进' },
  { value: 'hanging', label: '悬挂缩进' },
];

interface Draft {
  alignment: WordAlignment;
  leftChars: string;
  rightChars: string;
  special: SpecialIndentMode;
  specialChars: string;
  beforePt: string;
  afterPt: string;
  spacing: SpacingOption;
  spacingValue: string;
  pageBreakBefore: boolean;
}

function toDraft(style: ParagraphStyle | undefined, fontSizePt: number): Draft {
  const spacing = readLineSpacing(style);
  const special = readSpecialIndent(style, fontSizePt);
  return {
    alignment: style?.alignment ?? 'left',
    leftChars: formatPt(Math.round(pxToChars(style?.indentLeftPx ?? 0, fontSizePt) * 10) / 10),
    rightChars: formatPt(Math.round(pxToChars(style?.indentRightPx ?? 0, fontSizePt) * 10) / 10),
    special: special.mode,
    specialChars: special.mode === 'none' ? '2' : formatPt(special.chars),
    beforePt: formatPt(Math.round(pxToPt(style?.spaceBeforePx ?? 0) * 10) / 10),
    afterPt: formatPt(Math.round(pxToPt(style?.spaceAfterPx ?? 0) * 10) / 10),
    spacing:
      spacing.mode === 'exact'
        ? 'exact'
        : spacing.mode === 'atLeast'
          ? 'atLeast'
          : (String(spacing.value) as SpacingOption),
    spacingValue: formatPt(spacing.value),
    pageBreakBefore: !!style?.pageBreakBefore,
  };
}

export function ParagraphDialog({
  open,
  style,
  fontSizePt,
  selectionHint,
  onClose,
  onApply,
  applyChangedOnly,
}: {
  open: boolean;
  style: ParagraphStyle | undefined;
  fontSizePt: number;
  selectionHint?: string;
  onClose: () => void;
  onApply: (patch: Partial<ParagraphStyle>) => void;
  /** 功能：新核心只提交用户修改的字段，避免混合选区被初始值批量覆盖；时间：2026-11-10；作者：lq。 */
  applyChangedOnly?: boolean;
}) {
  const [draft, setDraft] = useState<Draft>(() => toDraft(style, fontSizePt));
  // 功能：记录对话框内实际修改的字段，不将未编辑的混合值写回正文；时间：2026-11-10；作者：lq。
  const changed = useRef(new Set<keyof Draft>());

  useEffect(() => {
    if (open) { setDraft(toDraft(style, fontSizePt)); changed.current.clear(); }
    // 打开时以当前段落为初值；对话框打开期间不跟随外部变化，避免输入被覆盖。
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const patch = (next: Partial<Draft>) => {
    for (const key of Object.keys(next) as Array<keyof Draft>) changed.current.add(key);
    setDraft((current) => ({ ...current, ...next }));
  };

  const isPresetMultiple = (value: SpacingOption): boolean => /^\d+(\.\d+)?$/.test(value);

  const submit = () => {
    const spacingMode = draft.spacing === 'exact' ? 'exact' : draft.spacing === 'atLeast' ? 'atLeast' : 'multiple';
    const spacingValue = isPresetMultiple(draft.spacing)
      ? Number(draft.spacing)
      : Number(draft.spacingValue) || 1.5;
    const result: Partial<ParagraphStyle> = {
      alignment: draft.alignment,
      indentLeftPx: Math.max(0, Number(draft.leftChars) || 0) * charWidthPx(fontSizePt),
      indentRightPx: Math.max(0, Number(draft.rightChars) || 0) * charWidthPx(fontSizePt),
      ...writeSpecialIndent(draft.special, Number(draft.specialChars) || 0, fontSizePt),
      spaceBeforePx: Math.max(0, Number(draft.beforePt) || 0) * (4 / 3),
      spaceAfterPx: Math.max(0, Number(draft.afterPt) || 0) * (4 / 3),
      ...writeLineSpacing({ mode: spacingMode, value: spacingValue }),
      pageBreakBefore: draft.pageBreakBefore,
    };
    // 功能：按编辑控件映射原有字段，一次确认仍只生成一个事务；时间：2026-11-10；作者：lq。
    const keys: Record<keyof Draft, Array<keyof ParagraphStyle>> = {
      alignment: ['alignment'], leftChars: ['indentLeftPx'], rightChars: ['indentRightPx'],
      special: ['firstLineIndentPx'], specialChars: ['firstLineIndentPx'],
      beforePt: ['spaceBeforePx'], afterPt: ['spaceAfterPx'],
      spacing: ['lineHeight', 'lineHeightExactPx', 'lineHeightAtLeastPx'],
      spacingValue: ['lineHeight', 'lineHeightExactPx', 'lineHeightAtLeastPx'], pageBreakBefore: ['pageBreakBefore'],
    };
    const edited = new Set([...changed.current].flatMap(key => keys[key]));
    onApply(applyChangedOnly ? Object.fromEntries(Object.entries(result).filter(([key]) => edited.has(key as keyof ParagraphStyle))) : result);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="段落"
      description={selectionHint ?? '设置选中段落的缩进与间距；未选中文字时作用于光标所在段落。'}
      width="md"
      footer={
        <>
          <Button onClick={onClose}>取消</Button>
          <Button variant="primary" onClick={submit}>
            确定
          </Button>
        </>
      }
    >
      <div className="space-y-4 pb-1">
        <section className="space-y-2">
          <h3 className="text-[11px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">缩进</h3>
          <div className="grid grid-cols-2 gap-3">
            <Field label="左侧（字符）">
              <Input
                value={draft.leftChars}
                inputMode="decimal"
                onChange={(event) => patch({ leftChars: event.target.value })}
              />
            </Field>
            <Field label="右侧（字符）">
              <Input
                value={draft.rightChars}
                inputMode="decimal"
                onChange={(event) => patch({ rightChars: event.target.value })}
              />
            </Field>
            <Field label="特殊格式">
              <Select
                value={draft.special}
                onChange={(event) => patch({ special: event.target.value as SpecialIndentMode })}
              >
                {SPECIAL_INDENTS.map((item) => (
                  <option key={item.value} value={item.value}>
                    {item.label}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="缩进值（字符）">
              <Input
                value={draft.specialChars}
                inputMode="decimal"
                disabled={draft.special === 'none'}
                onChange={(event) => patch({ specialChars: event.target.value })}
              />
            </Field>
          </div>
        </section>

        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="text-[11px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">间距</h3>
          <div className="grid grid-cols-2 gap-3">
            <Field label="段前（磅）">
              <Input
                value={draft.beforePt}
                inputMode="decimal"
                onChange={(event) => patch({ beforePt: event.target.value })}
              />
            </Field>
            <Field label="段后（磅）">
              <Input
                value={draft.afterPt}
                inputMode="decimal"
                onChange={(event) => patch({ afterPt: event.target.value })}
              />
            </Field>
            <Field label="行距">
              <Select
                value={draft.spacing}
                onChange={(event) => {
                  const value = event.target.value as SpacingOption;
                  patch({
                    spacing: value,
                    // 切到固定值/最小值时给一个可用的默认磅值，切到多倍时沿用原倍数。
                    spacingValue:
                      value === 'exact' || value === 'atLeast'
                        ? formatPt(Math.round(pxToPt(fontSizePt * 1.5) * 10) / 10)
                        : formatPt(Number(draft.spacing) || 1.5),
                  });
                }}
              >
                <option value="1">单倍行距</option>
                <option value="1.15">1.15 倍行距</option>
                <option value="1.5">1.5 倍行距</option>
                <option value="2">2 倍行距</option>
                <option value="2.5">2.5 倍行距</option>
                <option value="3">3 倍行距</option>
                <option value="multiple">多倍行距</option>
                <option value="atLeast">最小值</option>
                <option value="exact">固定值</option>
              </Select>
            </Field>
            <Field label={draft.spacing === 'exact' || draft.spacing === 'atLeast' ? '设置值（磅）' : '设置值（倍）'}>
              <Input
                value={isPresetMultiple(draft.spacing) ? draft.spacing : draft.spacingValue}
                inputMode="decimal"
                onChange={(event) => {
                  if (isPresetMultiple(draft.spacing)) patch({ spacing: 'multiple', spacingValue: event.target.value });
                  else patch({ spacingValue: event.target.value });
                }}
              />
            </Field>
          </div>
        </section>

        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <div className="grid grid-cols-2 items-end gap-3">
            <Field label="对齐方式">
              <Select
                value={draft.alignment}
                onChange={(event) => patch({ alignment: event.target.value as WordAlignment })}
              >
                {ALIGNMENTS.map((item) => (
                  <option key={item.value} value={item.value}>
                    {item.label}
                  </option>
                ))}
              </Select>
            </Field>
            <label className="flex items-center justify-between pb-1 text-[12px]">
              <span>段前分页</span>
              <Switch
                checked={draft.pageBreakBefore}
                onChange={(checked) => patch({ pageBreakBefore: checked })}
                label="段前分页"
              />
            </label>
          </div>
        </section>
      </div>
    </Dialog>
  );
}
