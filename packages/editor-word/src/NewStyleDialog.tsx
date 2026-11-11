/**
 * 功能：新建样式对话框，以当前段落格式为蓝本命名保存；时间：2026-10-29；作者：lq。
 *
 * 与 WPS 的差异：WPS 的新建样式要经过「格式→字体→段落」多级弹窗，
 * 这里默认直接取当前段落格式，只让用户填名字，需要微调时先调好段落再新建即可。
 */
import { useEffect, useMemo, useState } from 'react';
import type { DocModel, DocParagraph } from '@mini-office/document-model';
import { Button, Dialog, Field, Input } from '@mini-office/ui';
import { fontStackFor } from '@mini-office/design-system';
import { resolveStyleRun } from './styles';

export type NewStyleBase = 'current' | 'normal';

export function NewStyleDialog({
  open,
  paragraph,
  formatPreview,
  doc,
  existingNames,
  defaultName,
  onClose,
  onCreate,
}: {
  open: boolean;
  paragraph?: DocParagraph;
  /** 功能：新核心只提供格式摘要，预览不需要伪造旧段落对象；时间：2026-11-10；作者：lq。 */
  formatPreview?: { run: import('@mini-office/document-model').RunStyle; style: import('@mini-office/document-model').ParagraphStyle };
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>;
  existingNames: string[];
  defaultName: string;
  onClose: () => void;
  onCreate: (name: string, base: NewStyleBase) => void;
}) {
  const [name, setName] = useState(defaultName);
  const [base, setBase] = useState<NewStyleBase>('current');

  useEffect(() => {
    if (!open) return;
    setName(defaultName);
    setBase('current');
  }, [open, defaultName]);

  const trimmed = name.trim();
  const duplicate = existingNames.some((item) => item === trimmed);
  const invalid = !trimmed || trimmed.length > 40 || duplicate;

  // 预览使用当前段落的实际字体与段落格式；选择「正文」基准时回落到文档默认值。
  const preview = useMemo(() => {
    const run = base === 'current' ? resolveStyleRun({ run: formatPreview?.run ?? paragraph?.runs.find((item) => item.text)?.style ?? {} }, doc) : resolveStyleRun({ run: {} }, doc);
    const style = base === 'current' ? formatPreview?.style ?? paragraph?.style ?? {} : { alignment: 'left' as const, lineHeight: 1.5 };
    return { run, style };
  }, [base, doc, paragraph, formatPreview]);

  const submit = () => {
    if (invalid) return;
    onCreate(trimmed, base);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="新建样式"
      description="以当前段落的格式创建新样式，之后可在样式库中一键套用到其他段落。"
      width="sm"
      footer={
        <>
          <Button onClick={onClose}>取消</Button>
          <Button variant="primary" disabled={invalid} onClick={submit}>
            创建
          </Button>
        </>
      }
    >
      <div className="space-y-4">
        <Field label="样式名称" hint={duplicate ? '已存在同名样式，请换一个名称。' : '最多 40 个字符'}>
          <Input
            autoFocus
            value={name}
            maxLength={60}
            onChange={(event) => setName(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                submit();
              }
            }}
            placeholder="例如：合同条款"
          />
        </Field>

        <div className="space-y-1.5">
          <span className="text-[12px] text-[var(--mo-text-primary)]">样式基准</span>
          <div className="grid grid-cols-2 gap-2">
            {([
              { id: 'current', title: '当前段落格式', hint: '沿用光标处的字体与段落设置' },
              { id: 'normal', title: '正文', hint: '从文档默认格式开始' },
            ] as const).map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={() => setBase(item.id)}
                className={`rounded-md border px-2.5 py-2 text-left transition-colors ${
                  base === item.id
                    ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)]'
                    : 'border-[var(--mo-line)] hover:bg-[var(--mo-surface-2)]'
                }`}
              >
                <span className="block text-[12px] font-medium text-[var(--mo-text-primary)]">{item.title}</span>
                <span className="mt-0.5 block text-[10px] leading-snug text-[var(--mo-text-faint)]">{item.hint}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="space-y-1.5">
          <span className="text-[12px] text-[var(--mo-text-primary)]">预览</span>
          <div className="overflow-hidden rounded-md border border-[var(--mo-line)] bg-[var(--mo-canvas)] px-3 py-3">
            <div
              style={{
                fontFamily: fontStackFor(preview.run.fontFamily ?? doc.defaultFont),
                fontSize: `${preview.run.fontSizePt ?? doc.defaultFontSizePt}pt`,
                fontWeight: preview.run.bold ? 700 : 400,
                fontStyle: preview.run.italic ? 'italic' : 'normal',
                color: preview.run.color ?? 'var(--mo-text-primary)',
                textAlign: preview.style.alignment ?? 'left',
                lineHeight: preview.style.lineHeight ?? 1.5,
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
              }}
            >
              {trimmed || '样式预览'}
            </div>
          </div>
        </div>
      </div>
    </Dialog>
  );
}
