/** 功能：统一文件图标，以折角纸张、类型字标和内容纹理区分支持的格式；时间：2026-11-07；作者：lq。 */
import type { OfficeKind } from '@mini-office/shared';

const COLORS = { docx: '#3267CB', xlsx: '#218451', pptx: '#D65C35', pdf: '#D94851', md: '#7862C8', html: '#C78325', csv: '#218A83' };
const LABELS = { docx: 'W', xlsx: 'X', pptx: 'P', pdf: 'PDF', md: 'MD', html: '</>', csv: 'CSV' };

export function FileIcon({ kind, name, size = 32, className }: { kind: OfficeKind; name?: string; size?: number; className?: string }) {
  const type = name?.toLowerCase().endsWith('.csv') ? 'csv' : kind;
  const color = COLORS[type];
  return (
    <svg width={size} height={size} viewBox="0 0 32 36" fill="none" className={className} role="img" aria-label={`${type.toUpperCase()} 文件`}>
      <path d="M8 1.5h13L29 10v22a2.5 2.5 0 0 1-2.5 2.5h-18A2.5 2.5 0 0 1 6 32V4A2.5 2.5 0 0 1 8 1.5Z" fill={color} opacity=".16" />
      <path d="M21 1.5V8a2 2 0 0 0 2 2h6" stroke={color} strokeWidth="1.5" strokeLinejoin="round" />
      {type === 'xlsx' || type === 'csv' ? <path d="M12 12h12v8H12zm4 0v8m4-8v8m-8-4h12" stroke={color} strokeWidth="1.2" /> : type === 'pptx' ? <path d="M12 12h12v8H12zm6 8v3m-3 0h6" stroke={color} strokeWidth="1.4" /> : <path d="M12 13h11m-11 4h8" stroke={color} strokeWidth="1.5" strokeLinecap="round" />}
      <rect x="1" y="21" width={type === 'docx' || type === 'xlsx' || type === 'pptx' ? 17 : 25} height="13" rx="3" fill={color} />
      <text x={type === 'docx' || type === 'xlsx' || type === 'pptx' ? 9.5 : 13.5} y="30.5" textAnchor="middle" fill="white" fontFamily="Arial, sans-serif" fontWeight="700" fontSize={type.length > 2 && type !== 'docx' && type !== 'xlsx' && type !== 'pptx' ? 8 : 10}>{LABELS[type]}</text>
    </svg>
  );
}
