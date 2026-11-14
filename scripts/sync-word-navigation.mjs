/** 功能：只同步固定版本 GenOffice 的 jumpToAnchor 原函数体，不重写已接线的解析、分页和加载适配；时间：2026-11-14；作者：lq。 */
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const upstreamPath = 'apps/docs/src/renderer/App.tsx';
const localPath = 'packages/word-core/src/upstream/apps/docs/src/renderer/anchor-navigation.ts';
const source = readFileSync(resolve(root, 'output/word-core-m0.local/upstream', upstreamPath));
const sourceSha256 = createHash('sha256').update(source).digest('hex');
const baseline = JSON.parse(readFileSync(resolve(root, 'docs/word-core-m0/upstream.json'), 'utf8'));
if (baseline.files.find(file => file.path === upstreamPath)?.sha256 !== sourceSha256) throw new Error('GenOffice 固定来源指纹不匹配');
const ast = ts.createSourceFile(upstreamPath, source.toString('utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
let body;

/** 功能：按变量身份提取上游 useCallback 回调体，拒绝猜测函数位置；时间：2026-11-14；作者：lq。 */
function visit(node) {
  if (ts.isVariableDeclaration(node) && node.name.getText(ast) === 'jumpToAnchor' && node.initializer && ts.isCallExpression(node.initializer)) {
    const callback = node.initializer.arguments[0];
    if (callback && ts.isArrowFunction(callback) && ts.isBlock(callback.body)) body = callback.body.getText(ast).slice(1, -1);
  }
  ts.forEachChild(node, visit);
}
visit(ast);
if (!body?.includes("target?.scrollIntoView({ behavior: 'smooth', block: 'start' })")) throw new Error('上游目录跳转函数已变化');
const local = `/** 功能：同步 GenOffice 目录导航；时间：2026-11-14；作者：lq；上游版本及 Apache-2.0 许可见 UPSTREAM.md。 */
import type { ParsedDocFull } from '../../../../packages/docx-engine/src/index';
/** 功能：保留上游函数体，仅注入当前文档和真实编辑根，避免跨标签或快照误命中；时间：2026-11-14；作者：lq。 */
export function jumpToWordAnchor(root: HTMLElement, parsed: ParsedDocFull, anchor: string, tocTitle = ''): Element | null {
  const doc = { parsed };
  const document = root;
${body}
  return target;
}
`;
writeFileSync(resolve(root, localPath), local);
const manifestPath = resolve(root, 'packages/word-core/SOURCE-MANIFEST.json');
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
manifest.extractedFiles = manifest.extractedFiles.filter(file => file.localPath !== localPath);
manifest.extractedFiles.push({ upstreamPath, localPath, upstreamSha256: sourceSha256, localSha256: createHash('sha256').update(local).digest('hex'), modifications: ['从固定 App.tsx 提取 jumpToAnchor 原函数体，仅注入文档和查询根并返回命中节点'] });
writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n');
console.log(`Synced jumpToAnchor from ${baseline.commit}`);
