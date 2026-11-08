/** 功能：记录和核验 M0 工作区、上游源码、锁文件及资源基线；时间：2026-11-08；作者：lq。 */
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const local = join(root, 'output/word-core-m0.local');
const upstream = resolve(process.env.MINI_OFFICE_WORD_UPSTREAM ?? join(local, 'upstream'));
const destination = join(root, 'docs/word-core-m0');
const commit = 'e1f50e90d246f3ecc6427d2a0af90f04f37810ec';
const archive = join(local, `genoffice-${commit}.tar.gz`);
const comment = '功能：固定 Word 定制内核 M0 来源与回归基线；时间：2026-11-08；作者：lq。';

/** 功能：计算原始文件的 SHA-256，不读取私人文档；时间：2026-11-08；作者：lq。 */
function sha(path) { return createHash('sha256').update(readFileSync(path)).digest('hex'); }
/** 功能：以结构化参数执行只读 Git 查询；时间：2026-11-08；作者：lq。 */
function git(...args) { return execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim(); }
/** 功能：输出可复核 JSON，拒绝覆盖首次记录；时间：2026-11-08；作者：lq。 */
function record(name, value) {
  const target = join(destination, name);
  if (existsSync(target)) throw new Error(`基线已存在，拒绝覆盖：${target}`);
  writeFileSync(target, JSON.stringify({ _comment: comment, schemaVersion: 1, ...value }, null, 2) + '\n');
}
/** 功能：按稳定顺序枚举目录中的普通文件，拒绝符号链接；时间：2026-11-08；作者：lq。 */
function files(directory) {
  return readdirSync(directory, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name, 'en')).flatMap(entry => {
    const path = join(directory, entry.name);
    if (entry.isSymbolicLink()) throw new Error(`不接受符号链接：${path}`);
    return entry.isDirectory() ? files(path) : entry.isFile() ? [path] : [];
  });
}
/** 功能：从 TypeScript AST 提取静态、动态和类型导入，供 M1 划分运行依赖；时间：2026-11-08；作者：lq。 */
function importsOf(path) {
  const source = ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true);
  const imports = [];
  /** 功能：递归读取模块说明符，保留原始相对路径和类型导入标记；时间：2026-11-08；作者：lq。 */
  const visit = node => {
    if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) {
      imports.push({ specifier: node.moduleSpecifier.text, typeOnly: !!(node.isTypeOnly || node.importClause?.isTypeOnly), kind: ts.isImportDeclaration(node) ? 'import' : 'export' });
    } else if (ts.isCallExpression(node) && node.arguments[0] && ts.isStringLiteral(node.arguments[0]) && (node.expression.kind === ts.SyntaxKind.ImportKeyword || (ts.isIdentifier(node.expression) && node.expression.text === 'require'))) {
      imports.push({ specifier: node.arguments[0].text, typeOnly: false, kind: 'dynamic' });
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return imports;
}
/** 功能：枚举现有 Word UI 回调、静态菜单、可访问名称和键盘条件，避免手工清单遗漏入口；时间：2026-11-08；作者：lq。 */
function commands() {
  const paths = [...files(join(root, 'packages/editor-word/src')).filter(path => /\.(?:ts|tsx)$/.test(path)), ...['AppShell.tsx', 'CommandPalette.tsx', 'chrome.tsx'].map(name => join(root, 'packages/editor-shell/src', name))];
  const sources = paths.map(path => {
    const source = ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true);
    const entries = [];
    /** 功能：保留源码行号与入口类别，只收集实际源码中的静态标识；时间：2026-11-08；作者：lq。 */
    const visit = node => {
      const line = source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1;
      if (ts.isInterfaceDeclaration(node) && node.name.text === 'WordToolbarProps') {
        for (const member of node.members) if (member.name?.getText(source).startsWith('on')) entries.push({ kind: 'toolbar-callback', name: member.name.getText(source), line: source.getLineAndCharacterOfPosition(member.getStart(source)).line + 1 });
      }
      if (ts.isObjectLiteralExpression(node)) {
        const fields = node.properties.filter(ts.isPropertyAssignment);
        const id = fields.find(field => field.name.getText(source) === 'id');
        const label = fields.find(field => field.name.getText(source) === 'label');
        if (id && label && ts.isStringLiteral(id.initializer) && ts.isStringLiteral(label.initializer)) entries.push({ kind: 'menu', name: id.initializer.text, label: label.initializer.text, line });
      }
      if (ts.isJsxAttribute(node) && ['aria-label', 'label'].includes(node.name.getText(source)) && node.initializer && ts.isStringLiteral(node.initializer)) entries.push({ kind: 'ui-label', name: node.initializer.text, line });
      if (ts.isIfStatement(node) && /(?:event|e)\.key\b/.test(node.expression.getText(source))) entries.push({ kind: 'keyboard-condition', name: node.expression.getText(source), line });
      ts.forEachChild(node, visit);
    };
    visit(source);
    return { path: relative(root, path), sha256: sha(path), entries };
  }).filter(source => source.entries.length);
  record('commands.json', { scope: '静态源码入口审计，动态菜单、组合快捷键和运行条件另见功能迁移清单；不是行为测试通过证明。', sources });
  console.log(JSON.stringify({ status: 'recorded', commandSources: sources.length, entries: sources.reduce((sum, source) => sum + source.entries.length, 0) }));
}
/** 功能：核验首次记录的生产代码、上游文件和归档，检查 UI 图片哈希；时间：2026-11-08；作者：lq。 */
function verify() {
  const baseline = JSON.parse(readFileSync(join(destination, 'workspace.json'), 'utf8'));
  const origin = JSON.parse(readFileSync(join(destination, 'upstream.json'), 'utf8'));
  const errors = [];
  for (const file of baseline.protectedFiles) {
    if (!existsSync(join(root, file.path)) || sha(join(root, file.path)) !== file.sha256) errors.push(`工作区文件变化：${file.path}`);
  }
  for (const file of origin.files) {
    if (!existsSync(join(upstream, file.path)) || sha(join(upstream, file.path)) !== file.sha256) errors.push(`上游文件变化或缺失：${file.path}`);
  }
  // 功能：静态命令清单也必须对应同一份源码，不把过期行号当作当前入口；时间：2026-11-08；作者：lq。
  const commandPath = join(destination, 'commands.json');
  if (!existsSync(commandPath)) errors.push('缺少现有命令基线');
  else for (const source of JSON.parse(readFileSync(commandPath, 'utf8')).sources) {
    if (!existsSync(join(root, source.path)) || sha(join(root, source.path)) !== source.sha256) errors.push(`命令入口来源变化：${source.path}`);
  }
  if (!existsSync(archive) || sha(archive) !== origin.archive.sha256) errors.push('固定源码归档变化或缺失');
  const uiPath = join(root, 'output/playwright/word-core-m0/manifest.json');
  if (!existsSync(uiPath)) errors.push('缺少本次 UI 截图基线');
  else {
    const ui = JSON.parse(readFileSync(uiPath, 'utf8'));
    if (ui.workspaceCommit !== baseline.commit || ui.sample.sha256 !== sha(join(root, ui.sample.path))) errors.push('UI 基线版本或公共样本不一致');
    for (const shot of ui.screenshots) if (!existsSync(join(root, shot.path)) || sha(join(root, shot.path)) !== shot.sha256) errors.push(`UI 截图变化或缺失：${shot.path}`);
    if (ui.errors.length) errors.push('UI 捕获存在浏览器异常');
  }
  const result = { status: errors.length ? 'failed' : 'passed', workspaceCommit: baseline.commit, protectedFiles: baseline.protectedFiles.length, upstreamCommit: origin.commit, upstreamFiles: origin.files.length, errors };
  console.log(JSON.stringify(result, null, 2));
  if (errors.length) process.exitCode = 1;
}

if (process.argv[2] === '--verify') verify();
else if (process.argv[2] === '--commands') commands();
else if (process.argv[2] === '--record') {
  mkdirSync(destination, { recursive: true });
  const start = JSON.parse(readFileSync(join(local, 'start.json'), 'utf8'));
  if (start.commit !== git('rev-parse', 'HEAD')) throw new Error('开发起点已变化，不能混合记录');
  const listing = execFileSync('tar', ['-tzf', archive], { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }).trim().split('\n');
  if (listing.some(path => !path.startsWith(`genoffice-${commit}/`) || path.split('/').includes('..'))) throw new Error('归档根目录与固定提交不符');
  const roots = [
    'apps/docs', 'packages/docx-engine', 'packages/font-metrics', 'packages/doc2docx',
    'packages/zip-gate', 'packages/agent-core', 'packages/pptx-engine', 'packages/i18n',
    'packages/ui', 'packages/electron-utils',
  ];
  const rootFiles = ['package.json', 'package-lock.json', 'LICENSE', 'NOTICE'];
  const paths = [...rootFiles.map(path => join(upstream, path)), ...roots.flatMap(path => files(join(upstream, path)))];
  const inventory = paths.map(path => ({ path: relative(upstream, path), bytes: statSync(path).size, sha256: sha(path) })).sort((a, b) => a.path.localeCompare(b.path, 'en'));
  const lock = JSON.parse(readFileSync(join(upstream, 'package-lock.json'), 'utf8'));
  const packages = [];
  /** 功能：记录 workspace 包原始声明及每个直接依赖的锁定位置，不安装或启用依赖；时间：2026-11-08；作者：lq。 */
  for (const directory of ['apps', 'packages']) {
    for (const entry of readdirSync(join(upstream, directory), { withFileTypes: true })) {
      const path = `${directory}/${entry.name}/package.json`;
      if (!entry.isDirectory() || !existsSync(join(upstream, path))) continue;
      const pkg = JSON.parse(readFileSync(join(upstream, path), 'utf8'));
      packages.push({ path, name: pkg.name, version: pkg.version, dependencies: pkg.dependencies ?? {}, optionalDependencies: pkg.optionalDependencies ?? {}, peerDependencies: pkg.peerDependencies ?? {}, devDependencies: pkg.devDependencies ?? {}, sha256: sha(join(upstream, path)) });
    }
  }
  const lockEntries = Object.entries(lock.packages ?? {}).map(([path, pkg]) => ({ path, version: pkg.version ?? null, link: pkg.link === true, resolved: pkg.resolved ?? null, integrity: pkg.integrity ?? null, dependencies: pkg.dependencies ?? {}, optionalDependencies: pkg.optionalDependencies ?? {}, peerDependencies: pkg.peerDependencies ?? {} }));
  const imports = paths.filter(path => /\.(?:ts|tsx|mts|cts)$/.test(path) && !/\.d\.ts$/.test(path)).map(path => ({ path: relative(upstream, path), imports: importsOf(path) }));
  // 功能：基线只保护原有受版本控制的源码和锁文件，不采集业务输出或私人文件；时间：2026-11-08；作者：lq。
  const protectedFiles = git('ls-files', '-z').split('\0').filter(path => path && (/^(?:apps\/desktop|packages|crates|scripts|tests)\//.test(path) || /^(?:package\.json|pnpm-lock\.yaml|pnpm-workspace\.yaml|Cargo\.(?:toml|lock)|tsconfig\.json|vitest\.config\.ts)$/.test(path))).map(path => ({ path, sha256: sha(join(root, path)) }));
  record('workspace.json', { commit: start.commit, branch: start.branch, capturedAt: start.capturedAt, initialStatus: start.status, cleanAtStart: start.status === '', packageVersion: JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version, nodeVersion: process.versions.node, protectedFiles });
  record('upstream.json', { repository: 'https://github.com/genspark-ai/genoffice', commit, userValidatedVersionClaimed: false, retrieval: 'GitHub codeload 固定提交归档；每个文件 SHA-256 另行记录', archive: { path: relative(root, archive), url: `https://codeload.github.com/genspark-ai/genoffice/tar.gz/${commit}`, sha256: sha(archive), bytes: statSync(archive).size }, roots, rootsAreShippingSelection: false, files: inventory });
  record('dependencies.json', { commit, lockFile: 'package-lock.json', lockVersion: lock.lockfileVersion, lockSha256: sha(join(upstream, 'package-lock.json')), inventoryMeaning: '原始包声明、锁定后的直接/传递包记录与源文件导入；不代表这些依赖全部需要引入产品。M1 按运行时闭包提取。', workspacePackages: packages.sort((a, b) => a.path.localeCompare(b.path, 'en')), lockedPackages: lockEntries });
  record('imports.json', { commit, inventoryMeaning: '原始 TS 静态/动态导入和 re-export；混合导入需在 M1 逐项检查，CSS/worker/字体资源另见 upstream.json；不宣称为已提取的运行时依赖闭包。', files: imports });
  console.log(JSON.stringify({ status: 'recorded', workspaceCommit: start.commit, upstreamCommit: commit, upstreamFiles: inventory.length, sourceModules: imports.length, lockedPackages: lockEntries.length, protectedFiles: protectedFiles.length }));
} else throw new Error('用法：node scripts/word-core-baseline.mjs --record、--commands 或 --verify');
