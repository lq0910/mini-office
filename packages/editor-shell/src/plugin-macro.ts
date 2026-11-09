/**
 * 插件宏在宿主侧的落地：取当前会话 → 执行 → 写回 → 提示。
 *
 * 单独成文件而不是塞进命令面板，一是让"能不能跑"这条判断可以被测试直接覆盖，
 * 二是将来菜单栏、工具栏按钮触发同一条命令时复用同一段逻辑。
 */
import { MacroError, runMacro, usePluginRegistry, type PluginMacroCommand } from '@mini-office/plugin-host';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { collectShortcutBindings, matchShortcutFromEvent } from './shortcuts';

/** 宏能否在当前会话上执行。 */
export function canRunMacro(session: DocumentSession | undefined, macro: PluginMacroCommand): boolean {
  if (!session) return false;
  if (session.readonly) return false;
  // 功能：旧模型宏尚未迁移时不能修改新核心的工作区标识或误报执行成功；时间：2026-11-09；作者：lq。
  if (session.wordCore) return false;
  return session.kind === macro.kind;
}

export interface MacroRunResult {
  ok: boolean;
  error?: string;
}

/**
 * 执行一条插件宏。
 *
 * 失败时不改动文档：`runMacro` 会先整体校验再执行，抛错时原数据未被触碰。
 * 失败原因同时写进注册表，插件卡片上能看到"最近一次为什么失败"——
 * 只在 toast 里闪一下，用户切走就再也找不到了。
 */
export function executePluginMacro(macro: PluginMacroCommand, sessionId?: string): MacroRunResult {
  const state = useWorkspaceStore.getState();
  const id = sessionId ?? state.activeId;
  const session = id ? state.sessions[id] : undefined;

  if (!session) return { ok: false, error: '没有打开的文档' };
  if (!canRunMacro(session, macro)) {
    return {
      ok: false,
      error: session.readonly ? '当前文档为只读，无法执行命令' : '该命令不适用于当前文档类型',
    };
  }

  try {
    // 宏的操作名与文档类型在安装期已校验过，这里的断言只用于收窄联合类型。
    const next = runMacro(macro.kind as 'docx' | 'xlsx', session.data as never, macro.command.steps);
    state.updateData(session.id, next as never, { dirty: true });
    state.notify('success', `已执行「${macro.command.title}」`, `来自插件 ${macro.pluginName}`);
    usePluginRegistry.getState().reportPluginError(macro.pluginId, null);
    return { ok: true };
  } catch (error) {
    const message =
      error instanceof MacroError
        ? error.issues.length > 0
          ? `${error.message}：${error.issues[0].path} ${error.issues[0].message}`
          : error.message
        : error instanceof Error
          ? error.message
          : String(error);
    state.notify('error', '命令执行失败', message);
    usePluginRegistry.getState().reportPluginError(macro.pluginId, message);
    return { ok: false, error: message };
  }
}

/**
 * 把键盘事件分发到插件宏。
 *
 * 由 `AppShell` 在内置快捷键处理之后调用，因此内置命令始终优先：
 * 插件即使声明了冲突组合也抢不走 ⌘S 这类关键操作。
 * 返回 true 表示事件已被消费。
 */
export function dispatchPluginShortcut(event: KeyboardEvent, session: DocumentSession | undefined): boolean {
  const macros = usePluginRegistry.getState().macroCommands();
  if (macros.length === 0) return false;

  const binding = matchShortcutFromEvent(event, collectShortcutBindings(macros));
  if (!binding) return false;
  // 当前文档跑不了这条命令时不吞键，让系统按原样处理，避免"按了没反应也不报错"。
  if (!canRunMacro(session, binding.macro)) return false;

  event.preventDefault();
  executePluginMacro(binding.macro);
  return true;
}
