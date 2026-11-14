/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import type { PmNode } from './editor/convert'

/**
 * Word-style progressive open for large documents.
 *
 * Mounting a 60-page document in one setContent keeps the "Opening…" screen up
 * until the whole ProseMirror DOM is built and laid out — the dominant share of
 * cold-open time. Instead, mount enough blocks to overfill the first screens,
 * let the browser paint them, then append the tail in frame-sized chunks while
 * the user already reads page 1 (the page count in the status bar grows as the
 * tail streams in, like Word's background pagination).
 *
 * While the tail streams the document is incomplete, so:
 * - the editor stays editable except at the append boundary (the streaming
 *   tail guard refuses edits in or after the last mounted block),
 * - saves must wait on waitForFullContent() — a mid-stream save would
 *   serialize, and write to disk, a truncated document,
 * - appends bypass undo history and restore the dirty flag, so streaming is
 *   not an edit; history resets once the tail lands unless the user edited
 *   meanwhile (their undo stack must survive).
 */
export interface PhasedContentHost {
  /** full replace (Tiptap setContent) */
  setContent(doc: PmNode): void
  /** append top-level nodes at the document end, outside undo history */
  appendNodes(nodes: PmNode[]): void
  isDestroyed(): boolean
  /** 功能：组合输入期间不批量改变同一正文 DOM，下一帧继续；时间：2026-11-14；作者：lq。 */
  isComposing?(): boolean
  resetHistory(): void
  /** read-only + save-gate flag while the tail streams */
  setLoading(loading: boolean): void
  getDirty(): boolean
  setDirty(dirty: boolean): void
}

/** dispatched on document once a streamed tail has fully landed (layout
 *  measurers skip the tail's chunks and measure the whole document once) */
export const PHASED_CONTENT_SETTLED_EVENT = 'genoffice:phased-content-settled'

/** blocks in the first synchronous mount: overfills the first screens at any zoom */
export const PHASE1_BLOCKS = 64
/** blocks appended per scheduled chunk while the tail streams in */
export const PHASE_CHUNK_BLOCKS = 128
/** documents at or below this many top-level blocks mount in one pass as before */
export const PHASED_MIN_BLOCKS = 192

/** bumping the token cancels the pending tail (its chunks become no-ops) */
// 功能：按编辑器隔离上游加载状态，多标签并行打开/关闭不会取消另一份正文；时间：2026-11-13；作者：lq。
const defaultScope = {}
interface PhasedState { token: number; settled: Promise<void>; cancelPending: (() => void) | null; error?: unknown }
const states = new WeakMap<object, PhasedState>()
function stateOf(scope: object = defaultScope): PhasedState {
  let state = states.get(scope)
  if (!state) { state = { token: 0, settled: Promise.resolve(), cancelPending: null }; states.set(scope, state) }
  return state
}

/** Resolves when the streaming tail has fully landed (immediately when none is pending). */
export function waitForFullContent(scope: object = defaultScope): Promise<void> {
  const state = stateOf(scope)
  return state.settled.then(() => { if (state.error) throw state.error })
}

/** true while a phased open still has tail chunks to append */
export function isPhasedContentPending(scope: object = defaultScope): boolean {
  return stateOf(scope).cancelPending !== null
}

/** 功能：加载失败后明确阻止保存半份正文，检查不引入异步快照时序变化；时间：2026-11-13；作者：lq。 */
export function assertPhasedContentComplete(scope: object = defaultScope): void {
  const state = stateOf(scope)
  if (state.error) throw state.error
  if (state.cancelPending) throw new Error('Word 正文仍在加载')
}

/** Any full content replacement outside the phased path must drop a pending tail. */
export function cancelPhasedContent(scope: object = defaultScope): void {
  const state = stateOf(scope)
  state.error = undefined
  state.token++
  state.cancelPending?.()
  state.cancelPending = null
}

/**
 * Double-rAF: the browser paints the previous mount between the two callbacks.
 * A timer races the frames: a background tab (a hidden shell view keeps
 * document.hidden false but is handed about one frame a second) would
 * otherwise stretch a long open into hours, one chunk per second.
 */
export const PAINTED_FRAME_FALLBACK_MS = 80
const nextPaintedFrame = (cb: () => void): void => {
  let done = false
  let raf2 = 0
  const fire = (): void => {
    if (done) return
    done = true
    clearTimeout(timer)
    cancelAnimationFrame(raf1)
    cancelAnimationFrame(raf2)
    cb()
  }
  const raf1 = requestAnimationFrame(() => {
    raf2 = requestAnimationFrame(fire)
  })
  const timer = setTimeout(fire, PAINTED_FRAME_FALLBACK_MS)
}

export function setContentPhased(
  host: PhasedContentHost,
  pmDoc: PmNode,
  schedule: (cb: () => void) => void = nextPaintedFrame,
  // 功能：默认作用域兼容上游调用，产品会话传入自己的正文 DOM；时间：2026-11-13；作者：lq。
  scope: object = defaultScope,
): void {
  cancelPhasedContent(scope)
  const state = stateOf(scope)
  const content = pmDoc.content ?? []
  const my = ++state.token
  if (content.length <= PHASED_MIN_BLOCKS) {
    host.setContent(pmDoc)
    return
  }
  host.setContent({ ...pmDoc, content: content.slice(0, PHASE1_BLOCKS) })
  let settle!: () => void
  state.settled = new Promise<void>((resolve) => (settle = resolve))
  state.cancelPending = () => {
    host.setLoading(false)
    settle()
  }
  host.setLoading(true)
  let index = PHASE1_BLOCKS
  // a refused chunk remounted the whole document: whatever the user typed
  // before is gone with it, and an undo must not bring the truncated
  // pre-remount document back
  let remounted = false
  const finish = () => {
    if (my !== state.token) return
    // still "pending" through the history reset: it re-creates the plugin
    // views, whose constructors would otherwise measure the whole document
    // once before the settled event asks for the same pass again
    if (!host.isDestroyed() && (remounted || !host.getDirty())) host.resetHistory()
    state.cancelPending = null
    host.setLoading(false)
    settle()
    if (typeof document !== 'undefined')
      document.dispatchEvent(new CustomEvent(PHASED_CONTENT_SETTLED_EVENT, { detail: scope }))
  }
  const appendChunk = () => {
    if (my !== state.token) return
    if (host.isDestroyed()) return finish()
    if (host.isComposing?.()) { schedule(appendChunk); return }
    const chunk = content.slice(index, index + PHASE_CHUNK_BLOCKS)
    index += chunk.length
    const wasDirty = host.getDirty()
    try {
      host.appendNodes(chunk)
    } catch (error) {
      // 功能：加载中已有用户输入时禁止整篇回退覆盖；失败状态阻止保存截断正文；时间：2026-11-13；作者：lq。
      if (wasDirty) { state.error = error; return finish() }
      // a refused chunk falls back to the one-pass mount instead of leaving the file truncated
      remounted = true
      try {
        host.setContent(pmDoc)
      } catch (fallbackError) {
        state.error = fallbackError
      }
      host.setDirty(wasDirty)
      return finish()
    }
    host.setDirty(wasDirty)
    if (index < content.length) schedule(appendChunk)
    else finish()
  }
  schedule(appendChunk)
}
