/** 功能：将同一次来源尾部追加中的相邻 DOM 插入合并为片段，避免 WebKit 为每个段落重复执行 :has 后代失效；时间：2026-11-14；作者：lq。 */
export function batchWordTailDom<T>(root: HTMLElement, expectedBlocks: number, update: () => T): T {
  const fragment = root.ownerDocument.createDocumentFragment();
  const own = Object.getOwnPropertyDescriptor(root, 'insertBefore');
  const insert = root.insertBefore;
  let blocks = 0;
  const flush = () => {
    if (fragment.childNodes.length) insert.call(root, fragment, null);
    blocks = 0;
  };
  // 功能：仅拦截当前编辑器本次同步更新；移动既有节点或非尾部操作先提交待追加片段，保持原有顺序；时间：2026-11-14；作者：lq。
  root.insertBefore = function <N extends Node>(node: N, before: Node | null): N {
    if (before === null && node.parentNode !== root && node.nodeType !== Node.DOCUMENT_FRAGMENT_NODE) {
      fragment.appendChild(node);
      if ((node as Node & { pmViewDesc?: { node?: unknown } }).pmViewDesc?.node) blocks++;
      // 功能：最后一个来源块插入后立即放回渲染树，后续选区同步仍读取完整 DOM；时间：2026-11-14；作者：lq。
      if (blocks >= expectedBlocks) flush();
      return node;
    }
    flush();
    return insert.call(root, node, before) as N;
  };
  try { return update(); } finally {
    // 功能：异常和短批次同样提交已生成节点并恢复方法，不能让 DOM 与已应用事务脱节；时间：2026-11-14；作者：lq。
    try { flush(); } finally {
      if (own) Object.defineProperty(root, 'insertBefore', own);
      else delete (root as unknown as { insertBefore?: unknown }).insertBefore;
    }
  }
}
