/** 功能：用现有 React 18 挂载完整 Word 内核，给产品页面提供正文宿主；时间：2026-11-08；作者：lq。 */
import { useEffect, useRef } from 'react';
import { WordCoreSurface, type WordCoreSurfaceOptions } from './core-surface';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import '../../word-core/src/styles/document.css';
import '../../word-core/src/styles/fonts.css';

/** 功能：ready 回调直接交付具体会话，工具栏无需通用命令适配层；时间：2026-11-08；作者：lq。 */
export interface WordCoreSurfaceProps extends WordCoreSurfaceOptions {
  bytes?: Uint8Array;
  /** 功能：正式容器复用工作区的同一会话，组件不负责关闭文档；时间：2026-11-09；作者：lq。 */
  session?: WordCoreSession;
  onReady: (surface: WordCoreSurface) => void;
  onError: (error: unknown) => void;
}
/** 功能：支持 StrictMode、文件切换和异步卸载；不替换产品工具栏或主题；时间：2026-11-08；作者：lq。 */
export function WordCoreSurfaceView(props: WordCoreSurfaceProps) {
  const root = useRef<HTMLDivElement>(null);
  const callbacks = useRef(props);
  callbacks.current = props;
  useEffect(() => {
    const abort = new AbortController();
    let mounted: WordCoreSurface | undefined;
    const mount = props.session
      ? WordCoreSurface.attach.bind(WordCoreSurface, root.current!, props.session)
      : WordCoreSurface.mount.bind(WordCoreSurface, root.current!, props.bytes);
    void mount({
      readonly: props.readonly,
      onState: state => callbacks.current.onState?.(state),
      // 功能：异步分页失败交给现有错误入口，不能静默显示错误页数；时间：2026-11-09；作者：lq。
      onLayout: state => callbacks.current.onLayout?.(state),
      onLayoutError: error => (callbacks.current.onLayoutError ?? callbacks.current.onError)(error),
    }, abort.signal).then(surface => {
      if (abort.signal.aborted) { surface.dispose(); return; }
      mounted = surface;
      callbacks.current.onReady(surface);
    }).catch(error => { if (!abort.signal.aborted) callbacks.current.onError(error); });
    return () => { abort.abort(); mounted?.dispose(); };
  }, [props.bytes, props.readonly, props.session]);
  return <div ref={root} className="mo-word-core" />;
}
