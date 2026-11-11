//! 功能：仅调试构建执行本地 WebView 验收脚本并落盘报告，不进入发布构建；时间：2026-11-11；作者：lq。

/// 功能：报告只写入显式验收目录，不开放通用文件写入入口；时间：2026-11-11；作者：lq。
#[tauri::command]
pub fn word_smoke_report(report: String) -> Result<(), String> {
    let directory = std::env::var("MINI_OFFICE_WORD_SMOKE_DIR").map_err(|_| "未启用本机验收")?;
    std::fs::write(std::path::Path::new(&directory).join("native-report.json"), report).map_err(|error| error.to_string())
}

/// 功能：仅用户显式设置本地脚本时在页面就绪后执行，复用实际 Tauri 窗口与 IPC；时间：2026-11-11；作者：lq。
pub fn plugin() -> tauri::plugin::TauriPlugin<tauri::Wry> {
    tauri::plugin::Builder::new("word-smoke")
        .on_page_load(|webview, payload| {
            eprintln!("Word smoke page: {:?} {}", payload.event(), payload.url());
            if payload.event() != tauri::webview::PageLoadEvent::Finished { return; }
            if let Ok(path) = std::env::var("MINI_OFFICE_WORD_SMOKE_SCRIPT") {
                if let Ok(script) = std::fs::read_to_string(path) { eprintln!("Word smoke eval: {:?}", webview.eval(script)); }
            }
        }).build()
}
