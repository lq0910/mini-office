use office_core::{FileInfo, OfficeKind, PlatformInfo, MAX_INPUT_BYTES};
use office_storage::{
    atomic_write, RecentFile, RecentStore, RecoverySnapshot, RecoveryStore, SaveOutcome,
};
use serde::Serialize;
use std::path::Path;
use std::sync::{atomic::{AtomicBool, Ordering}, Mutex};
use tauri::{Emitter, Manager, State};

struct AppState {
    recent: RecentStore,
    recovery: RecoveryStore,
    // 功能：在前端就绪前暂存系统文件事件，并统一保护窗口关闭和退出；时间：2026-10-18；作者：lq。
    pending_files: Mutex<Vec<String>>,
    close_authorized: AtomicBool,
}

/// 功能：托管状态尚未建立时（macOS 冷启动的 odoc 事件早于 setup）先落到全局队列，setup 时并入；时间：2026-10-29；作者：lq。
static EARLY_PENDING_FILES: Mutex<Vec<String>> = Mutex::new(Vec::new());

/// 功能：原子领取待打开文件，避免冷启动时系统事件先于页面监听而丢失；时间：2026-10-18；作者：lq。
#[tauri::command]
fn take_pending_files(state: State<'_, AppState>) -> Result<Vec<String>, String> {
    let mut paths = {
        let mut pending = state.pending_files.lock().map_err(|error| error.to_string())?;
        std::mem::take(&mut *pending)
    };
    if let Ok(mut early) = EARLY_PENDING_FILES.lock() {
        for path in std::mem::take(&mut *early) {
            if !paths.contains(&path) {
                paths.push(path);
            }
        }
    }
    Ok(paths)
}

/// 功能：前端逐份完成保存确认后允许应用退出；时间：2026-10-18；作者：lq。
#[tauri::command]
fn confirm_app_close(app: tauri::AppHandle, state: State<'_, AppState>) {
    state.close_authorized.store(true, Ordering::SeqCst);
    app.exit(0);
}

/// 功能：把待打开文件交给前端，并唤醒主窗口，冷启动与运行中打开共用同一队列；时间：2026-10-29；作者：lq。
fn focus_main_window(app: &tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.unminimize();
        let _ = window.show();
        let _ = window.set_focus();
    }
}

/// 功能：原子入队并通知前端领取，重复路径只保留一次；时间：2026-10-29；作者：lq。
fn queue_paths(app: &tauri::AppHandle, paths: Vec<String>) {
    if paths.is_empty() {
        focus_main_window(app);
        return;
    }
    if let Some(state) = app.try_state::<AppState>() {
        if let Ok(mut pending) = state.pending_files.lock() {
            for path in paths {
                if !pending.contains(&path) {
                    pending.push(path);
                }
            }
        }
    } else if let Ok(mut early) = EARLY_PENDING_FILES.lock() {
        // 功能：冷启动阶段状态未建立，先存全局队列，setup 与 take_pending_files 都会并入；时间：2026-10-29；作者：lq。
        for path in paths {
            if !early.contains(&path) {
                early.push(path);
            }
        }
    }
    focus_main_window(app);
    let _ = app.emit("mo:open-files", ());
}

/// 功能：接收 macOS 文件 URL 并唤醒应用，中文与空格路径由 URL 库解码；时间：2026-10-18；作者：lq。
#[cfg(target_os = "macos")]
fn queue_open_files(app: &tauri::AppHandle, urls: Vec<tauri::Url>) {
    let paths: Vec<String> = urls
        .into_iter()
        .filter_map(|url| url.to_file_path().ok())
        .filter(|path| path.is_file())
        .map(|path| {
            path.canonicalize()
                .unwrap_or(path)
                .to_string_lossy()
                .into_owned()
        })
        .collect();
    queue_paths(app, paths);
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DocumentInspectionDto {
    kind: OfficeKind,
    level: u8,
    main_part: String,
    entry_count: usize,
    warnings: Vec<office_core::CompatibilityWarning>,
    features: Vec<office_core::CapabilityFeature>,
    theme_colors: Vec<(String, String)>,
    major_font: Option<String>,
    minor_font: Option<String>,
    relationship_count: usize,
    compatibility_score: u8,
}

fn allowed_output(path: &str) -> bool {
    let extension = Path::new(path)
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    matches!(
        extension.as_str(),
        "docx" | "xlsx" | "pptx" | "csv" | "pdf" | "zip" | "png" | "jpg" | "jpeg"
    )
}

#[tauri::command]
fn read_file(path: String) -> Result<Vec<u8>, String> {
    let metadata = std::fs::metadata(&path).map_err(|error| format!("无法读取文件：{error}"))?;
    if !metadata.is_file() {
        return Err("目标不是文件。".to_string());
    }
    if metadata.len() > MAX_INPUT_BYTES {
        return Err("文件超过 250 MB 的 V1 上限。".to_string());
    }
    std::fs::read(&path).map_err(|error| format!("无法读取文件：{error}"))
}

#[tauri::command]
fn write_file(path: String, bytes: Vec<u8>) -> Result<SaveOutcome, String> {
    if bytes.is_empty() {
        return Err("拒绝写入空文件。".to_string());
    }
    if !allowed_output(&path) {
        return Err("不允许写入该文件类型。".to_string());
    }
    let bytes = if OfficeKind::from_extension(&path).is_some() && Path::new(&path).is_file() {
        // 功能：保存前把原包中输出遗漏的部件补回，保护复杂对象和厂商扩展；时间：2026-10-28；作者：lq。
        let source = std::fs::read(&path).map_err(|error| format!("无法读取原文件：{error}"))?;
        office_ooxml::preserve_missing_parts(&source, &bytes).map_err(|error| format!("兼容部件合并失败：{error}"))?
    } else {
        bytes
    };
    if OfficeKind::from_extension(&path).is_some() {
        office_ooxml::validate_package(&bytes)
            .map_err(|error| format!("保存前容器校验失败：{error}"))?;
    }
    atomic_write(&path, &bytes, true).map_err(|error| error.to_string())
}

/// 功能：在线程池扫描字体，避免冷启动字体文件读取阻塞原生窗口事件和加载动画；时间：2026-10-24；作者：lq。
#[tauri::command]
async fn list_fonts() -> Result<Vec<office_font::FontInfo>, String> {
    tauri::async_runtime::spawn_blocking(office_font::cached_fonts)
        .await
        .map_err(|error| format!("字体扫描失败：{error}"))
}

/// 功能：强制重新扫描系统字体，用户在应用运行期间安装字体后可立即识别；时间：2026-10-29；作者：lq。
#[tauri::command]
async fn refresh_fonts() -> Result<Vec<office_font::FontInfo>, String> {
    tauri::async_runtime::spawn_blocking(office_font::refresh_fonts)
        .await
        .map_err(|error| format!("字体扫描失败：{error}"))
}

#[tauri::command]
fn inspect_document(path: String) -> Result<DocumentInspectionDto, String> {
    let bytes = read_file(path.clone())?;
    let name = Path::new(&path)
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("document");
    let inspection = office_ooxml::analyze(&bytes, name).map_err(|error| error.to_string())?;
    let compatibility_score = inspection.capability.compatibility_score();
    // 功能：把主题和关系摘要暴露给兼容性报告，便于后续编辑器复用统一的主题继承结果；时间：2026-10-28；作者：lq。
    let theme_part = match inspection.kind {
        OfficeKind::Docx => "word/theme/theme1.xml",
        OfficeKind::Xlsx => "xl/theme/theme1.xml",
        OfficeKind::Pptx => "ppt/theme/theme1.xml",
    };
    let theme = office_ooxml::read_text_entry(&bytes, theme_part)
        .map(|xml| office_ooxml::parse_theme(&xml))
        .unwrap_or_default();
    let relationship_count = office_ooxml::entry_names(&bytes)
        .unwrap_or_default()
        .into_iter()
        .filter(|entry| entry.ends_with(".rels"))
        .filter_map(|entry| office_ooxml::read_text_entry(&bytes, &entry).ok())
        .map(|xml| office_ooxml::relationship_records(&xml).len())
        .sum();
    Ok(DocumentInspectionDto {
        kind: inspection.kind,
        level: inspection.capability.level,
        main_part: inspection.main_part,
        entry_count: inspection.entry_count,
        warnings: inspection.capability.warnings,
        features: inspection.capability.features,
        theme_colors: theme.colors.into_iter().collect(),
        major_font: theme.major_font,
        minor_font: theme.minor_font,
        relationship_count,
        compatibility_score,
    })
}

/// 功能：在内存中把旧版 DOC/XLS/PPT 转换为 OOXML 字节，源目录不产生任何文件；时间：2026-11-02；作者：lq。
#[tauri::command]
fn convert_legacy_document(input: String) -> Result<ConvertedDocumentDto, String> {
    let path = Path::new(&input);
    if !path.is_file() {
        return Err("源文件不存在。".to_string());
    }
    let bytes = std::fs::read(path).map_err(|error| format!("无法读取源文件：{error}"))?;
    let source_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| "无法识别源文件名。".to_string())?
        .to_string();
    let converted = office_compat::convert_legacy_bytes(&bytes, &source_name)
        .map_err(|error| error.to_string())?;
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("未命名文档");
    let extension = office_compat::LegacyKind::from_path(path)
        .map(|kind| kind.target_extension())
        .unwrap_or("docx");
    Ok(ConvertedDocumentDto {
        name: format!("{stem}.{extension}"),
        source_name,
        bytes: converted,
    })
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ConvertedDocumentDto {
    /// 转换后可直接交给编辑器打开的文件名（扩展名为 OOXML）。
    name: String,
    /// 用户原始文件名（旧格式扩展名），用于标签显示与只读兼容模式判定。
    source_name: String,
    bytes: Vec<u8>,
}

/// 功能：转换拖拽到桌面的旧版 Office 字节，避免必须先保存为临时文件；时间：2026-10-28；作者：lq。
#[tauri::command]
fn convert_legacy_bytes(name: String, bytes: Vec<u8>) -> Result<ConvertedDocumentDto, String> {
    let extension = Path::new(&name)
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let output_extension = match extension.as_str() {
        "doc" => "docx",
        "xls" => "xlsx",
        "ppt" => "pptx",
        _ => return Err("只支持 DOC、XLS、PPT 旧版文件转换。".to_string()),
    };
    let converted = office_compat::convert_legacy_bytes(&bytes, &name).map_err(|error| error.to_string())?;
    let stem = Path::new(&name)
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("未命名文档");
    let display_name = format!("{stem}.{output_extension}");
    Ok(ConvertedDocumentDto {
        name: display_name,
        source_name: name,
        bytes: converted,
    })
}

#[tauri::command]
fn platform_info() -> PlatformInfo {
    office_platform::platform_info()
}

/// 功能：统一读取产品名与包标识，供默认应用登记与状态判断使用；时间：2026-10-29；作者：lq。
fn app_identity(app: &tauri::AppHandle) -> (String, String) {
    let config = app.config();
    let name = config
        .product_name
        .clone()
        .unwrap_or_else(|| "Mini Office".to_string());
    (name, config.identifier.clone())
}

/// 功能：查询各 Office 类型是否已由本应用默认打开；时间：2026-10-29；作者：lq。
#[tauri::command]
async fn default_app_status(
    app: tauri::AppHandle,
) -> Result<office_platform::default_app::DefaultAppStatus, String> {
    let (name, identifier) = app_identity(&app);
    tauri::async_runtime::spawn_blocking(move || office_platform::default_app::status(&name, &identifier))
        .await
        .map_err(|error| format!("默认应用状态查询失败：{error}"))
}

/// 功能：申请默认应用，Windows 登记候选并打开系统设置，macOS 直接调用 LaunchServices；时间：2026-10-29；作者：lq。
#[tauri::command]
async fn request_default_app(
    app: tauri::AppHandle,
) -> Result<office_platform::default_app::DefaultAppOutcome, String> {
    let (name, identifier) = app_identity(&app);
    tauri::async_runtime::spawn_blocking(move || {
        office_platform::default_app::request_default(&name, &identifier)
    })
    .await
    .map_err(|error| format!("默认应用申请失败：{error}"))
}

/// 功能：打开系统默认应用设置页，供 Windows 用户在设置中逐项确认；时间：2026-10-29；作者：lq。
#[tauri::command]
async fn open_default_app_settings(app: tauri::AppHandle) -> Result<(), String> {
    let (name, _) = app_identity(&app);
    tauri::async_runtime::spawn_blocking(move || {
        office_platform::default_app::open_system_settings(&name)
    })
    .await
    .map_err(|error| format!("打开系统设置失败：{error}"))?
}

/// 功能：重新登记文件关联，修复“右键 → 打开方式 → Mini Office”只启动软件不打开文档；时间：2026-10-29；作者：lq。
#[tauri::command]
async fn repair_file_associations(app: tauri::AppHandle) -> Result<String, String> {
    let (name, _) = app_identity(&app);
    tauri::async_runtime::spawn_blocking(move || office_platform::default_app::register(&name))
        .await
        .map_err(|error| format!("文件关联修复失败：{error}"))?
}

#[tauri::command]
fn app_version() -> String {
    office_platform::app_version().to_string()
}

#[tauri::command]
fn recent_list(state: State<'_, AppState>, limit: Option<usize>) -> Result<Vec<RecentFile>, String> {
    state
        .recent
        .list(limit.unwrap_or(30))
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn recent_record(state: State<'_, AppState>, path: String) -> Result<RecentFile, String> {
    let info: FileInfo = office_storage::inspect_file(&path).map_err(|error| error.to_string())?;
    let kind = OfficeKind::from_extension(&path)
        .map(|kind| match kind {
            OfficeKind::Docx => "docx",
            OfficeKind::Xlsx => "xlsx",
            OfficeKind::Pptx => "pptx",
        })
        .unwrap_or("bin")
        .to_string();
    let entry = RecentFile {
        id: uuid::Uuid::new_v4().to_string(),
        path: info.path,
        name: info.name,
        kind,
        size: info.size,
        updated_at: chrono::Utc::now().timestamp_millis(),
        level: 3,
    };
    state
        .recent
        .record(&entry)
        .map_err(|error| error.to_string())?;
    Ok(entry)
}

#[tauri::command]
fn recovery_save(
    state: State<'_, AppState>,
    display_name: String,
    kind: String,
    bytes: Vec<u8>,
) -> Result<RecoverySnapshot, String> {
    if bytes.is_empty() || bytes.len() as u64 > MAX_INPUT_BYTES {
        return Err("快照大小无效。".to_string());
    }
    state
        .recovery
        .save(&display_name, &kind, &bytes)
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_list(state: State<'_, AppState>) -> Result<Vec<RecoverySnapshot>, String> {
    state.recovery.list().map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_load(state: State<'_, AppState>, id: String) -> Result<Vec<u8>, String> {
    state.recovery.load(&id).map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_delete(state: State<'_, AppState>, id: String) -> Result<(), String> {
    state
        .recovery
        .delete(&id)
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn print_window(window: tauri::WebviewWindow) -> Result<(), String> {
    window.print().map_err(|error| format!("打印失败：{error}"))
}

#[tauri::command]
fn export_pdf(
    input: String,
    output_directory: String,
) -> Result<office_pdf::PdfExportResult, String> {
    let converter = office_pdf::LibreOfficeConverter::discover().ok_or_else(|| {
        "未找到 LibreOffice。Mini Office 前端会使用内置 PDF 导出；LibreOffice 转换仅作为 PoC 对照。"
            .to_string()
    })?;
    converter
        .convert_to_pdf(&input, &output_directory)
        .map_err(|error| error.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let builder = tauri::Builder::default();
    // 功能：Windows/Linux 拖拽文件到图标或双击关联文件会启动第二个进程，由单实例插件转发参数到已运行窗口；时间：2026-10-29；作者：lq。
    #[cfg(any(target_os = "windows", target_os = "linux"))]
    let builder = builder.plugin(tauri_plugin_single_instance::init(|app, argv, _cwd| {
        let paths = office_platform::default_app::openable_paths(argv.into_iter().skip(1));
        queue_paths(app, paths);
    }));
    builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            office_platform::ensure_directories()?;
            let app_data = office_platform::app_data_dir();
            // 功能：冷启动参数即待打开文件，Windows 拖拽到图标与命令行打开都走这条路径；时间：2026-10-29；作者：lq。
            let mut initial_files = office_platform::default_app::openable_paths(std::env::args().skip(1));
            // 功能：合并 setup 之前收到的系统文件事件（macOS 冷启动 odoc 早于托管状态）；时间：2026-10-29；作者：lq。
            if let Ok(mut early) = EARLY_PENDING_FILES.lock() {
                for path in std::mem::take(&mut *early) {
                    if !initial_files.contains(&path) {
                        initial_files.push(path);
                    }
                }
            }
            let state = AppState {
                recent: RecentStore::open(&app_data)?,
                recovery: RecoveryStore::open(&app_data)?,
                pending_files: Mutex::new(initial_files),
                close_authorized: AtomicBool::new(false),
            };
            app.manage(state);
            // 功能：macOS 每次启动后台刷新一次 LaunchServices 登记，避免覆盖安装后关联指向旧路径；时间：2026-10-29；作者：lq。
            #[cfg(target_os = "macos")]
            {
                std::thread::spawn(|| {
                    let _ = office_platform::default_app::register("Mini Office");
                });
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            read_file,
            write_file,
            list_fonts,
            refresh_fonts,
            inspect_document,
            convert_legacy_document,
            convert_legacy_bytes,
            platform_info,
            app_version,
            default_app_status,
            request_default_app,
            open_default_app_settings,
            repair_file_associations,
            recent_list,
            recent_record,
            recovery_save,
            recovery_list,
            recovery_load,
            recovery_delete,
            export_pdf,
            print_window,
            take_pending_files,
            confirm_app_close
        ])
        .build(tauri::generate_context!())
        .expect("Mini Office failed to start")
        .run(|app, event| {
            // 功能：原生关窗和 Cmd+Q 均等待前端保存确认，不直接结束进程；时间：2026-10-18；作者：lq。
            match event {
                tauri::RunEvent::WindowEvent { event: tauri::WindowEvent::CloseRequested { api, .. }, .. } => {
                    if !app.state::<AppState>().close_authorized.load(Ordering::SeqCst) {
                        api.prevent_close();
                        let _ = app.emit("mo:request-close", ());
                    }
                }
                tauri::RunEvent::ExitRequested { api, .. } => {
                    if !app.state::<AppState>().close_authorized.load(Ordering::SeqCst) {
                        api.prevent_exit();
                        let _ = app.emit("mo:request-close", ());
                    }
                }
                #[cfg(target_os = "macos")]
                tauri::RunEvent::Opened { urls } => queue_open_files(app, urls),
                _ => {}
            }
        });
}
