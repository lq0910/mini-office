//! 功能：注册系统文件关联并申请默认应用，Windows 走注册表与系统设置页，macOS 走 LaunchServices，Linux 走 xdg-mime；时间：2026-10-29；作者：lq。
//!
//! 设计约束：现代操作系统均禁止程序静默改写默认打开方式。Windows 只允许登记候选应用后由用户在
//! 设置页确认，macOS 允许调用 LaunchServices 申请但系统仍可拒绝，因此本模块统一返回“申请结果”
//! 而不是“已成功”，由界面告知用户后续动作。

use serde::Serialize;
use std::path::Path;

/// 关联目标：界面标签、扩展名与各平台标识。
#[derive(Debug, Clone, Copy)]
pub struct AssociationTarget {
    /// 主扩展名，用于界面展示、默认值判断与 ProgID 命名。
    pub extension: &'static str,
    /// 同族扩展名，登记时与主扩展名一并写入，避免 `.markdown` / `.htm` 被漏掉。
    pub extra_extensions: &'static [&'static str],
    pub label: &'static str,
    pub mime: &'static str,
    /// macOS 统一类型标识，系统通过它判断默认打开程序。
    pub uti: &'static str,
    /// Windows 程序标识（ProgID），同时用于注册表登记与默认值判断。
    pub prog_id: &'static str,
    /// 是否参与“设为默认应用”：HTML 只登记到“打开方式”，不抢占系统默认浏览器。
    pub apply_default: bool,
}

impl AssociationTarget {
    /// 功能：主扩展名与同族扩展名的完整登记清单；时间：2026-11-01；作者：lq。
    pub fn extensions(&self) -> impl Iterator<Item = &'static str> {
        std::iter::once(self.extension).chain(self.extra_extensions.iter().copied())
    }
}

pub const TARGETS: [AssociationTarget; 7] = [
    AssociationTarget {
        extension: "docx",
        extra_extensions: &["docm", "dotx"],
        label: "Word 文档",
        mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        uti: "org.openxmlformats.wordprocessingml.document",
        prog_id: "MiniOffice.docx",
        apply_default: true,
    },
    AssociationTarget {
        extension: "xlsx",
        extra_extensions: &["xlsm", "xltx"],
        label: "Excel 工作簿",
        mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        uti: "org.openxmlformats.spreadsheetml.sheet",
        prog_id: "MiniOffice.xlsx",
        apply_default: true,
    },
    AssociationTarget {
        extension: "pptx",
        extra_extensions: &["pptm", "potx"],
        label: "PowerPoint 演示文稿",
        mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        uti: "org.openxmlformats.presentationml.presentation",
        prog_id: "MiniOffice.pptx",
        apply_default: true,
    },
    AssociationTarget {
        extension: "csv",
        extra_extensions: &[],
        label: "CSV 文本",
        mime: "text/csv",
        uti: "public.comma-separated-values-text",
        prog_id: "MiniOffice.csv",
        apply_default: true,
    },
    // 功能：登记 PDF 关联，右键“打开方式”可直接用本应用查看；时间：2026-11-01；作者：lq。
    AssociationTarget {
        extension: "pdf",
        extra_extensions: &[],
        label: "PDF 文档",
        mime: "application/pdf",
        uti: "com.adobe.pdf",
        prog_id: "MiniOffice.pdf",
        apply_default: true,
    },
    // 功能：登记 Markdown 关联，覆盖 .md 与 .markdown 两种写法；时间：2026-11-01；作者：lq。
    AssociationTarget {
        extension: "md",
        extra_extensions: &["markdown"],
        label: "Markdown 文本",
        mime: "text/markdown",
        uti: "net.daringfireball.markdown",
        prog_id: "MiniOffice.md",
        apply_default: true,
    },
    // 功能：登记 HTML 关联，覆盖 .html 与 .htm 两种写法；时间：2026-11-01；作者：lq。
    AssociationTarget {
        extension: "html",
        extra_extensions: &["htm"],
        label: "HTML 网页",
        mime: "text/html",
        uti: "public.html",
        prog_id: "MiniOffice.html",
        // 功能：HTML 不参与默认应用申请，避免抢占系统默认浏览器；时间：2026-11-01；作者：lq。
        apply_default: false,
    },
];

/// 功能：可直接通过系统关联打开的全部扩展名，冷启动参数与拖拽校验共用；时间：2026-10-29；作者：lq。
pub const OPENABLE_EXTENSIONS: [&str; 18] = [
    "docx", "docm", "dotx", "xlsx", "xlsm", "xltx", "pptx", "pptm", "potx", "csv", "doc", "xls", "ppt",
    // 功能：PDF / Markdown / HTML 也走同一套“打开方式”与命令行打开路径；时间：2026-11-01；作者：lq。
    "pdf", "md", "markdown", "html", "htm",
];

/// 功能：把已登记类型拼成面向用户的扩展名清单，避免提示文案与 TARGETS 脱节；时间：2026-11-01；作者：lq。
pub fn supported_extension_list() -> String {
    TARGETS
        .iter()
        .map(|target| target.extension.to_ascii_uppercase())
        .collect::<Vec<_>>()
        .join(" / ")
}

/// 功能：仅列出参与“设为默认应用”的类型，与 status 面板条目保持一致；时间：2026-11-01；作者：lq。
pub fn default_extension_list() -> String {
    TARGETS
        .iter()
        .filter(|target| target.apply_default)
        .map(|target| target.extension.to_ascii_uppercase())
        .collect::<Vec<_>>()
        .join(" / ")
}

/// 功能：参与默认应用申请的类型清单，status 与申请流程共用；时间：2026-11-01；作者：lq。
pub fn default_targets() -> impl Iterator<Item = &'static AssociationTarget> {
    TARGETS.iter().filter(|target| target.apply_default)
}

pub fn is_openable_extension(extension: &str) -> bool {
    let normalized = extension.trim_start_matches('.').to_ascii_lowercase();
    OPENABLE_EXTENSIONS.contains(&normalized.as_str())
}

/// 功能：仅接受真实存在的受支持文件，过滤 `-psn_` 等系统参数与目录；时间：2026-10-29；作者：lq。
pub fn is_openable_path(candidate: &str) -> bool {
    let path = Path::new(candidate);
    if !path.is_file() {
        return false;
    }
    path.extension()
        .and_then(|value| value.to_str())
        .is_some_and(is_openable_extension)
}

/// 功能：去掉 Windows 规范化路径的 `\\?\` 前缀，避免最近文件与标题栏出现无法识别的路径；时间：2026-10-29；作者：lq。
pub fn normalize_display_path(path: &str) -> String {
    #[cfg(target_os = "windows")]
    {
        if let Some(rest) = path.strip_prefix(r"\\?\UNC\") {
            return format!(r"\\{rest}");
        }
        if let Some(rest) = path.strip_prefix(r"\\?\") {
            return rest.to_string();
        }
    }
    path.to_string()
}

/// 功能：从进程参数中挑选待打开文件，兼容 Windows 拖拽到图标与 Linux 文件管理器打开；时间：2026-10-29；作者：lq。
pub fn openable_paths<I, S>(args: I) -> Vec<String>
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    let mut paths = Vec::new();
    for arg in args {
        let candidate = arg.into();
        if !is_openable_path(&candidate) {
            continue;
        }
        let normalized = std::fs::canonicalize(&candidate)
            .map(|path| normalize_display_path(&path.to_string_lossy()))
            .unwrap_or_else(|_| normalize_display_path(&candidate));
        if !paths.contains(&normalized) {
            paths.push(normalized);
        }
    }
    paths
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct AssociationStatus {
    pub extension: String,
    pub label: String,
    pub is_default: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefaultAppStatus {
    pub platform: String,
    pub app_name: String,
    /// 能否由程序直接申请，Windows 恒为 false。
    pub can_set_programmatically: bool,
    pub entries: Vec<AssociationStatus>,
    pub all_default: bool,
    pub default_count: usize,
    /// 面向用户的操作说明。
    pub detail: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefaultAppOutcome {
    /// applied（已生效）| settings（已打开系统设置）| manual（需手动确认）| unsupported | error
    pub outcome: String,
    pub message: String,
    pub status: DefaultAppStatus,
}

fn platform_id() -> &'static str {
    if cfg!(target_os = "windows") {
        "windows"
    } else if cfg!(target_os = "macos") {
        "macos"
    } else if cfg!(target_os = "linux") {
        "linux"
    } else {
        "unknown"
    }
}

fn can_set_programmatically() -> bool {
    cfg!(any(target_os = "macos", target_os = "linux"))
}

fn platform_detail(app_name: &str) -> String {
    if cfg!(target_os = "windows") {
        format!(
            "Windows 不允许程序静默修改默认应用。点击后会打开系统设置，请在列表中选择 {app_name} 并逐项确认。"
        )
    } else if cfg!(target_os = "macos") {
        "macOS 由系统统一管理默认应用，点击后直接向系统申请，个别类型可能仍需在“显示简介”中确认。"
            .to_string()
    } else if cfg!(target_os = "linux") {
        "Linux 通过 xdg-mime 写入用户级默认应用，仅影响当前用户。".to_string()
    } else {
        "当前平台不支持设置默认应用。".to_string()
    }
}

/// 功能：当前系统上本应用是否为各类型的默认打开程序；时间：2026-10-29；作者：lq。
pub fn status(app_name: &str, bundle_id: &str) -> DefaultAppStatus {
    // 功能：面板只展示参与默认应用申请的类型，HTML 仅登记“打开方式”不在此列；时间：2026-11-01；作者：lq。
    let entries: Vec<AssociationStatus> = default_targets()
        .map(|target| AssociationStatus {
            extension: target.extension.to_string(),
            label: target.label.to_string(),
            is_default: is_default_handler(target, bundle_id),
        })
        .collect();
    let default_count = entries.iter().filter(|entry| entry.is_default).count();
    DefaultAppStatus {
        platform: platform_id().to_string(),
        app_name: app_name.to_string(),
        can_set_programmatically: can_set_programmatically(),
        all_default: default_count == entries.len(),
        default_count,
        entries,
        detail: platform_detail(app_name),
    }
}

/// 功能：重新登记本应用到系统文件关联数据库，修复覆盖安装/移动应用后“打开方式”失效；时间：2026-10-29；作者：lq。
///
/// 说明：这一步只登记“本应用支持这些类型”，不会改变系统默认应用，因此可以安全地在启动时或用户点击时执行。
pub fn register(app_name: &str) -> Result<String, String> {
    #[cfg(target_os = "windows")]
    {
        windows::register(app_name)?;
        Ok(format!(
            "已在当前用户范围登记 {} 关联，右键“打开方式”会列出 Mini Office。",
            supported_extension_list()
        ))
    }
    #[cfg(target_os = "macos")]
    {
        let _ = app_name;
        macos::register_bundle()?;
        Ok("已向 LaunchServices 重新登记应用，右键“打开方式”列表会立即刷新。".to_string())
    }
    #[cfg(target_os = "linux")]
    {
        let _ = app_name;
        Err("Linux 的文件关联由安装包写入的 .desktop 文件决定，请重新安装以刷新关联。".to_string())
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    {
        let _ = app_name;
        Err("当前平台不支持登记文件关联。".to_string())
    }
}

/// 功能：申请默认应用，Windows 先登记候选再打开系统设置，其他平台直接调用系统接口；时间：2026-10-29；作者：lq。
pub fn request_default(app_name: &str, bundle_id: &str) -> DefaultAppOutcome {
    #[cfg(target_os = "windows")]
    {
        let mut notes: Vec<String> = Vec::new();
        if let Err(error) = windows::register(app_name) {
            notes.push(format!("登记候选应用未完全成功：{error}"));
        }
        match windows::open_default_apps_settings() {
            Ok(()) => DefaultAppOutcome {
                outcome: "settings".to_string(),
                message: if notes.is_empty() {
                    format!(
                        "已打开 Windows 默认应用设置，请为 {} 选择 Mini Office。",
                        default_extension_list()
                    )
                } else {
                    format!("已打开 Windows 默认应用设置；{}", notes.join("；"))
                },
                status: status(app_name, bundle_id),
            },
            Err(error) => DefaultAppOutcome {
                outcome: "error".to_string(),
                message: format!("无法打开系统设置：{error}"),
                status: status(app_name, bundle_id),
            },
        }
    }
    #[cfg(not(target_os = "windows"))]
    {
        let mut failures: Vec<String> = Vec::new();
        // 功能：申请默认前先刷新 LaunchServices 登记，避免系统仍指向旧路径导致申请无效；时间：2026-10-29；作者：lq。
        #[cfg(target_os = "macos")]
        {
            if let Err(error) = macos::register_bundle() {
                failures.push(format!("刷新应用登记：{error}"));
            }
        }
        // 功能：只申请参与默认应用的类型，HTML 仅登记“打开方式”；时间：2026-11-01；作者：lq。
        for target in default_targets() {
            if let Err(error) = set_default_handler(target, app_name, bundle_id) {
                failures.push(format!("{}：{error}", target.extension));
            }
        }
        let refreshed = status(app_name, bundle_id);
        if refreshed.all_default {
            DefaultAppOutcome {
                outcome: "applied".to_string(),
                message: format!(
                    "已将 {app_name} 设为 {} 的默认打开方式。",
                    default_extension_list()
                ),
                status: refreshed,
            }
        } else if failures.is_empty() {
            DefaultAppOutcome {
                outcome: "manual".to_string(),
                message: "系统未接受全部默认应用申请，请在“显示简介 → 打开方式 → 全部更改”中确认。"
                    .to_string(),
                status: refreshed,
            }
        } else {
            DefaultAppOutcome {
                outcome: "manual".to_string(),
                message: format!(
                    "部分类型未能自动设置（{}），请在系统设置或“显示简介”中手动确认。",
                    failures.join("、")
                ),
                status: refreshed,
            }
        }
    }
}

/// 功能：打开系统默认应用设置页，macOS 无对应设置面板时明确返回不支持；时间：2026-10-29；作者：lq。
pub fn open_system_settings(_app_name: &str) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        windows::open_default_apps_settings()
    }
    #[cfg(target_os = "macos")]
    {
        Err("macOS 通过系统接口直接申请默认应用，无需打开设置页。".to_string())
    }
    #[cfg(target_os = "linux")]
    {
        Err("Linux 请使用“设为默认应用”写入 xdg-mime 配置。".to_string())
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
    {
        Err("当前平台不支持默认应用设置。".to_string())
    }
}

/// 功能：解析 `reg query` 输出，取出指定值名对应的数据，供 Windows 默认应用判断复用；时间：2026-10-29；作者：lq。
#[cfg_attr(not(target_os = "windows"), allow(dead_code))]
pub fn parse_reg_query(stdout: &str, value_name: &str) -> Option<String> {
    for line in stdout.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        let mut parts = trimmed.split_whitespace();
        let name = parts.next()?;
        if !value_name.is_empty() && !name.eq_ignore_ascii_case(value_name) {
            continue;
        }
        // 第二段是类型（REG_SZ / REG_NONE 等），其后全部是数据。
        let kind = parts.next().unwrap_or_default();
        if !kind.to_ascii_uppercase().starts_with("REG_") {
            continue;
        }
        let data = parts.collect::<Vec<_>>().join(" ");
        return Some(data);
    }
    None
}

/// 功能：判断默认值是否指向本应用，兼容安装版 ProgID 与开发版可执行文件路径；时间：2026-10-29；作者：lq。
#[cfg_attr(not(target_os = "windows"), allow(dead_code))]
pub fn prog_id_belongs_to_app(prog_id: &str) -> bool {
    prog_id.to_ascii_lowercase().contains("minioffice")
}

// ---------------------------------------------------------------------------
// macOS：LaunchServices
// ---------------------------------------------------------------------------

#[cfg(target_os = "macos")]
mod macos {
    use std::ffi::{c_char, c_void, CString};
    use std::ptr;

    type CFTypeRef = *const c_void;

    #[link(name = "CoreFoundation", kind = "framework")]
    extern "C" {
        fn CFStringCreateWithCString(
            allocator: CFTypeRef,
            c_str: *const c_char,
            encoding: u32,
        ) -> CFTypeRef;
        fn CFStringGetCString(
            string: CFTypeRef,
            buffer: *mut c_char,
            buffer_size: isize,
            encoding: u32,
        ) -> u8;
        fn CFStringGetLength(string: CFTypeRef) -> isize;
        fn CFRelease(value: CFTypeRef);
    }

    #[link(name = "CoreServices", kind = "framework")]
    extern "C" {
        fn LSCopyDefaultRoleHandlerForContentType(content_type: CFTypeRef, role: u32) -> CFTypeRef;
        fn LSSetDefaultRoleHandlerForContentType(
            content_type: CFTypeRef,
            role: u32,
            handler: CFTypeRef,
        ) -> i32;
    }

    /// kCFStringEncodingUTF8
    const UTF8: u32 = 0x0800_0100;
    /// kLSRolesAll
    const ROLES_ALL: u32 = 0xFFFF_FFFF;

    fn cf_string(value: &str) -> Option<CFTypeRef> {
        let raw = CString::new(value).ok()?;
        let handle = unsafe { CFStringCreateWithCString(ptr::null(), raw.as_ptr(), UTF8) };
        if handle.is_null() {
            None
        } else {
            Some(handle)
        }
    }

    fn cf_string_to_string(value: CFTypeRef) -> Option<String> {
        if value.is_null() {
            return None;
        }
        let length = unsafe { CFStringGetLength(value) };
        let capacity = length.saturating_mul(4).saturating_add(1);
        let mut buffer = vec![0 as c_char; capacity as usize];
        let copied = unsafe { CFStringGetCString(value, buffer.as_mut_ptr(), capacity, UTF8) };
        if copied == 0 {
            return None;
        }
        let bytes: Vec<u8> = buffer
            .iter()
            .take_while(|byte| **byte != 0)
            .map(|byte| *byte as u8)
            .collect();
        String::from_utf8(bytes).ok()
    }

    pub fn default_handler(uti: &str) -> Option<String> {
        let content_type = cf_string(uti)?;
        let handler = unsafe { LSCopyDefaultRoleHandlerForContentType(content_type, ROLES_ALL) };
        unsafe { CFRelease(content_type) };
        if handler.is_null() {
            return None;
        }
        let value = cf_string_to_string(handler);
        unsafe { CFRelease(handler) };
        value
    }

    pub fn set_default_handler(uti: &str, bundle_id: &str) -> Result<(), String> {
        let content_type = cf_string(uti).ok_or_else(|| "无法创建内容类型标识".to_string())?;
        let bundle = cf_string(bundle_id).ok_or_else(|| "无法创建应用标识".to_string())?;
        let status = unsafe { LSSetDefaultRoleHandlerForContentType(content_type, ROLES_ALL, bundle) };
        unsafe {
            CFRelease(content_type);
            CFRelease(bundle);
        }
        if status == 0 {
            Ok(())
        } else {
            Err(format!("系统返回状态码 {status}"))
        }
    }

    /// 功能：定位 .app 包根目录，当前可执行文件位于 <app>/Contents/MacOS/ 下；时间：2026-10-29；作者：lq。
    fn bundle_root() -> Option<std::path::PathBuf> {
        let exe = std::env::current_exe().ok()?;
        // MacOS -> Contents -> <App>.app
        exe.ancestors()
            .find(|path| path.extension().is_some_and(|ext| ext.eq_ignore_ascii_case("app")))
            .map(|path| path.to_path_buf())
    }

    /// 功能：调用 lsregister 重新登记本应用，修复“右键打开方式选了我们却打不开文档”；时间：2026-10-29；作者：lq。
    ///
    /// 根因：应用被移动、覆盖安装或从 DMG 直接运行时，LaunchServices 里仍保留旧路径记录，
    /// 系统会把文档交给旧路径或不把文档交给新进程，表现就是“只启动软件、不打开文档”。
    pub fn register_bundle() -> Result<(), String> {
        let bundle = bundle_root().ok_or_else(|| "无法定位应用包路径（可能以裸可执行文件运行）。".to_string())?;
        // App Translocation：从 DMG 直接运行时路径是随机只读目录，登记不会生效，需提示用户拖入应用程序目录。
        if bundle.to_string_lossy().contains("/AppTranslocation/") {
            return Err("应用正从只读镜像位置运行，请先把 Mini Office 拖入“应用程序”文件夹再修复关联。".to_string());
        }
        let tool = "/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister";
        if !std::path::Path::new(tool).exists() {
            return Err("系统缺少 lsregister 工具，无法自动修复。".to_string());
        }
        let output = std::process::Command::new(tool)
            .arg("-f")
            .arg(&bundle)
            .output()
            .map_err(|error| format!("无法调用 lsregister：{error}"))?;
        if output.status.success() {
            Ok(())
        } else {
            Err(String::from_utf8_lossy(&output.stderr).trim().to_string())
        }
    }
}

#[cfg(target_os = "macos")]
fn is_default_handler(target: &AssociationTarget, bundle_id: &str) -> bool {
    macos::default_handler(target.uti)
        .map(|handler| handler.eq_ignore_ascii_case(bundle_id))
        .unwrap_or(false)
}

#[cfg(target_os = "macos")]
fn set_default_handler(target: &AssociationTarget, _app_name: &str, bundle_id: &str) -> Result<(), String> {
    macos::set_default_handler(target.uti, bundle_id)
}

// ---------------------------------------------------------------------------
// Windows：注册表登记 + 系统设置页
// ---------------------------------------------------------------------------

#[cfg(target_os = "windows")]
mod windows {
    use super::{parse_reg_query, prog_id_belongs_to_app, AssociationTarget, TARGETS};
    use std::os::windows::process::CommandExt;
    use std::path::PathBuf;
    use std::process::Command;

    /// CREATE_NO_WINDOW，避免调用 reg.exe 时闪出控制台窗口。
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;
    const CAPABILITY_KEY: &str = "MiniOffice";
    /// Win11 设置页支持按已登记应用直达，名称需与 RegisteredApplications 的值名一致。
    const REGISTERED_APP_NAME: &str = "MiniOffice";

    fn run(program: &str, args: &[String]) -> Result<String, String> {
        let output = Command::new(program)
            .args(args)
            .creation_flags(CREATE_NO_WINDOW)
            .output()
            .map_err(|error| format!("无法执行 {program}：{error}"))?;
        if !output.status.success() {
            let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
            return Err(if detail.is_empty() {
                format!("{program} 退出码 {:?}", output.status.code())
            } else {
                detail
            });
        }
        Ok(String::from_utf8_lossy(&output.stdout).into_owned())
    }

    fn query_value(key: &str, value: Option<&str>) -> Option<String> {
        let mut args = vec!["query".to_string(), key.to_string()];
        match value {
            Some(name) => {
                args.push("/v".to_string());
                args.push(name.to_string());
            }
            None => args.push("/ve".to_string()),
        }
        let stdout = run("reg", &args).ok()?;
        parse_reg_query(&stdout, value.unwrap_or(""))
    }

    fn write_value(key: &str, value: Option<&str>, data: &str, kind: Option<&str>) -> Result<(), String> {
        let mut args = vec!["add".to_string(), key.to_string()];
        match value {
            Some(name) => {
                args.push("/v".to_string());
                args.push(name.to_string());
            }
            None => args.push("/ve".to_string()),
        }
        args.push("/t".to_string());
        args.push(kind.unwrap_or("REG_SZ").to_string());
        args.push("/d".to_string());
        args.push(data.to_string());
        args.push("/f".to_string());
        run("reg", &args).map(|_| ())
    }

    fn executable() -> Result<PathBuf, String> {
        std::env::current_exe().map_err(|error| format!("无法定位应用可执行文件：{error}"))
    }

    /// 功能：在当前用户范围登记候选应用，使系统设置与“打开方式”列表出现 Mini Office；时间：2026-10-29；作者：lq。
    pub fn register(app_name: &str) -> Result<(), String> {
        let exe = executable()?;
        let exe = exe.display().to_string();
        let command = format!("\"{exe}\" \"%1\"");
        let icon = format!("\"{exe}\",0");
        let mut failures: Vec<String> = Vec::new();

        for target in TARGETS {
            let base = format!("HKCU\\Software\\Classes\\{}", target.prog_id);
            let description = format!("{}（{}）", target.label, app_name);
            for (key, data) in [
                (base.clone(), description.clone()),
                (format!("{base}\\DefaultIcon"), icon.clone()),
                (format!("{base}\\shell\\open\\command"), command.clone()),
            ] {
                if let Err(error) = write_value(&key, None, &data, None) {
                    failures.push(error);
                }
            }
            // 登记到 OpenWithProgids 与能力表，右键“打开方式”立即出现本应用，系统设置页也会列出该类型。
            // 同族扩展名（.docm / .markdown / .htm 等）共用同一个 ProgID，一并写入。
            let capabilities = format!("HKCU\\Software\\{CAPABILITY_KEY}\\Capabilities\\FileAssociations");
            for extension in target.extensions() {
                // 功能：仅参与默认应用的类型写入能力表，HTML 只登记“打开方式”；时间：2026-11-01；作者：lq。
                if target.apply_default {
                    let value_name = format!(".{extension}");
                    if let Err(error) = write_value(&capabilities, Some(&value_name), target.prog_id, None) {
                        failures.push(error);
                    }
                }
                let open_with = format!("HKCU\\Software\\Classes\\.{extension}\\OpenWithProgids");
                if let Err(error) = write_value(&open_with, Some(target.prog_id), "", Some("REG_NONE")) {
                    failures.push(error);
                }
            }
        }

        let capability_root = format!("HKCU\\Software\\{CAPABILITY_KEY}\\Capabilities");
        for (name, data) in [
            ("ApplicationName", app_name.to_string()),
            (
                "ApplicationDescription",
                "轻量、离线优先的跨平台办公套件".to_string(),
            ),
            ("ApplicationIcon", icon.clone()),
        ] {
            if let Err(error) = write_value(&capability_root, Some(name), &data, None) {
                failures.push(error);
            }
        }
        if let Err(error) = write_value(
            "HKCU\\Software\\RegisteredApplications",
            Some(REGISTERED_APP_NAME),
            &format!("Software\\{CAPABILITY_KEY}\\Capabilities"),
            None,
        ) {
            failures.push(error);
        }

        if failures.is_empty() {
            Ok(())
        } else {
            Err(failures.join("；"))
        }
    }

    pub fn open_default_apps_settings() -> Result<(), String> {
        // Win11 22H2+ 支持按已登记应用直达；旧版本会忽略参数并停留在默认应用列表。
        let uri = format!("ms-settings:defaultapps?registeredAppUser={REGISTERED_APP_NAME}");
        Command::new("explorer")
            .arg(&uri)
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("无法打开系统设置：{error}"))
    }

    pub fn is_default(target: &AssociationTarget) -> bool {
        let key = format!(
            "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.{}\\UserChoice",
            target.extension
        );
        let Some(prog_id) = query_value(&key, Some("ProgId")) else {
            return false;
        };
        if prog_id_belongs_to_app(&prog_id) {
            return true;
        }
        // 安装版可能使用打包器生成的 ProgID，回退比较打开命令是否指向本应用。
        let command_key = format!("HKCU\\Software\\Classes\\{prog_id}\\shell\\open\\command");
        let Some(command) = query_value(&command_key, None) else {
            return false;
        };
        executable()
            .map(|exe| {
                let name = exe
                    .file_name()
                    .map(|value| value.to_string_lossy().to_ascii_lowercase())
                    .unwrap_or_default();
                !name.is_empty() && command.to_ascii_lowercase().contains(&name)
            })
            .unwrap_or(false)
    }
}

#[cfg(target_os = "windows")]
fn is_default_handler(target: &AssociationTarget, _bundle_id: &str) -> bool {
    windows::is_default(target)
}

#[cfg(target_os = "windows")]
#[allow(dead_code)]
fn set_default_handler(_target: &AssociationTarget, _app_name: &str, _bundle_id: &str) -> Result<(), String> {
    Err("Windows 需要用户在系统设置中确认".to_string())
}

// ---------------------------------------------------------------------------
// Linux：xdg-mime
// ---------------------------------------------------------------------------

#[cfg(target_os = "linux")]
mod linux {
    use super::AssociationTarget;
    use std::process::Command;

    fn run(program: &str, args: &[&str]) -> Option<String> {
        let output = Command::new(program).args(args).output().ok()?;
        if !output.status.success() {
            return None;
        }
        Some(String::from_utf8_lossy(&output.stdout).trim().to_string())
    }

    pub fn desktop_file(app_name: &str) -> Option<String> {
        let candidates = [
            format!("{}.desktop", app_name),
            format!("{}.desktop", app_name.replace(' ', "-").to_ascii_lowercase()),
            "mini-office.desktop".to_string(),
        ];
        let home = dirs::data_dir().unwrap_or_default().join("applications");
        for candidate in candidates {
            if home.join(&candidate).is_file() || std::path::Path::new("/usr/share/applications").join(&candidate).is_file() {
                return Some(candidate);
            }
        }
        None
    }

    pub fn default_handler(mime: &str) -> Option<String> {
        run("xdg-mime", &["query", "default", mime]).filter(|value| !value.is_empty())
    }

    pub fn is_default(target: &AssociationTarget) -> bool {
        default_handler(target.mime)
            .map(|value| value.to_ascii_lowercase().contains("mini"))
            .unwrap_or(false)
    }

    pub fn set_default(target: &AssociationTarget, app_name: &str) -> Result<(), String> {
        let desktop = desktop_file(app_name)
            .ok_or_else(|| "未找到已安装的 .desktop 启动项，请先通过安装包部署应用".to_string())?;
        run("xdg-mime", &["default", &desktop, target.mime])
            .map(|_| ())
            .ok_or_else(|| "xdg-mime 执行失败".to_string())
    }
}

#[cfg(target_os = "linux")]
fn is_default_handler(target: &AssociationTarget, _bundle_id: &str) -> bool {
    linux::is_default(target)
}

#[cfg(target_os = "linux")]
fn set_default_handler(target: &AssociationTarget, app_name: &str, _bundle_id: &str) -> Result<(), String> {
    linux::set_default(target, app_name)
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
fn is_default_handler(_target: &AssociationTarget, _bundle_id: &str) -> bool {
    false
}

#[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
fn set_default_handler(_target: &AssociationTarget, _app_name: &str, _bundle_id: &str) -> Result<(), String> {
    Err("当前平台不支持默认应用设置".to_string())
}

impl DefaultAppStatus {
    /// 功能：供前端在系统状态缺失时构造占位结果，保证界面不出现空白；时间：2026-10-29；作者：lq。
    pub fn unsupported(app_name: &str) -> Self {
        Self {
            platform: platform_id().to_string(),
            app_name: app_name.to_string(),
            can_set_programmatically: false,
            entries: Vec::new(),
            all_default: false,
            default_count: 0,
            detail: "当前平台不支持默认应用设置。".to_string(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_cover_office_types_and_readonly_formats() {
        let extensions: Vec<&str> = TARGETS.iter().map(|target| target.extension).collect();
        assert_eq!(
            extensions,
            vec!["docx", "xlsx", "pptx", "csv", "pdf", "md", "html"]
        );
        for target in TARGETS {
            assert!(target.mime.contains('/'));
            assert!(!target.uti.is_empty());
            assert!(target.prog_id.starts_with("MiniOffice."));
        }
    }

    #[test]
    fn target_extensions_include_family_variants() {
        // 功能：同族扩展名（含宏、模板、.markdown/.htm）必须一并登记，否则右键“打开方式”缺项；时间：2026-11-01；作者：lq。
        let markdown = TARGETS.iter().find(|target| target.extension == "md").expect("缺少 md 关联");
        assert_eq!(markdown.extensions().collect::<Vec<_>>(), vec!["md", "markdown"]);
        let html = TARGETS.iter().find(|target| target.extension == "html").expect("缺少 html 关联");
        assert_eq!(html.extensions().collect::<Vec<_>>(), vec!["html", "htm"]);
        let word = TARGETS.iter().find(|target| target.extension == "docx").expect("缺少 docx 关联");
        assert_eq!(word.extensions().collect::<Vec<_>>(), vec!["docx", "docm", "dotx"]);
        // 每个登记扩展名都必须能被冷启动参数识别，避免“能关联却打不开”。
        for target in TARGETS {
            for extension in target.extensions() {
                assert!(is_openable_extension(extension), "{extension} 未加入可打开清单");
            }
        }
        assert_eq!(supported_extension_list(), "DOCX / XLSX / PPTX / CSV / PDF / MD / HTML");
        // 功能：HTML 只登记“打开方式”，不参与默认应用申请；时间：2026-11-01；作者：lq。
        assert_eq!(default_extension_list(), "DOCX / XLSX / PPTX / CSV / PDF / MD");
        assert!(!default_targets().any(|target| target.extension == "html"));
    }

    #[test]
    fn openable_extension_matches_case_and_dot() {
        assert!(is_openable_extension("DOCX"));
        assert!(is_openable_extension(".xlsx"));
        // 功能：PDF 等只读格式也要能通过“打开方式”进入应用；时间：2026-11-01；作者：lq。
        assert!(is_openable_extension("pdf"));
        assert!(is_openable_extension(".PDF"));
        assert!(is_openable_extension("md"));
        assert!(!is_openable_extension(""));
        assert!(!is_openable_extension("exe"));
    }

    #[test]
    fn openable_paths_ignore_flags_and_missing_files() {
        let missing = std::env::temp_dir().join("mini-office-not-exists.docx");
        let args = vec![
            "-psn_0_12345".to_string(),
            missing.to_string_lossy().into_owned(),
            "--flag".to_string(),
        ];
        assert!(openable_paths(args).is_empty());
    }

    #[test]
    fn openable_paths_keep_existing_document() {
        let path = std::env::temp_dir().join("mini-office-openable-test.docx");
        std::fs::write(&path, b"stub").expect("写入临时文件失败");
        let found = openable_paths(vec![path.to_string_lossy().into_owned()]);
        assert_eq!(found.len(), 1);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn display_path_keeps_plain_paths() {
        assert_eq!(normalize_display_path("/Users/me/合同.docx"), "/Users/me/合同.docx");
    }

    #[test]
    fn reg_query_parses_value_data() {
        let stdout = "\r\nHKEY_CURRENT_USER\\Software\\Classes\\.docx\\OpenWithProgids\r\n    MiniOffice.docx    REG_NONE                        \r\n";
        assert_eq!(parse_reg_query(stdout, "MiniOffice.docx").as_deref(), Some(""));
        let stdout = "HKEY_CURRENT_USER\\...\r\n    ProgId    REG_SZ    Word.Document.12\r\n";
        assert_eq!(parse_reg_query(stdout, "ProgId").as_deref(), Some("Word.Document.12"));
        assert_eq!(parse_reg_query(stdout, "Other"), None);
    }

    #[test]
    fn prog_id_belongs_to_mini_office_only() {
        assert!(prog_id_belongs_to_app("MiniOffice.docx"));
        assert!(prog_id_belongs_to_app("minioffice.desktop"));
        assert!(!prog_id_belongs_to_app("Word.Document.12"));
    }

    #[test]
    fn status_lists_every_target() {
        let report = status("Mini Office", "com.minioffice.desktop");
        assert_eq!(report.entries.len(), default_targets().count());
        assert!(report.entries.iter().all(|entry| entry.extension != "html"));
        assert_eq!(report.default_count, report.entries.iter().filter(|entry| entry.is_default).count());
        assert!(!report.app_name.is_empty());
    }
}
