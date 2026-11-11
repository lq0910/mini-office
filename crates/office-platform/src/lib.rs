use office_core::{OfficeError, PlatformInfo};
use std::fs;
use std::path::PathBuf;

// 功能：文件关联登记与默认应用申请按平台实现，供桌面壳直接调用；时间：2026-10-29；作者：lq。
pub mod default_app;

pub fn app_data_dir() -> PathBuf {
    // 功能：调试验收使用独立恢复/最近文件目录，避免测试触碰用户工作区；时间：2026-11-11；作者：lq。
    #[cfg(debug_assertions)]
    if let Some(directory) = std::env::var_os("MINI_OFFICE_WORD_SMOKE_DATA") { return PathBuf::from(directory); }
    if let Some(directory) = dirs::data_dir() {
        directory.join("MiniOffice")
    } else {
        std::env::temp_dir().join("MiniOffice")
    }
}

pub fn cache_dir() -> PathBuf {
    if let Some(directory) = dirs::cache_dir() {
        directory.join("MiniOffice")
    } else {
        std::env::temp_dir().join("MiniOffice-cache")
    }
}

pub fn recovery_dir() -> PathBuf {
    app_data_dir().join("recovery")
}

pub fn print_dir() -> PathBuf {
    cache_dir().join("print")
}

pub fn ensure_directories() -> Result<(), OfficeError> {
    fs::create_dir_all(app_data_dir())?;
    fs::create_dir_all(cache_dir())?;
    fs::create_dir_all(recovery_dir())?;
    Ok(())
}

pub fn platform_info() -> PlatformInfo {
    PlatformInfo::current(
        app_data_dir().to_string_lossy().to_string(),
        std::env::temp_dir().to_string_lossy().to_string(),
    )
}

pub fn app_version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn info_has_expected_family() {
        let info = platform_info();
        assert!(!info.arch.is_empty());
        assert!(!info.app_data_dir.is_empty());
    }
}
