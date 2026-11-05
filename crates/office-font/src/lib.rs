use office_core::OfficeError;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{Duration, Instant};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontInfo {
    pub family: String,
    pub style: String,
    pub path: String,
    pub has_chinese: bool,
    /// 功能：同一字体族的其他名称（中英文/不同平台记录），用于缺失字体判断；时间：2026-10-29；作者：lq。
    #[serde(default)]
    pub aliases: Vec<String>,
}

const CJK_FONT_HINTS: &[&str] = &[
    "songti",
    "heiti",
    "kaiti",
    "fangsong",
    "pingfang",
    "hiragino",
    "simsun",
    "simhei",
    "kaiti",
    "yahei",
    "dengxian",
    "source han",
    "noto sans cjk",
    "noto serif cjk",
    "wenquanyi",
    "微软雅黑",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "等线",
    "思源",
];

pub fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    #[cfg(target_os = "macos")]
    {
        dirs.push(PathBuf::from("/System/Library/Fonts"));
        dirs.push(PathBuf::from("/System/Library/Fonts/Supplemental"));
        dirs.push(PathBuf::from("/Library/Fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join("Library/Fonts"));
        }
    }
    #[cfg(target_os = "windows")]
    {
        if let Some(windows) = std::env::var_os("WINDIR") {
            dirs.push(PathBuf::from(windows).join("Fonts"));
        }
        if let Some(local) = dirs::data_local_dir() {
            dirs.push(local.join("Microsoft/Windows/Fonts"));
        }
    }
    #[cfg(target_os = "linux")]
    {
        dirs.push(PathBuf::from("/usr/share/fonts"));
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join(".fonts"));
            dirs.push(home.join(".local/share/fonts"));
        }
    }
    dirs
}

fn has_chinese_name(family: &str) -> bool {
    family.chars().any(|character| {
        matches!(character, '\u{3400}'..='\u{9fff}' | '\u{f900}'..='\u{faff}')
    }) || {
        let lower = family.to_ascii_lowercase();
        CJK_FONT_HINTS
            .iter()
            .any(|hint| lower.contains(&hint.to_ascii_lowercase()))
    }
}

/// 功能：按字体记录的语言选择名称，优先简体中文、其他中文、英文，全部有效名称仍保留为别名；时间：2026-11-03；作者：lq。
fn preferred_font_names<'a>(
    names: impl IntoIterator<Item = ttf_parser::name::Name<'a>>,
    name_ids: [u16; 2],
) -> Option<(String, Vec<String>)> {
    use ttf_parser::{Language, PlatformId};
    let mut values = Vec::new();
    let mut primary: Option<(String, (u8, usize))> = None;
    for name in names {
        let Some(kind) = name_ids.iter().position(|id| *id == name.name_id) else {
            continue;
        };
        let Some(value) = name.to_string() else { continue };
        let value = value.trim();
        if value.is_empty() || value.chars().any(|c| c.is_control() || c == '\u{fffd}') {
            continue;
        }
        // 功能：语言判断使用 name 表的元数据，不用 Songti/Hiragino 等字体名关键字代替语言；时间：2026-11-03；作者：lq。
        let language = match name.language() {
            Language::Chinese_PeoplesRepublicOfChina | Language::Chinese_Singapore => 0,
            Language::Chinese_Taiwan | Language::Chinese_HongKongSAR | Language::Chinese_MacaoSAR => 1,
            _ if name.platform_id == PlatformId::Windows
                && name.language_id < 0x8000
                && name.language_id & 0x03ff == 0x09 => 2,
            _ if value.is_ascii() => 3,
            _ => 4,
        };
        let priority = (language, kind);
        if primary.as_ref().map_or(true, |(_, previous)| priority < *previous) {
            primary = Some((value.to_string(), priority));
        }
        if !values.iter().any(|item| item == value) {
            values.push(value.to_string());
        }
    }
    let (primary, _) = primary?;
    values.retain(|value| value != &primary);
    Some((primary, values))
}

/// 功能：分别读取字体族与字面样式；同语言优先排版族/子族，并保留中英文族别名供文档匹配；时间：2026-11-03；作者：lq。
fn face_names(data: &[u8], index: u32) -> Option<(String, String, Vec<String>)> {
    use ttf_parser::name_id;
    let face = ttf_parser::Face::parse(data, index).ok()?;
    let (family, aliases) = preferred_font_names(
        face.names(),
        [name_id::TYPOGRAPHIC_FAMILY, name_id::FAMILY],
    )?;
    let style = preferred_font_names(
        face.names(),
        [name_id::TYPOGRAPHIC_SUBFAMILY, name_id::SUBFAMILY],
    ).map(|(style, _)| style).unwrap_or_else(|| "Regular".to_string());
    Some((family, style, aliases))
}

fn parse_font_file(path: &Path, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    let data = match fs::read(path) {
        Ok(data) => data,
        Err(_) => return,
    };
    let count = ttf_parser::fonts_in_collection(&data).unwrap_or(1);
    for index in 0..count {
        if let Some((family, style, aliases)) = face_names(&data, index) {
            let key = format!("{}::{}", family.to_ascii_lowercase(), style.to_ascii_lowercase());
            if !seen.insert(key) {
                continue;
            }
            output.push(FontInfo {
                has_chinese: has_chinese_name(&family) || aliases.iter().any(|alias| has_chinese_name(alias)),
                family,
                style,
                path: path.to_string_lossy().to_string(),
                aliases,
            });
        }
    }
}

fn walk(directory: &Path, depth: usize, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    if depth > 2 || output.len() > 4_000 {
        return;
    }
    let entries = match fs::read_dir(directory) {
        Ok(entries) => entries,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            walk(&path, depth + 1, output, seen);
            continue;
        }
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if matches!(extension.as_str(), "ttf" | "otf" | "ttc" | "otc") {
            parse_font_file(&path, output, seen);
        }
    }
}

pub fn list_fonts() -> Vec<FontInfo> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for directory in font_directories() {
        walk(&directory, 0, &mut output, &mut seen);
    }
    output.sort_by(|a, b| {
        b.has_chinese
            .cmp(&a.has_chinese)
            .then_with(|| a.family.to_lowercase().cmp(&b.family.to_lowercase()))
    });
    output
}

/// 功能：字体缓存，带过期时间与手动失效，用户新装字体后无需重启应用；时间：2026-10-29；作者：lq。
///
/// 旧实现用 `OnceLock` 永久缓存：用户在系统里安装字体后，应用仍按旧清单判断缺失字体，
/// 于是“自己安装的字体”一直显示兼容性告警。改为 30 秒 TTL + 显式失效。
const FONT_CACHE_TTL: Duration = Duration::from_secs(30);

static FONT_CACHE: Mutex<Option<(Instant, Vec<FontInfo>)>> = Mutex::new(None);

pub fn cached_fonts() -> Vec<FontInfo> {
    let mut guard = match FONT_CACHE.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };
    if let Some((stamp, fonts)) = guard.as_ref() {
        if stamp.elapsed() < FONT_CACHE_TTL {
            return fonts.clone();
        }
    }
    let fonts = list_fonts();
    *guard = Some((Instant::now(), fonts.clone()));
    fonts
}

/// 功能：清空字体缓存，供“重新扫描字体”与字体目录变化后立即生效；时间：2026-10-29；作者：lq。
pub fn invalidate_fonts() {
    let mut guard = match FONT_CACHE.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };
    *guard = None;
}

/// 功能：强制重新扫描并回填缓存；时间：2026-10-29；作者：lq。
pub fn refresh_fonts() -> Vec<FontInfo> {
    invalidate_fonts();
    cached_fonts()
}

pub fn font_families() -> Vec<String> {
    let mut families = Vec::new();
    let mut seen = HashSet::new();
    for font in list_fonts() {
        let key = font.family.to_ascii_lowercase();
        if seen.insert(key) {
            families.push(font.family);
        }
    }
    families
}

pub fn require_font(family: &str) -> Result<FontInfo, OfficeError> {
    let lower = family.to_ascii_lowercase();
    list_fonts()
        .into_iter()
        .find(|font| font.family.to_ascii_lowercase() == lower)
        .ok_or_else(|| OfficeError::Other(format!("缺少字体：{family}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 功能：构造多语言 name 表记录，验证字体选择与别名保留，不依赖测试机安装字体；时间：2026-11-03；作者：lq。
    fn family_names(entries: &[(u16, u16, &str)]) -> Option<(String, Vec<String>)> {
        let encoded: Vec<Vec<u8>> = entries.iter().map(|(_, _, value)| {
            value.encode_utf16().flat_map(u16::to_be_bytes).collect()
        }).collect();
        let names = entries.iter().zip(&encoded).map(|((name_id, language_id, _), bytes)| {
            ttf_parser::name::Name {
                platform_id: ttf_parser::PlatformId::Windows,
                encoding_id: 1,
                language_id: *language_id,
                name_id: *name_id,
                name: bytes,
            }
        });
        preferred_font_names(names, [ttf_parser::name_id::TYPOGRAPHIC_FAMILY, ttf_parser::name_id::FAMILY])
    }

    /// 功能：简体中文展示名优先，其他语言仍保留为文档匹配别名；时间：2026-11-03；作者：lq。
    #[test]
    fn prefers_simplified_chinese_without_losing_other_names() {
        let (family, aliases) = family_names(&[
            (1, 0x0409, "Songti SC"),
            (1, 0x0404, "宋體-簡"),
            (1, 0x0804, "宋体-简"),
        ]).unwrap();
        assert_eq!(family, "宋体-简");
        assert_eq!(aliases, ["Songti SC", "宋體-簡"]);
    }

    /// 功能：字体有英文名时不随记录顺序显示阿拉伯文；只有外文名时仍保留正常字体；时间：2026-11-03；作者：lq。
    #[test]
    fn prefers_english_but_preserves_foreign_only_fonts() {
        let (family, aliases) = family_names(&[
            (1, 0x0401, "دمشق"), (1, 0x0409, "Damascus"),
        ]).unwrap();
        assert_eq!(family, "Damascus");
        assert_eq!(aliases, ["دمشق"]);
        assert_eq!(family_names(&[(1, 0x0401, "دمشق")]).unwrap().0, "دمشق");
    }

    /// 功能：同语言下优先排版族，忽略异常名称，保留旧字体族名以兼容原文档；时间：2026-11-03；作者：lq。
    #[test]
    fn prefers_typographic_family_and_ignores_invalid_names() {
        let (family, aliases) = family_names(&[
            (1, 0x0804, "坏\0字体"), (1, 0x0804, "坏\u{fffd}字体"),
            (1, 0x0409, "Example Bold"), (16, 0x0409, "Example"),
        ]).unwrap();
        assert_eq!(family, "Example");
        assert_eq!(aliases, ["Example Bold"]);
    }

    #[test]
    fn scans_system_font_directories() {
        let fonts = list_fonts();
        assert!(!fonts.is_empty(), "system should provide at least one font");
        // Chinese coverage depends on the host's installed fonts.
    }

    #[test]
    fn directories_are_platform_specific() {
        assert!(!font_directories().is_empty());
    }
}
