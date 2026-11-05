use office_core::{OfficeError, MAX_INPUT_BYTES};
use office_oxide::Document;
use std::path::{Path, PathBuf};

/// 功能：表示可转换的旧版 Office 文件类型；时间：2026-10-27；作者：lq。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyKind {
    Doc,
    Xls,
    Ppt,
}

impl LegacyKind {
    pub fn from_path(path: impl AsRef<Path>) -> Option<Self> {
        match path.as_ref().extension()?.to_str()?.to_ascii_lowercase().as_str() {
            "doc" => Some(Self::Doc),
            "xls" => Some(Self::Xls),
            "ppt" => Some(Self::Ppt),
            _ => None,
        }
    }

    pub fn target_extension(self) -> &'static str {
        match self {
            Self::Doc => "docx",
            Self::Xls => "xlsx",
            Self::Ppt => "pptx",
        }
    }
}

/// 功能：为旧版 Office 文件生成默认 OOXML 输出路径；时间：2026-10-27；作者：lq。
pub fn default_output_path(input: impl AsRef<Path>) -> Option<PathBuf> {
    let kind = LegacyKind::from_path(&input)?;
    let path = input.as_ref();
    let stem = path.file_stem()?.to_string_lossy();
    Some(path.with_file_name(format!("{}_兼容转换.{}", stem, kind.target_extension())))
}

/// 功能：使用纯 Rust office_oxide 将 DOC/XLS/PPT 转换为 OOXML，保持主应用不依赖大型办公套件；时间：2026-10-27；作者：lq。
pub fn convert_legacy(input: impl AsRef<Path>, output: impl AsRef<Path>) -> Result<(), OfficeError> {
    let input = input.as_ref();
    let output = output.as_ref();
    if LegacyKind::from_path(input).is_none() {
        return Err(OfficeError::Unsupported("只支持 DOC、XLS、PPT 旧版文件转换。".to_string()));
    }
    let size = std::fs::metadata(input).map_err(OfficeError::Io)?.len();
    if size > MAX_INPUT_BYTES {
        return Err(OfficeError::Security("旧版 Office 文件超过 250 MB 安全上限。".to_string()));
    }
    if !matches!(output.extension().and_then(|value| value.to_str()).map(|value| value.to_ascii_lowercase()).as_deref(), Some("docx" | "xlsx" | "pptx")) {
        return Err(OfficeError::Unsupported("输出文件必须是 DOCX、XLSX 或 PPTX。".to_string()));
    }
    let document = Document::open(input).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
    document.save_as(output).map_err(|error| OfficeError::Storage(error.to_string()))?;
    // DOC 的 SEPX 页尺寸/页边距在 office_oxide 0.1.15 的转换中丢失；
    // 从原件的 PlcfSed 恢复到 OOXML，仅修改内存转换结果，不触碰原文件。
    if LegacyKind::from_path(input) == Some(LegacyKind::Doc) {
        let bytes = std::fs::read(input).map_err(OfficeError::Io)?;
        let converted = std::fs::read(output).map_err(OfficeError::Io)?;
        let restored = restore_doc_sections(&bytes, &converted)?;
        std::fs::write(output, restored).map_err(OfficeError::Io)?;
    }
    Ok(())
}

/// 功能：在内存与临时目录之间完成旧版 Office 转换，供拖拽和非文件路径入口复用；时间：2026-10-28；作者：lq。
pub fn convert_legacy_bytes(bytes: &[u8], name: &str) -> Result<Vec<u8>, OfficeError> {
    let kind = LegacyKind::from_path(name)
        .ok_or_else(|| OfficeError::Unsupported("只支持 DOC、XLS、PPT 旧版文件转换。".to_string()))?;
    if bytes.len() as u64 > MAX_INPUT_BYTES {
        return Err(OfficeError::Security("旧版 Office 文件超过 250 MB 安全上限。".to_string()));
    }
    let directory = tempfile::tempdir().map_err(OfficeError::Io)?;
    // 功能：只使用文件名创建隔离副本，绝对路径参数不能覆盖源文件；时间：2026-11-03；作者：lq。
    let file_name = Path::new(name).file_name()
        .ok_or_else(|| OfficeError::Unsupported("旧格式转换缺少有效文件名。".to_string()))?;
    let input = directory.path().join(file_name);
    let output = directory.path().join(format!("converted.{}", kind.target_extension()));
    std::fs::write(&input, bytes).map_err(OfficeError::Io)?;
    convert_legacy(&input, &output)?;
    std::fs::read(output).map_err(OfficeError::Io)
}

// [MS-DOC] PlcfSed: (n + 1) × CP(u32) + n × Sed(12 bytes),
// Sed.fcSepx 位于第 2 字节；WordDocument[fcSepx] 为 cb(u16) + SEPX SPRM。
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6777a859-d537-486b-96c7-fb0c959a5bf9
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct SectionGeometry {
    width: Option<u16>,
    height: Option<u16>,
    left: Option<u16>,
    right: Option<u16>,
    top: Option<u16>,
    bottom: Option<u16>,
    header: Option<u16>,
    footer: Option<u16>,
    line_pitch: Option<u16>,
}

impl SectionGeometry {
    fn has_page_setup(&self) -> bool {
        self.width.is_some() || self.height.is_some() || self.left.is_some()
            || self.right.is_some() || self.top.is_some() || self.bottom.is_some()
    }

    fn xml(&self) -> String {
        let mut result = String::new();
        if self.width.is_some() || self.height.is_some() {
            let width = self.width.unwrap_or(12240);
            let height = self.height.unwrap_or(15840);
            let orientation = if width > height { " w:orient=\"landscape\"" } else { "" };
            result.push_str(&format!("<w:pgSz w:w=\"{width}\" w:h=\"{height}\"{orientation}/>"));
        }
        if self.left.is_some() || self.right.is_some() || self.top.is_some() || self.bottom.is_some() {
            let (top, right, bottom, left) = (
                self.top.unwrap_or(1440), self.right.unwrap_or(1440),
                self.bottom.unwrap_or(1440), self.left.unwrap_or(1440),
            );
            let header = self.header.unwrap_or(720);
            let footer = self.footer.unwrap_or(720);
            result.push_str(&format!("<w:pgMar w:top=\"{top}\" w:right=\"{right}\" w:bottom=\"{bottom}\" w:left=\"{left}\" w:header=\"{header}\" w:footer=\"{footer}\"/>"));
        }
        if let Some(pitch) = self.line_pitch {
            result.push_str(&format!("<w:docGrid w:type=\"lines\" w:linePitch=\"{pitch}\"/>"));
        }
        result
    }
}

fn le_u16(data: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes(data.get(at..at + 2)?.try_into().ok()?))
}

fn le_u32(data: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes(data.get(at..at + 4)?.try_into().ok()?))
}

fn doc_sections(input: &[u8]) -> Vec<SectionGeometry> {
    let Ok(mut cfb) = office_oxide::cfb::CfbReader::new(std::io::Cursor::new(input)) else { return Vec::new() };
    let Ok(word) = cfb.open_stream("WordDocument") else { return Vec::new() };
    let Some(offset) = le_u32(&word, 0xCA).map(|v| v as usize) else { return Vec::new() };
    let Some(size) = le_u32(&word, 0xCE).map(|v| v as usize) else { return Vec::new() };
    let table_name = if le_u16(&word, 0x0A).unwrap_or(0) & 0x0200 != 0 { "1Table" } else { "0Table" };
    let Ok(table) = cfb.open_stream(table_name).or_else(|_| cfb.open_stream(if table_name == "1Table" { "0Table" } else { "1Table" })) else { return Vec::new() };
    if size < 20 || size > table.len().saturating_sub(offset) || (size - 4) % 16 != 0 { return Vec::new() }
    let count = (size - 4) / 16;
    if count > 4096 { return Vec::new() }
    (0..count).map(|i| {
        let sed = offset + (count + 1) * 4 + i * 12;
        let Some(fc_sepx) = le_u32(&table, sed + 2).map(|v| v as usize) else { return SectionGeometry::default() };
        let Some(cb) = le_u16(&word, fc_sepx).map(|v| v as usize) else { return SectionGeometry::default() };
        let Some(sprms) = word.get(fc_sepx + 2..fc_sepx.saturating_add(2).saturating_add(cb)) else { return SectionGeometry::default() };
        parse_section_sprms(sprms)
    }).collect()
}

fn parse_section_sprms(sprms: &[u8]) -> SectionGeometry {
    let mut result = SectionGeometry::default();
    let mut at = 0;
    while at + 2 <= sprms.len() {
        let opcode = le_u16(sprms, at).unwrap();
        at += 2;
        let size = match opcode >> 13 {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            7 => 3,
            6 => {
                let Some(&size) = sprms.get(at) else { break };
                at += 1;
                size as usize
            }
            _ => break,
        };
        let Some(value) = sprms.get(at..at.saturating_add(size)) else { break };
        if value.len() >= 2 {
            let raw = u16::from_le_bytes([value[0], value[1]]);
            // [MS-DOC] Section Properties: sprmSXaPage 等操作码；负页边距取绝对值。
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/46c3ec54-53ff-4c0a-b0d6-07ad15d2546e
            match opcode {
                0xB01F if (144..=31680).contains(&raw) => result.width = Some(raw),
                0xB020 if (144..=31680).contains(&raw) => result.height = Some(raw),
                0xB021 if raw <= 31680 => result.left = Some(raw),
                0xB022 if raw <= 31680 => result.right = Some(raw),
                0x9023 => result.top = Some((raw as i16).unsigned_abs().min(31665)),
                0x9024 => result.bottom = Some((raw as i16).unsigned_abs().min(31665)),
                0xB017 if raw <= 31680 => result.header = Some(raw),
                0xB018 if raw <= 31680 => result.footer = Some(raw),
                0x9031 if (1..=31680).contains(&raw) => result.line_pitch = Some(raw),
                _ => {}
            }
        }
        at += size;
    }
    result
}

fn patch_document_sections(xml: &str, sections: &[SectionGeometry]) -> String {
    if !sections.iter().any(SectionGeometry::has_page_setup) { return xml.to_string() }
    // 转换器保留原节边界时逐节恢复；若丢失节边界，宁可不捏造中间节，
    // 最后一节的几何信息仍用于页面整体显示。
    let count = xml.match_indices("<w:sectPr").count();
    let mut output = String::with_capacity(xml.len() + sections.len() * 160);
    let mut rest = xml;
    let mut index = 0;
    while let Some(start) = rest.find("<w:sectPr") {
        output.push_str(&rest[..start]);
        let fragment = &rest[start..];
        let Some(open_end) = fragment.find('>') else { return xml.to_string() };
        let self_closing = fragment[..=open_end].ends_with("/>");
        let end = if self_closing { open_end + 1 } else {
            let Some(close_at) = fragment.find("</w:sectPr>") else { return xml.to_string() };
            close_at + "</w:sectPr>".len()
        };
        let current = &fragment[..end];
        let geometry = if count == sections.len() { &sections[index] } else { sections.last().unwrap() };
        let properties = geometry.xml();
        if properties.is_empty() {
            output.push_str(current);
        } else if self_closing {
            output.push_str("<w:sectPr>");
            output.push_str(&properties);
            output.push_str("</w:sectPr>");
        } else {
            // pgSz/pgMar/docGrid 按 CT_SectPr 顺序放在 w:type/页眉引用之后。
            let insert = current.rfind("</w:sectPr>").unwrap();
            output.push_str(&current[..insert]);
            output.push_str(&properties);
            output.push_str(&current[insert..]);
        }
        rest = &fragment[end..];
        index += 1;
    }
    output.push_str(rest);
    if count == 0 {
        if let Some(body_end) = output.rfind("</w:body>") {
            output.insert_str(body_end, &format!("<w:sectPr>{}</w:sectPr>", sections.last().unwrap().xml()));
        }
    }
    output
}

fn restore_doc_sections(input: &[u8], converted: &[u8]) -> Result<Vec<u8>, OfficeError> {
    use std::io::{Cursor, Read, Write};
    use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};
    let sections = doc_sections(input);
    if !sections.iter().any(SectionGeometry::has_page_setup) { return Ok(converted.to_vec()) }
    let mut archive = ZipArchive::new(Cursor::new(converted)).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
    let mut output = ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
        let name = entry.name().to_string();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).map_err(OfficeError::Io)?;
        if name == "word/document.xml" {
            let xml = String::from_utf8(bytes).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
            bytes = patch_document_sections(&xml, &sections).into_bytes();
        }
        if name.ends_with('/') {
            output.add_directory(name, SimpleFileOptions::default()).map_err(|error| OfficeError::Storage(error.to_string()))?;
        } else {
            output.start_file(name, SimpleFileOptions::default().compression_method(CompressionMethod::Deflated))
                .map_err(|error| OfficeError::Storage(error.to_string()))?;
            output.write_all(&bytes).map_err(OfficeError::Io)?;
        }
    }
    Ok(output.finish().map_err(|error| OfficeError::Storage(error.to_string()))?.into_inner())
}

#[cfg(test)]
mod tests {
    use super::{convert_legacy_bytes, default_output_path, LegacyKind};

    #[test]
    fn detects_legacy_extensions_and_output_name() {
        assert_eq!(LegacyKind::from_path("报告.XLS"), Some(LegacyKind::Xls));
        assert_eq!(LegacyKind::from_path("演示.ppt"), Some(LegacyKind::Ppt));
        assert_eq!(default_output_path("合同.doc").unwrap().to_string_lossy(), "合同_兼容转换.docx");
    }

    // 功能：验证内存入口先校验格式并拒绝不支持的扩展名；时间：2026-10-28；作者：lq。
    #[test]
    fn rejects_unsupported_memory_input() {
        assert!(convert_legacy_bytes(b"invalid", "sample.txt").is_err());
    }

    #[test]
    fn preserves_real_doc_page_geometry_and_tables() {
        use std::io::Read;
        let samples = [
            ("/Users/liqiang/Downloads/采购文件.doc", true),
            ("/Users/liqiang/Downloads/测试分包合同.doc", false),
        ];
        for (name, has_tables) in samples {
            if !std::path::Path::new(name).exists() { continue }
            let input = std::fs::read(name).unwrap();
            let geometry = super::doc_sections(&input);
            eprintln!("{name}: {geometry:?}");
            assert!(geometry.iter().any(super::SectionGeometry::has_page_setup));
            let converted = convert_legacy_bytes(&input, name).unwrap();
            let mut zip = zip::ZipArchive::new(std::io::Cursor::new(converted)).unwrap();
            let mut xml = String::new();
            zip.by_name("word/document.xml").unwrap().read_to_string(&mut xml).unwrap();
            assert!(xml.contains("<w:pgSz"), "{name}: page size missing");
            assert!(xml.contains("<w:pgMar"), "{name}: margins missing");
            if has_tables { assert!(xml.contains("<w:tbl>"), "{name}: tables missing"); }
        }
    }

    #[test]
    fn restores_section_geometry_without_erasing_existing_properties() {
        let geometry = super::parse_section_sprms(&[
            0x1F, 0xB0, 0x82, 0x2E, // A4 width: 11906 twips
            0x20, 0xB0, 0xC6, 0x41, // A4 height: 16838 twips
            0x21, 0xB0, 0x00, 0x07, // left margin: 1792 twips
            0x23, 0x90, 0x00, 0xFA, // top margin: -1536 twips
        ]);
        assert_eq!(geometry.width, Some(11906));
        assert_eq!(geometry.top, Some(1536));
        let output = super::patch_document_sections(
            "<w:body><w:p/><w:sectPr><w:footerReference r:id=\"rId1\"/></w:sectPr></w:body>",
            &[geometry],
        );
        assert!(output.contains("<w:pgSz w:w=\"11906\" w:h=\"16838\"/>"));
        assert!(output.contains("<w:pgMar w:top=\"1536\""));
        assert!(output.contains("<w:footerReference r:id=\"rId1\"/>"));
    }
}
