use std::collections::HashMap;

use serde::{Deserialize, Deserializer, Serialize, Serializer};

mod byte_asset_map {
    use super::*;

    pub fn serialize<S>(map: &HashMap<String, Vec<u8>>, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        use serde::ser::SerializeMap;
        let mut ser_map = serializer.serialize_map(Some(map.len()))?;
        for (key, value) in map {
            ser_map.serialize_entry(key, &serde_bytes::Bytes::new(value))?;
        }
        ser_map.end()
    }

    pub fn deserialize<'de, D>(deserializer: D) -> Result<HashMap<String, Vec<u8>>, D::Error>
    where
        D: Deserializer<'de>,
    {
        let raw = HashMap::<String, serde_bytes::ByteBuf>::deserialize(deserializer)?;
        Ok(raw.into_iter().map(|(key, value)| (key, value.into_vec())).collect())
    }
}

/// Mirrors TypeScript `OoxmlPart`.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OoxmlPart {
    pub name: String,
    pub content: String,
}

/// Mirrors TypeScript `OoxmlPackage`.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OoxmlPackage {
    pub parts: HashMap<String, OoxmlPart>,
    #[serde(with = "byte_asset_map")]
    pub binary_assets: HashMap<String, Vec<u8>>,
    /// Import diagnostics that cannot be represented by OOXML itself. This is
    /// empty for native DOCX packages and populated by format adapters such as
    /// the legacy binary DOC importer.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub warnings: Vec<String>,
}

pub fn get_part<'a>(pkg: &'a OoxmlPackage, part_name: &str) -> Option<&'a OoxmlPart> {
    pkg.parts.get(part_name)
}

pub fn with_part(pkg: &OoxmlPackage, part: OoxmlPart) -> OoxmlPackage {
    let mut parts = pkg.parts.clone();
    parts.insert(part.name.clone(), part);
    OoxmlPackage {
        parts,
        binary_assets: pkg.binary_assets.clone(),
        warnings: pkg.warnings.clone(),
    }
}

pub fn relationship_target(pkg: &OoxmlPackage, source: &str, kind: &str) -> Option<String> {
    let rel_name = if source.is_empty() { "_rels/.rels".to_string() } else { crate::parse::relationship_part_name_for_part(source) };
    let part = pkg.parts.get(&rel_name)?;
    let xml = roxmltree::Document::parse(&part.content).ok()?;
    let rel = xml.descendants().find(|node| node.is_element() && node.tag_name().name() == "Relationship"
        && node.attribute("Type").is_some_and(|value| value.rsplit('/').next() == Some(kind))
        && node.attribute("TargetMode") != Some("External"))?;
    Some(crate::parse::resolve_part_path(source, rel.attribute("Target")?))
}

pub fn main_document_part_name(pkg: &OoxmlPackage) -> String {
    relationship_target(pkg, "", "officeDocument").unwrap_or_else(|| "word/document.xml".into())
}

pub fn related_document_part_name(pkg: &OoxmlPackage, kind: &str) -> String {
    relationship_target(pkg, &main_document_part_name(pkg), kind).unwrap_or_else(|| {
        if kind == "theme" { "word/theme/theme1.xml".into() } else { format!("word/{kind}.xml") }
    })
}

pub fn related_document_part<'a>(pkg: &'a OoxmlPackage, kind: &str) -> Option<&'a OoxmlPart> {
    pkg.parts.get(&related_document_part_name(pkg, kind))
}
