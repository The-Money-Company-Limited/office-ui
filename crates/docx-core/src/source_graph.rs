use serde::{Deserialize, Serialize};

use crate::{structured::package_identity, OoxmlPackage};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceAttribute {
    pub namespace: Option<String>,
    pub name: String,
    pub value: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceElement {
    pub id: String,
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub namespace: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    pub byte_start: usize,
    pub byte_end: usize,
    pub attributes: Vec<SourceAttribute>,
    pub children: Vec<SourceElement>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SourcePart {
    pub name: String,
    pub byte_length: usize,
    pub root: Option<SourceElement>,
    pub error: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSourceGraph {
    pub schema_version: u32,
    pub document_id: String,
    pub parts: Vec<SourcePart>,
    pub assets: Vec<SourcePart>,
    pub warnings: Vec<String>,
}

pub fn document_source_graph(package: &OoxmlPackage) -> DocumentSourceGraph {
    fn element(node: roxmltree::Node<'_, '_>, id: String) -> SourceElement {
        let range = node.range();
        SourceElement {
            id: id.clone(),
            kind: if node.is_element() {
                "element"
            } else if node.is_text() {
                "text"
            } else if node.is_comment() {
                "comment"
            } else {
                "processing_instruction"
            }
            .into(),
            namespace: node.tag_name().namespace().map(str::to_string),
            name: node
                .is_element()
                .then(|| node.tag_name().name().to_string()),
            text: if node.is_element() {
                None
            } else {
                node.text().map(str::to_string)
            },
            byte_start: range.start,
            byte_end: range.end,
            attributes: node
                .attributes()
                .map(|attribute| SourceAttribute {
                    namespace: attribute.namespace().map(str::to_string),
                    name: attribute.name().to_string(),
                    value: attribute.value().to_string(),
                })
                .collect(),
            children: node
                .children()
                .enumerate()
                .map(|(index, child)| element(child, format!("{id}/{index}")))
                .collect(),
        }
    }
    let document_id = package_identity(package);
    let mut parts = package
        .parts
        .values()
        .map(|part| {
            let parsed = roxmltree::Document::parse(&part.content);
            SourcePart {
                name: part.name.clone(),
                byte_length: part.content.len(),
                root: parsed.as_ref().ok().map(|doc| {
                    element(doc.root_element(), format!("{document_id}/{}/0", part.name))
                }),
                error: parsed.err().map(|e| e.to_string()),
            }
        })
        .collect::<Vec<_>>();
    let mut assets = package
        .binary_assets
        .iter()
        .map(|(name, bytes)| SourcePart {
            name: name.clone(),
            byte_length: bytes.len(),
            root: None,
            error: None,
        })
        .collect::<Vec<_>>();
    parts.sort_by(|a, b| a.name.cmp(&b.name));
    assets.sort_by(|a, b| a.name.cmp(&b.name));
    DocumentSourceGraph {
        schema_version: 1,
        document_id,
        parts,
        assets,
        warnings: package.warnings.clone(),
    }
}
