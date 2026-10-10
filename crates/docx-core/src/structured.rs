use std::collections::HashSet;

use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};

use crate::model::*;
use crate::package::OoxmlPackage;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum BlockType {
    Text,
    Heading,
    SectionHeading,
    Figure,
    Table,
    TableHead,
    TableCell,
    KeyValue,
    PageNumber,
    Barcode,
    Formula,
    Header,
    Footer,
    Watermark,
    Legend,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceRef {
    pub document_id: String,
    pub node_id: String,
    pub part_name: String,
    pub story: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end: Option<usize>,
    pub offset_encoding: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentBlock {
    pub object: String,
    pub id: String,
    pub r#type: BlockType,
    pub content: String,
    pub details: Value,
    pub source_refs: Vec<SourceRef>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent_block_id: Option<String>,
    pub children: Vec<DocumentBlock>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StructuredDocument {
    pub schema_version: u32,
    pub document_id: String,
    pub blocks: Vec<DocumentBlock>,
    pub warnings: Vec<String>,
    pub geometry_status: String,
}

pub fn package_identity(package: &OoxmlPackage) -> String {
    let mut hash = Sha256::new();
    let mut parts: Vec<_> = package.parts.iter().collect();
    parts.sort_by_key(|(name, _)| *name);
    for (name, part) in parts {
        hash.update((name.len() as u64).to_le_bytes());
        hash.update(name.as_bytes());
        hash.update((part.content.len() as u64).to_le_bytes());
        hash.update(part.content.as_bytes());
    }
    let mut assets: Vec<_> = package.binary_assets.iter().collect();
    assets.sort_by_key(|(name, _)| *name);
    for (name, bytes) in assets {
        hash.update((name.len() as u64).to_le_bytes());
        hash.update(name.as_bytes());
        hash.update((bytes.len() as u64).to_le_bytes());
        hash.update(bytes);
    }
    format!("{:x}", hash.finalize())
}

pub fn assign_source_ids(model: &mut DocModel, document_id: &str) {
    fn paragraph(node: &mut ParagraphNode, id: String) {
        node.block_id = Some(id.clone());
        for (index, child) in node.children.iter_mut().enumerate() {
            if let ParagraphChildNode::Image(image) = child {
                let image_id = format!("{id}/child/{index}");
                image.block_id = Some(image_id.clone());
                if let Some(content) = &mut image.content_nodes {
                    nodes(content, &format!("{image_id}/content"));
                }
            }
        }
    }
    fn table(node: &mut TableNode, id: String) {
        node.block_id = Some(id.clone());
        for (ri, row) in node.rows.iter_mut().enumerate() {
            for (ci, cell) in row.cells.iter_mut().enumerate() {
                for (ni, child) in cell.nodes.iter_mut().enumerate() {
                    let id = format!("{id}/row/{ri}/cell/{ci}/{ni}");
                    match child {
                        TableCellContentNode::Paragraph(p) => paragraph(p, id),
                        TableCellContentNode::Table(t) => table(t, id),
                    }
                }
            }
        }
    }
    fn nodes(content: &mut [DocNode], root: &str) {
        for (index, node) in content.iter_mut().enumerate() {
            let id = format!("{root}/{index}");
            match node {
                DocNode::Paragraph(p) => paragraph(p, id),
                DocNode::Table(t) => table(t, id),
            }
        }
    }
    model.metadata.document_id = Some(document_id.to_string());
    let root = format!("docx:{document_id}");
    nodes(&mut model.nodes, &format!("{root}/word/document.xml/body"));
    for header in &mut model.metadata.header_sections {
        nodes(&mut header.nodes, &format!("{root}/{}", header.part_name));
    }
    for footer in &mut model.metadata.footer_sections {
        nodes(&mut footer.nodes, &format!("{root}/{}", footer.part_name));
    }
    for section in model.metadata.sections.iter_mut().flatten() {
        for header in &mut section.header_sections {
            nodes(&mut header.nodes, &format!("{root}/{}", header.part_name));
        }
        for footer in &mut section.footer_sections {
            nodes(&mut footer.nodes, &format!("{root}/{}", footer.part_name));
        }
    }
    for (kind, notes) in [
        ("footnote", &mut model.metadata.footnotes),
        ("endnote", &mut model.metadata.endnotes),
    ] {
        for note in notes.iter_mut().flatten() {
            if let Some(content) = &mut note.nodes {
                nodes(
                    content,
                    &format!("{root}/word/{kind}s.xml/{kind}/{}", note.id),
                );
            }
        }
    }
    for comment in model.metadata.comments.iter_mut().flatten() {
        if let Some(content) = &mut comment.nodes {
            nodes(
                content,
                &format!("{root}/word/comments.xml/comment/{}", comment.id),
            );
        }
    }
}

struct Extractor {
    document_id: String,
    part_name: String,
    story: String,
}

impl Extractor {
    fn block(
        &self,
        id: String,
        kind: BlockType,
        content: String,
        details: Value,
        children: Vec<DocumentBlock>,
        parent: Option<&str>,
    ) -> DocumentBlock {
        DocumentBlock {
            object: "block".into(),
            id: id.clone(),
            r#type: kind,
            content,
            details,
            source_refs: vec![SourceRef {
                document_id: self.document_id.clone(),
                node_id: id,
                part_name: self.part_name.clone(),
                story: self.story.clone(),
                start: None,
                end: None,
                offset_encoding: "utf-16".into(),
            }],
            parent_block_id: parent.map(str::to_string),
            children,
        }
    }

    fn nodes(&self, nodes: &[DocNode], parent: Option<&str>) -> Vec<DocumentBlock> {
        nodes
            .iter()
            .enumerate()
            .map(|(index, node)| {
                let fallback = format!("{}/{}", parent.unwrap_or("document"), index);
                match node {
                    DocNode::Paragraph(p) => self.paragraph(p, &fallback, parent),
                    DocNode::Table(t) => self.table(t, &fallback, parent),
                }
            })
            .collect()
    }

    fn paragraph(&self, p: &ParagraphNode, fallback: &str, parent: Option<&str>) -> DocumentBlock {
        let id = p.block_id.as_deref().unwrap_or(fallback);
        let style = p.style.as_ref();
        let style_id = style.and_then(|s| s.style_id.as_deref()).unwrap_or("");
        let field_instructions = p
            .source_xml
            .as_deref()
            .map(|xml| {
                crate::xml::extract_balanced_tag_blocks(xml, "w:instrText")
                    .iter()
                    .map(|field| {
                        field
                            .split_once('>')
                            .map(|(_, value)| {
                                crate::xml::decode_xml_entities(
                                    value.split("</").next().unwrap_or(""),
                                )
                            })
                            .unwrap_or_default()
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let joined_instructions = field_instructions.join("").to_uppercase();
        let instruction_kind = joined_instructions.split_whitespace().next().unwrap_or("");
        let kind = if matches!(instruction_kind, "PAGE" | "NUMPAGES" | "SECTIONPAGES") {
            BlockType::PageNumber
        } else if matches!(instruction_kind, "DISPLAYBARCODE" | "MERGEBARCODE") {
            BlockType::Barcode
        } else if style.and_then(|s| s.heading_level).is_some() {
            BlockType::Heading
        } else if style_id.eq_ignore_ascii_case("title")
            || style_id.eq_ignore_ascii_case("subtitle")
        {
            BlockType::SectionHeading
        } else if style_id.eq_ignore_ascii_case("caption") {
            BlockType::Legend
        } else {
            BlockType::Text
        };
        let mut content = String::new();
        let mut inline = Vec::new();
        let mut children = Vec::new();
        for (index, child) in p.children.iter().enumerate() {
            let child_id = format!("{id}/child/{index}");
            match child {
                ParagraphChildNode::Text(run) => {
                    let start = content.encode_utf16().count();
                    content.push_str(&run.text);
                    inline.push(json!({"type":"text", "start":start, "end":content.encode_utf16().count(), "text":run.text, "style":run.style, "link":run.link, "noteReference":run.note_reference}));
                }
                ParagraphChildNode::FormField(field) => {
                    let value = field.value.clone().unwrap_or_else(|| {
                        field
                            .checked
                            .map(|checked| if checked { "☒" } else { "☐" }.to_string())
                            .unwrap_or_default()
                    });
                    content.push_str(&value);
                    children.push(self.block(
                        child_id,
                        BlockType::KeyValue,
                        value,
                        json!({"kind":"form_field", "field":field}),
                        Vec::new(),
                        Some(id),
                    ));
                }
                ParagraphChildNode::Image(image) => {
                    let kind = match image.semantic_kind.as_deref() {
                        Some("formula") => BlockType::Formula,
                        Some("watermark") => BlockType::Watermark,
                        Some("barcode") => BlockType::Barcode,
                        _ => BlockType::Figure,
                    };
                    let image_id = image.block_id.as_deref().unwrap_or(&child_id);
                    let nested = image
                        .content_nodes
                        .as_ref()
                        .map(|nodes| self.nodes(nodes, Some(image_id)))
                        .unwrap_or_default();
                    let text = if nested.is_empty() {
                        image
                            .text_box_text
                            .clone()
                            .or_else(|| image.alt.clone())
                            .unwrap_or_default()
                    } else {
                        nested
                            .iter()
                            .map(|b| b.content.as_str())
                            .collect::<Vec<_>>()
                            .join("\n")
                    };
                    children.push(self.block(image_id.to_string(), kind, text, json!({"kind":image.semantic_kind.as_deref().unwrap_or(if image.synthetic_text_box == Some(true) {"text_box"} else {"image"}), "asset":image.part_name, "alt":image.alt, "contentType":image.content_type, "width":image.width_px, "height":image.height_px, "floating":image.floating}), nested, Some(id)));
                }
            }
        }
        self.block(id.to_string(), kind, content, json!({"kind":"paragraph", "styleId":style_id, "headingLevel":style.and_then(|s|s.heading_level), "numbering":style.and_then(|s|s.numbering.as_ref()), "inlines":inline, "fieldInstructions":field_instructions}), children, parent)
    }

    fn table(&self, table: &TableNode, fallback: &str, parent: Option<&str>) -> DocumentBlock {
        let id = table.block_id.as_deref().unwrap_or(fallback);
        let mut children = Vec::new();
        let mut row_text = Vec::new();
        for (ri, row) in table.rows.iter().enumerate() {
            let row_id = format!("{id}/row/{ri}");
            let is_header = row
                .style
                .as_ref()
                .and_then(|s| s.is_header)
                .unwrap_or(false);
            let mut cells = Vec::new();
            let mut cell_text = Vec::new();
            for (ci, cell) in row.cells.iter().enumerate() {
                let cell_id = format!("{row_id}/cell/{ci}");
                let content = cell
                    .nodes
                    .iter()
                    .enumerate()
                    .map(|(ni, node)| match node {
                        TableCellContentNode::Paragraph(p) => {
                            self.paragraph(p, &format!("{cell_id}/{ni}"), Some(&cell_id))
                        }
                        TableCellContentNode::Table(t) => {
                            self.table(t, &format!("{cell_id}/{ni}"), Some(&cell_id))
                        }
                    })
                    .collect::<Vec<_>>();
                let text = content
                    .iter()
                    .map(|b| b.content.as_str())
                    .collect::<Vec<_>>()
                    .join("\n");
                cell_text.push(text.clone());
                cells.push(self.block(cell_id, BlockType::TableCell, text, json!({"rowIndex":ri, "columnIndex":ci, "isHeader":is_header, "style":cell.style}), content, Some(if is_header { &row_id } else { id })));
            }
            let text = cell_text.join("\t");
            row_text.push(text.clone());
            if is_header {
                children.push(self.block(
                    row_id,
                    BlockType::TableHead,
                    text,
                    json!({"rowIndex":ri}),
                    cells,
                    Some(id),
                ));
            } else {
                children.extend(cells);
            }
        }
        self.block(
            id.to_string(),
            BlockType::Table,
            row_text.join("\n"),
            json!({"rowCount":table.rows.len()}),
            children,
            parent,
        )
    }
}

pub fn extract_document_blocks(model: &DocModel) -> StructuredDocument {
    let document_id = model
        .metadata
        .document_id
        .clone()
        .unwrap_or_else(|| "unsaved".into());
    let mut extractor = Extractor {
        document_id: document_id.clone(),
        part_name: "word/document.xml".into(),
        story: "body".into(),
    };
    let mut blocks = extractor.nodes(&model.nodes, None);
    let mut seen = HashSet::new();
    let headers = model.metadata.header_sections.iter().chain(
        model
            .metadata
            .sections
            .iter()
            .flatten()
            .flat_map(|s| s.header_sections.iter()),
    );
    for header in headers {
        if !seen.insert(header.part_name.clone()) {
            continue;
        }
        extractor.part_name = header.part_name.clone();
        extractor.story = "header".into();
        blocks.extend(
            extractor
                .nodes(&header.nodes, None)
                .into_iter()
                .map(|mut b| {
                    b.details["contentType"] = json!(b.r#type);
                    b.r#type = BlockType::Header;
                    b
                }),
        );
    }
    let footers = model.metadata.footer_sections.iter().chain(
        model
            .metadata
            .sections
            .iter()
            .flatten()
            .flat_map(|s| s.footer_sections.iter()),
    );
    for footer in footers {
        if !seen.insert(footer.part_name.clone()) {
            continue;
        }
        extractor.part_name = footer.part_name.clone();
        extractor.story = "footer".into();
        blocks.extend(
            extractor
                .nodes(&footer.nodes, None)
                .into_iter()
                .map(|mut b| {
                    b.details["contentType"] = json!(b.r#type);
                    b.r#type = BlockType::Footer;
                    b
                }),
        );
    }
    for (kind, notes) in [
        ("footnote", &model.metadata.footnotes),
        ("endnote", &model.metadata.endnotes),
    ] {
        extractor.part_name = format!("word/{kind}s.xml");
        extractor.story = kind.into();
        for note in notes.iter().flatten() {
            if let Some(nodes) = &note.nodes {
                blocks.extend(extractor.nodes(nodes, None));
            }
        }
    }
    extractor.part_name = "word/comments.xml".into();
    extractor.story = "comment".into();
    for comment in model.metadata.comments.iter().flatten() {
        if let Some(nodes) = &comment.nodes {
            blocks.extend(extractor.nodes(nodes, None));
        }
    }
    StructuredDocument {
        schema_version: 1,
        document_id,
        blocks,
        warnings: model.metadata.warnings.clone(),
        geometry_status: "requires_layout".into(),
    }
}
