use docx_core::{build_doc_model, serialize_doc_model};
use docx_core::model::NumberingDefinitionSet;
use docx_core::zip::create_minimal_docx_package;
use docx_core::package::OoxmlPart;

fn typed_list_definitions() -> NumberingDefinitionSet {
    serde_json::from_value(serde_json::json!({
        "abstracts": [{ "abstractNumId": 1, "levels": [{
            "ilvl": 0, "start": 1, "format": "bullet", "text": "•", "suffix": "tab",
            "indent": { "leftTwips": 720, "hangingTwips": 360 }
        }] }],
        "instances": [{ "numId": 2, "abstractNumId": 1, "levelStartOverrides": { "0": 1 } }]
    })).unwrap()
}

#[test]
fn typed_list_definitions_survive_save_and_reopen() {
    let pkg = create_minimal_docx_package(None);
    let mut model = build_doc_model(&pkg);
    model.metadata.numbering_definitions = Some(typed_list_definitions());
    let saved = serialize_doc_model(&model, Some(&pkg));
    assert_eq!(build_doc_model(&saved).metadata.numbering_definitions, model.metadata.numbering_definitions);
    assert!(saved.parts["[Content_Types].xml"].content.contains("/word/numbering.xml"));
    assert!(saved.parts["word/_rels/document.xml.rels"].content.contains("/numbering\""));
    let saved_again = serialize_doc_model(&build_doc_model(&saved), Some(&saved));
    assert_eq!(saved_again.parts["word/numbering.xml"].content, saved.parts["word/numbering.xml"].content);
}

#[test]
fn typed_lists_preserve_imported_numbering_markup_and_part_location() {
    let mut pkg = create_minimal_docx_package(None);
    let original_abstract = r#"<w:abstractNum w:abstractNumId="0"><w:nsid w:val="12345678"/><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/></w:lvl></w:abstractNum>"#;
    let original_instance = r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;
    for (name, content) in [
        ("word/lists/custom.xml", format!(r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{original_abstract}{original_instance}</w:numbering>"#)),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="lists/custom.xml"/></Relationships>"#.to_string()),
    ] {
        pkg.parts.insert(name.to_string(), OoxmlPart { name: name.to_string(), content });
    }
    let mut model = build_doc_model(&pkg);
    let new_definitions = typed_list_definitions();
    let definitions = model.metadata.numbering_definitions.as_mut().unwrap();
    definitions.abstracts.extend(new_definitions.abstracts);
    definitions.instances.extend(new_definitions.instances);
    let saved = serialize_doc_model(&model, Some(&pkg));
    let xml = &saved.parts["word/lists/custom.xml"].content;
    assert!(xml.contains(original_abstract));
    assert!(xml.contains(original_instance));
    assert_eq!(build_doc_model(&saved).metadata.numbering_definitions, model.metadata.numbering_definitions);
    assert!(!saved.parts.contains_key("word/numbering.xml"));
    assert!(xml.find("<w:abstractNum w:abstractNumId=\"1\"").unwrap() < xml.find("<w:num w:numId=\"1\"").unwrap());
}


#[test]
fn explicit_numbering_cancellation_survives_save_and_reopen() {
    use docx_core::model::{DocNode, ParagraphNumbering};

    let mut pkg = create_minimal_docx_package(Some(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="ListPara"/></w:pPr></w:p></w:body></w:document>"#));
    let name = "word/styles.xml";
    pkg.parts.insert(name.to_string(), OoxmlPart {
        name: name.to_string(),
        content: r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="ListPara"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>"#.to_string(),
    });
    let mut model = build_doc_model(&pkg);
    let DocNode::Paragraph(paragraph) = &mut model.nodes[0] else { panic!("expected paragraph"); };
    assert_eq!(paragraph.style.as_ref().unwrap().numbering.as_ref().unwrap().num_id, 1);
    paragraph.style.as_mut().unwrap().numbering = Some(ParagraphNumbering { num_id: 0, ilvl: 0 });
    paragraph.source_xml = None;
    for _ in 0..2 {
        pkg = serialize_doc_model(&model, Some(&pkg));
        assert!(pkg.parts["word/document.xml"].content.contains(r#"<w:numId w:val="0"/>"#));
        model = build_doc_model(&pkg);
        let DocNode::Paragraph(paragraph) = &model.nodes[0] else { panic!("expected paragraph"); };
        let style = paragraph.style.as_ref().unwrap();
        assert_eq!(style.style_id.as_deref(), Some("ListPara"));
        assert!(style.numbering.is_none());
    }
}
