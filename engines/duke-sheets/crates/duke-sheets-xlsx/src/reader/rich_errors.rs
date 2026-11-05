//! Excel saves the errors that came with dynamic arrays as #VALUE! and points the cell's `vm` at a
//! rich value holding the real one: metadata.xml maps the cell's value metadata to a rich value,
//! rdrichvalue.xml holds it, and rdrichvaluestructure.xml names its keys (an `_error` structure
//! with an `errorType`).

use std::collections::HashMap;

use duke_sheets_core::CellError;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

/// The error each `vm` index (1-based, as on a cell) stands for, where it is one the cell's
/// `<v>` can't hold. Parts that are missing or unreadable give no entries.
pub(crate) fn value_errors(metadata: &[u8], rich_values: &[u8], structures: &[u8]) -> HashMap<u32, CellError> {
    // metadata.xml: <metadataType name=…> in order (rc t is 1-based into them), the XLRICHVALUE
    // <futureMetadata> blocks' <xlrd:rvb i=…> (rc v is 0-based into them), and <valueMetadata>
    // blocks' <rc t v> (a cell's vm is 1-based into them).
    let mut types = Vec::new();
    let mut rich_value_of_block = Vec::new();
    let mut value_blocks = Vec::new();
    let mut section = "";
    each_element(metadata, |name, e| match name {
        b"metadataType" => types.push(attr(e, b"name").unwrap_or_default()),
        b"futureMetadata" => section = if attr(e, b"name").as_deref() == Some("XLRICHVALUE") { "rich" } else { "" },
        b"valueMetadata" => section = "values",
        b"cellMetadata" => section = "",
        b"rvb" if section == "rich" => rich_value_of_block.push(number(e, b"i")),
        b"rc" if section == "values" => value_blocks.push((number(e, b"t"), number(e, b"v"))),
        _ => {}
    });

    // rdrichvaluestructure.xml: <s t=…> with its <k n=…> keys in order.
    let mut keys: Vec<(String, Vec<String>)> = Vec::new();
    each_element(structures, |name, e| match name {
        b"s" => keys.push((attr(e, b"t").unwrap_or_default(), Vec::new())),
        b"k" => {
            if let Some((_, names)) = keys.last_mut() {
                names.push(attr(e, b"n").unwrap_or_default());
            }
        }
        _ => {}
    });

    // rdrichvalue.xml: <rv s=…> with one <v> per key.
    let mut values: Vec<(usize, Vec<String>)> = Vec::new();
    let mut reader = Reader::from_reader(rich_values);
    let mut buf = Vec::new();
    let mut in_value = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"rv" => {
                values.push((number(&e, b"s").unwrap_or(usize::MAX), Vec::new()))
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"v" => in_value = true,
            Ok(Event::Text(t)) if in_value => {
                if let (Some((_, vs)), Ok(text)) = (values.last_mut(), t.unescape()) {
                    vs.push(text.into_owned());
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"v" => in_value = false,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let rich_type = types.iter().position(|t| t == "XLRICHVALUE").map(|i| i + 1);
    let mut errors = HashMap::new();
    for (vm, (t, v)) in value_blocks.into_iter().enumerate() {
        if t.is_none() || t != rich_type {
            continue;
        }
        let Some(&Some(rich)) = v.and_then(|v| rich_value_of_block.get(v)) else { continue };
        let Some((structure, fields)) = values.get(rich) else { continue };
        let Some((kind, names)) = keys.get(*structure) else { continue };
        if kind != "_error" {
            continue;
        }
        let error_type = names.iter().position(|n| n == "errorType").and_then(|i| fields.get(i));
        // 8 is #SPILL!, the one Excel shows where a dynamic array can't spill.
        if error_type.map(String::as_str) == Some("8") {
            errors.insert(vm as u32 + 1, CellError::Spill);
        }
    }
    errors
}

fn each_element(xml: &[u8], mut visit: impl FnMut(&[u8], &BytesStart)) {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => visit(e.local_name().as_ref(), &e),
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn attr(e: &BytesStart, key: &[u8]) -> Option<String> {
    e.attributes()
        .flatten()
        .find(|a| a.key.local_name().as_ref() == key)
        .and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
}

fn number(e: &BytesStart, key: &[u8]) -> Option<usize> {
    attr(e, key)?.parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    // As Excel saved `=SUMPRODUCT(A:A=A3)*H:J`, which can't spill.
    const METADATA: &str = r#"<metadata xmlns:xlrd="x"><metadataTypes count="2"><metadataType name="XLDAPR"/><metadataType name="XLRICHVALUE"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext><xda:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata><valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata></metadata>"#;
    const VALUES: &str = r#"<rvData count="1"><rv s="0"><v>0</v><v>8</v><v>0</v><v>3</v></rv></rvData>"#;
    const STRUCTURES: &str = r#"<rvStructures count="1"><s t="_error"><k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/></s></rvStructures>"#;

    #[test]
    fn spill_error_saved_as_a_rich_value() {
        let errors = value_errors(METADATA.as_bytes(), VALUES.as_bytes(), STRUCTURES.as_bytes());
        assert_eq!(errors.get(&1), Some(&CellError::Spill));
        assert_eq!(errors.len(), 1);
    }

    #[test]
    fn other_rich_values_are_not_errors() {
        let values = r#"<rvData count="1"><rv s="0"><v>0</v><v>2</v><v>0</v><v>0</v></rv></rvData>"#;
        assert!(value_errors(METADATA.as_bytes(), values.as_bytes(), STRUCTURES.as_bytes()).is_empty());
        assert!(value_errors(b"", b"", b"").is_empty());
    }
}
