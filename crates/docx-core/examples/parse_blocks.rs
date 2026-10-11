fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut path = None;
    let mut options = docx_core::ExtractDocumentBlocksOptions::default();
    let mut positional_only = false;
    for argument in std::env::args_os().skip(1) {
        if !positional_only && argument == "--" {
            positional_only = true;
        } else if !positional_only && argument == "--include-words" {
            options.include_words = true;
        } else if path.is_none() {
            path = Some(argument);
        } else {
            return Err("Expected one document path and optional --include-words".into());
        }
    }
    let path = path.ok_or("A document path is required")?;
    let bytes = std::fs::read(path)?;
    let parsed = docx_core::parse_document_blocks_with_options(&bytes, options)?;
    serde_json::to_writer(std::io::stdout().lock(), &parsed)?;
    Ok(())
}
