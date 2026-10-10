fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args_os().nth(1).ok_or("A document path is required")?;
    let bytes = std::fs::read(path)?;
    let parsed = docx_core::parse_document_blocks(&bytes)?;
    serde_json::to_writer(std::io::stdout().lock(), &parsed)?;
    Ok(())
}
