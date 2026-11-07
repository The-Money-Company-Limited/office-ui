//! BIFF8 formula token decompiler.

pub mod ptg;
pub mod token_parser;

pub use duke_sheets_formula::decompile::decompiler;
pub use duke_sheets_formula::decompile::function_table;
pub use duke_sheets_formula::decompile::{
    ExternName, ExternSheetEntry, FormulaContext, NameRecord, SupBook, BIFF8_LAST_CELL,
    BUILTIN_NAMES,
};

/// Decompile BIFF8 formula token bytes into a human-readable formula string.
pub fn decompile(data: &[u8], ctx: &FormulaContext) -> String {
    decompile_with_extra(data, &[], ctx)
}

/// Decompile a shared formula's tokens for its cell at `base_cell`.
pub fn decompile_shared(data: &[u8], ctx: &FormulaContext, base_cell: (u32, u16)) -> String {
    if data.is_empty() {
        return String::new();
    }
    decompiler::decompile_at(&token_parser::parse_tokens(data), ctx, Some(base_cell))
}

/// Decompile formula tokens with an extra-data section for tArray constants.
pub fn decompile_with_extra(data: &[u8], extra_data: &[u8], ctx: &FormulaContext) -> String {
    if data.is_empty() {
        return String::new();
    }
    let tokens = token_parser::parse_tokens_with_extra(data, extra_data);
    decompiler::decompile(&tokens, ctx)
}
