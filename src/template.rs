use docx_rs::*;

// This module provides Word-like default styling for docx-rs
// since it doesn't include default templates like python-docx

pub fn apply_page_layout(docx: Docx) -> Docx {
    docx.page_size(12240, 15840) // 8.5 x 11 inches in twentieths of a point
        .page_margin(PageMargin::new().top(1080).bottom(1080).left(1080).right(1080).header(720).footer(720))
}

// Note: docx-rs has limited spacing API compared to python-docx
// For line spacing, we'll use the available API
pub fn apply_line_spacing(para: Paragraph, spacing: f32) -> Paragraph {
    // Use the line_spacing method with a multiplier value
    let spacing_val = (spacing * 240.0) as i32;
    para.line_spacing(LineSpacing::new().line(spacing_val))
}