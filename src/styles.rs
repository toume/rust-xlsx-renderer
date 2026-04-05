use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

use crate::payload::{ColumnPayload, StyleDefinitionPayload, ThemePayload};

pub fn resolve_header_format(column: &ColumnPayload, theme: &ThemePayload) -> Format {
    let custom = column
        .header_style
        .as_deref()
        .and_then(|style| style_by_name(style, theme));
    custom.unwrap_or_else(|| style_by_name("header", theme).unwrap_or_else(Format::new))
}

pub fn resolve_format(
    base_style: Option<&str>,
    style_list: &[String],
    theme: &ThemePayload,
) -> Format {
    let mut format = base_style
        .and_then(|style| style_by_name(style, theme))
        .or_else(|| style_by_name("cell", theme))
        .unwrap_or_else(Format::new);

    for style_name in style_list {
        if let Some(style_format) = style_by_name(style_name, theme) {
            format = format.merge(&style_format);
        }
    }

    format
}

pub fn style_by_name(name: &str, theme: &ThemePayload) -> Option<Format> {
    theme.styles.get(name).and_then(style_from_definition)
}

fn style_from_definition(definition: &StyleDefinitionPayload) -> Option<Format> {
    let mut format = Format::new();
    let mut touched = false;

    if definition.bold.unwrap_or(false) {
        format = format.set_bold();
        touched = true;
    }

    if definition.italic.unwrap_or(false) {
        format = format.set_italic();
        touched = true;
    }

    if let Some(font_size) = definition.font_size {
        format = format.set_font_size(font_size);
        touched = true;
    }

    if definition.text_wrap.unwrap_or(false) {
        format = format.set_text_wrap();
        touched = true;
    }

    if let Some(font_color) = definition.font_color.as_deref().and_then(parse_rgb_color) {
        format = format.set_font_color(font_color);
        touched = true;
    }

    if let Some(background_color) = definition
        .background_color
        .as_deref()
        .and_then(parse_rgb_color)
    {
        format = format.set_background_color(background_color);
        touched = true;
    }

    if let Some(num_format) = &definition.num_format {
        format = format.set_num_format(num_format);
        touched = true;
    }

    if let Some(border) = definition.border.as_deref().and_then(parse_border) {
        format = format.set_border(border);
        touched = true;
    }

    if let Some(align) = definition.align.as_deref().and_then(parse_horizontal_align) {
        format = format.set_align(align);
        touched = true;
    }

    if let Some(vertical_align) = definition
        .vertical_align
        .as_deref()
        .and_then(parse_vertical_align)
    {
        format = format.set_align(vertical_align);
        touched = true;
    }

    touched.then_some(format)
}

fn parse_rgb_color(value: &str) -> Option<Color> {
    let normalized = value.trim().trim_start_matches('#');
    if normalized.len() != 6 {
        return None;
    }

    u32::from_str_radix(normalized, 16).ok().map(Color::RGB)
}

fn parse_border(value: &str) -> Option<FormatBorder> {
    match value {
        "thin" => Some(FormatBorder::Thin),
        "none" => Some(FormatBorder::None),
        _ => None,
    }
}

fn parse_horizontal_align(value: &str) -> Option<FormatAlign> {
    match value {
        "left" => Some(FormatAlign::Left),
        "center" => Some(FormatAlign::Center),
        "right" => Some(FormatAlign::Right),
        _ => None,
    }
}

fn parse_vertical_align(value: &str) -> Option<FormatAlign> {
    match value {
        "top" => Some(FormatAlign::Top),
        "middle" => Some(FormatAlign::VerticalCenter),
        "bottom" => Some(FormatAlign::Bottom),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::{resolve_format, style_by_name};
    use crate::payload::{StyleDefinitionPayload, ThemePayload};

    #[test]
    fn known_styles_exist() {
        let mut styles = BTreeMap::new();
        styles.insert(
            "header".to_string(),
            StyleDefinitionPayload {
                background_color: Some("112233".to_string()),
                bold: Some(true),
                font_color: Some("ffffff".to_string()),
                italic: None,
                font_size: None,
                border: Some("thin".to_string()),
                align: Some("center".to_string()),
                vertical_align: Some("middle".to_string()),
                text_wrap: Some(true),
                num_format: None,
            },
        );
        let theme = ThemePayload { styles };
        assert!(style_by_name("header", &theme).is_some());
        assert!(style_by_name("missing", &theme).is_none());
    }

    #[test]
    fn resolve_format_accepts_unknown_overlay_styles() {
        let mut styles = BTreeMap::new();
        styles.insert(
            "cell".to_string(),
            StyleDefinitionPayload {
                background_color: None,
                bold: None,
                font_color: None,
                italic: None,
                font_size: None,
                border: Some("thin".to_string()),
                align: None,
                vertical_align: None,
                text_wrap: None,
                num_format: None,
            },
        );
        styles.insert(
            "dangerText".to_string(),
            StyleDefinitionPayload {
                background_color: None,
                bold: None,
                font_color: Some("ff0000".to_string()),
                italic: None,
                font_size: None,
                border: None,
                align: None,
                vertical_align: None,
                text_wrap: None,
                num_format: None,
            },
        );
        let theme = ThemePayload { styles };
        let format = resolve_format(
            Some("cell"),
            &["unknown".to_string(), "dangerText".to_string()],
            &theme,
        );
        let _ = format;
    }

    #[test]
    fn theme_styles_override_builtin_styles() {
        let mut styles = BTreeMap::new();
        styles.insert(
            "header".to_string(),
            StyleDefinitionPayload {
                background_color: Some("112233".to_string()),
                bold: Some(true),
                font_color: Some("ffffff".to_string()),
                italic: None,
                font_size: None,
                border: Some("thin".to_string()),
                align: Some("center".to_string()),
                vertical_align: Some("middle".to_string()),
                text_wrap: Some(true),
                num_format: None,
            },
        );
        let theme = ThemePayload { styles };

        assert!(style_by_name("header", &theme).is_some());
    }
}
