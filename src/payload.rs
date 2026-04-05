use serde::Deserialize;

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RenderWorkbookPayload {
    pub workbook: WorkbookPayload,
    pub theme: ThemePayload,
    #[serde(default)]
    pub memory_mode: Option<String>,
    #[serde(default)]
    pub temp_dir: Option<String>,
    pub target: RenderTargetPayload,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookPayload {
    pub sheet_list: Vec<SheetPayload>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemePayload {
    #[serde(default)]
    pub styles: std::collections::BTreeMap<String, StyleDefinitionPayload>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StyleDefinitionPayload {
    #[serde(default)]
    pub font_color: Option<String>,
    #[serde(default)]
    pub background_color: Option<String>,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub font_size: Option<f64>,
    #[serde(default)]
    pub border: Option<String>,
    #[serde(default)]
    pub align: Option<String>,
    #[serde(default)]
    pub vertical_align: Option<String>,
    #[serde(default)]
    pub text_wrap: Option<bool>,
    #[serde(default)]
    pub num_format: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetPayload {
    pub sheet_name: String,
    #[serde(default)]
    pub header_row_index: Option<u32>,
    #[serde(default)]
    pub header_row_height: Option<f64>,
    #[serde(default)]
    pub freeze_header_row: Option<bool>,
    #[serde(default)]
    pub auto_filter: Option<bool>,
    pub columns: Vec<ColumnPayload>,
    pub rows: Vec<RowPayload>,
    #[serde(default)]
    pub merge_range_list: Vec<MergeRangePayload>,
    #[serde(default)]
    pub conditional_format_list: Vec<ConditionalFormatPayload>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnPayload {
    pub key: String,
    pub title: String,
    pub width: f64,
    #[serde(default)]
    pub data_style: Option<String>,
    #[serde(default)]
    pub header_style: Option<String>,
    #[serde(default)]
    pub success_label_list: Vec<String>,
    #[serde(default)]
    pub danger_label_list: Vec<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RowPayload {
    pub values: serde_json::Map<String, serde_json::Value>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellPayload {
    #[serde(rename = "type")]
    pub cell_type: String,
    #[serde(default)]
    pub value: Option<serde_json::Value>,
    #[serde(default)]
    pub formula: Option<String>,
    #[serde(default)]
    pub url: Option<String>,
    #[serde(default)]
    pub label: Option<String>,
    #[serde(default)]
    pub style_list: Vec<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeRangePayload {
    pub start_row: u32,
    pub start_col: u16,
    pub end_row: u32,
    pub end_col: u16,
    #[serde(default)]
    pub value: Option<String>,
    #[serde(default)]
    pub style_list: Vec<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatPayload {
    pub start_row: u32,
    pub start_col: u16,
    pub end_row: u32,
    pub end_col: u16,
    pub formula: String,
    pub style: String,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum RenderTargetPayload {
    File { path: String },
    Buffer,
}
