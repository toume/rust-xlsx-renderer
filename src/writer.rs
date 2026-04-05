use std::io::{Seek, Write};

use napi::Result;
use rust_xlsxwriter::{ConditionalFormatFormula, Formula, Url, Workbook, Worksheet};

use crate::error::{
    invalid_cell, invalid_cell_type, invalid_conditional_style, invalid_link, invalid_number,
    xlsx_error,
};
use crate::memory_mode::MemoryMode;
use crate::payload::{CellPayload, ColumnPayload, RenderWorkbookPayload, SheetPayload};
use crate::styles::{resolve_format, resolve_header_format, style_by_name};
use crate::utils::{
    escape_excel_text, excel_column_name, extract_number, extract_string, parse_iso_date,
};

pub fn write_workbook<W: Write + Seek + Send>(
    payload: &RenderWorkbookPayload,
    writer: &mut W,
) -> Result<()> {
    let mut workbook = Workbook::new();

    if let Some(temp_dir) = &payload.temp_dir {
        workbook.set_tempdir(temp_dir).map_err(xlsx_error)?;
    }

    let memory_mode = MemoryMode::parse(payload.memory_mode.as_deref())?;

    for sheet in &payload.workbook.sheet_list {
        let worksheet = match memory_mode {
            MemoryMode::Standard => workbook.add_worksheet(),
            MemoryMode::LowMemory => workbook.add_worksheet_with_low_memory(),
            MemoryMode::ConstantMemory => workbook.add_worksheet_with_constant_memory(),
        };

        write_sheet(worksheet, sheet, &payload.theme)?;
    }

    workbook.save_to_writer(writer).map_err(xlsx_error)
}

fn write_sheet(
    worksheet: &mut Worksheet,
    sheet: &SheetPayload,
    theme: &crate::payload::ThemePayload,
) -> Result<()> {
    let header_row_index = sheet.header_row_index.unwrap_or(0);
    let header_row_height = sheet.header_row_height.unwrap_or(34.0);

    worksheet.set_name(&sheet.sheet_name).map_err(xlsx_error)?;

    for (idx, column) in sheet.columns.iter().enumerate() {
        let col_index = idx as u16;
        worksheet
            .set_column_width(col_index, column.width)
            .map_err(xlsx_error)?;

        let header_format = resolve_header_format(column, theme);
        worksheet
            .write_with_format(
                header_row_index,
                col_index,
                column.title.as_str(),
                &header_format,
            )
            .map_err(xlsx_error)?;
    }

    worksheet
        .set_row_height(header_row_index, header_row_height)
        .map_err(xlsx_error)?;

    if sheet.freeze_header_row.unwrap_or(true) {
        worksheet
            .set_freeze_panes(header_row_index + 1, 0)
            .map_err(xlsx_error)?;
    }

    for merge in &sheet.merge_range_list {
        let format = resolve_format(None, &merge.style_list, theme);
        worksheet
            .merge_range(
                merge.start_row,
                merge.start_col,
                merge.end_row,
                merge.end_col,
                merge.value.clone().unwrap_or_default().as_str(),
                &format,
            )
            .map_err(xlsx_error)?;
    }

    for conditional in &sheet.conditional_format_list {
        let format = style_by_name(&conditional.style, theme).ok_or_else(|| {
            invalid_conditional_style(format!(
                "unknown xlsx conditional style: {}",
                conditional.style
            ))
        })?;

        let conditional_format = ConditionalFormatFormula::new()
            .set_rule(conditional.formula.as_str())
            .set_format(&format);

        worksheet
            .add_conditional_format(
                conditional.start_row,
                conditional.start_col,
                conditional.end_row,
                conditional.end_col,
                &conditional_format,
            )
            .map_err(xlsx_error)?;
    }

    let first_data_row = header_row_index + 1;
    for (row_offset, row) in sheet.rows.iter().enumerate() {
        let row_index = first_data_row + row_offset as u32;
        write_row(worksheet, row_index, &sheet.columns, row, theme)?;
    }

    apply_label_conditional_formats(worksheet, sheet, first_data_row, theme)?;

    if sheet.auto_filter.unwrap_or(true) && !sheet.columns.is_empty() && !sheet.rows.is_empty() {
        let last_row = first_data_row + sheet.rows.len() as u32 - 1;
        let last_col = sheet.columns.len() as u16 - 1;
        worksheet
            .autofilter(header_row_index, 0, last_row, last_col)
            .map_err(xlsx_error)?;
    }

    Ok(())
}

fn write_row(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[ColumnPayload],
    row: &crate::payload::RowPayload,
    theme: &crate::payload::ThemePayload,
) -> Result<()> {
    for (col_index, column) in columns.iter().enumerate() {
        let col_index = col_index as u16;
        let base_style = column.data_style.as_deref();
        let base_format = resolve_format(base_style, &[], theme);
        let raw_value = row.values.get(&column.key);

        match raw_value {
            None | Some(serde_json::Value::Null) => {
                worksheet
                    .write_blank(row_index, col_index, &base_format)
                    .map_err(xlsx_error)?;
            }
            Some(serde_json::Value::String(value)) => {
                worksheet
                    .write_with_format(row_index, col_index, value.as_str(), &base_format)
                    .map_err(xlsx_error)?;
            }
            Some(serde_json::Value::Number(value)) => {
                let number = value.as_f64().ok_or_else(|| {
                    invalid_number(format!("invalid xlsx number for column {}", column.key))
                })?;
                worksheet
                    .write_with_format(row_index, col_index, number, &base_format)
                    .map_err(xlsx_error)?;
            }
            Some(serde_json::Value::Bool(value)) => {
                worksheet
                    .write_with_format(row_index, col_index, *value, &base_format)
                    .map_err(xlsx_error)?;
            }
            Some(serde_json::Value::Object(_)) => {
                let cell: CellPayload = serde_json::from_value(raw_value.cloned().unwrap())
                    .map_err(|error| {
                        invalid_cell(format!(
                            "invalid xlsx cell for column {}: {}",
                            column.key, error
                        ))
                    })?;

                let format = resolve_format(base_style, &cell.style_list, theme);
                write_cell_object(worksheet, row_index, col_index, &cell, &format)?;
            }
            Some(_) => {
                return Err(invalid_cell(format!(
                    "unsupported xlsx value for column {}",
                    column.key
                )))
            }
        }
    }

    Ok(())
}

fn write_cell_object(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    cell: &CellPayload,
    format: &rust_xlsxwriter::Format,
) -> Result<()> {
    match cell.cell_type.as_str() {
        "empty" => {
            worksheet
                .write_blank(row, col, format)
                .map_err(xlsx_error)?;
            Ok(())
        }
        "string" => worksheet
            .write_with_format(
                row,
                col,
                extract_string(cell.value.as_ref())
                    .unwrap_or_default()
                    .as_str(),
                format,
            )
            .map_err(xlsx_error)
            .map(|_| ()),
        "number" => worksheet
            .write_with_format(
                row,
                col,
                extract_number(cell.value.as_ref())
                    .ok_or_else(|| invalid_number("number cell is missing a numeric value"))?,
                format,
            )
            .map_err(xlsx_error)
            .map(|_| ()),
        "formula" => worksheet
            .write_with_format(
                row,
                col,
                Formula::new(cell.formula.clone().unwrap_or_default().as_str()),
                format,
            )
            .map_err(xlsx_error)
            .map(|_| ()),
        "link" => {
            let url = cell
                .url
                .clone()
                .ok_or_else(|| invalid_link("link cell is missing a url"))?;
            let mut link = Url::new(url.as_str());
            if let Some(label) = &cell.label {
                link = link.set_text(label.as_str());
            }
            worksheet
                .write_with_format(row, col, link, format)
                .map_err(xlsx_error)
                .map(|_| ())
        }
        "date" => {
            if let Some(value) = extract_string(cell.value.as_ref()) {
                if let Some(date) = parse_iso_date(&value) {
                    worksheet
                        .write_with_format(row, col, &date, format)
                        .map_err(xlsx_error)?;
                    return Ok(());
                }
                worksheet
                    .write_with_format(row, col, value.as_str(), format)
                    .map_err(xlsx_error)?;
                return Ok(());
            }

            worksheet
                .write_blank(row, col, format)
                .map_err(xlsx_error)?;
            Ok(())
        }
        invalid => Err(invalid_cell_type(format!(
            "unsupported xlsx cell type: {}",
            invalid
        ))),
    }
}

fn apply_label_conditional_formats(
    worksheet: &mut Worksheet,
    sheet: &SheetPayload,
    first_data_row: u32,
    theme: &crate::payload::ThemePayload,
) -> Result<()> {
    if sheet.rows.is_empty() {
        return Ok(());
    }

    let last_row = first_data_row + sheet.rows.len() as u32 - 1;

    for (col_index, column) in sheet.columns.iter().enumerate() {
        let col_index = col_index as u16;
        let col_name = excel_column_name(col_index);
        let excel_row = first_data_row + 1;

        for label in &column.success_label_list {
            let success_format = style_by_name("success", theme)
                .ok_or_else(|| invalid_conditional_style("missing theme style: success"))?;
            let rule = format!(
                "=ISNUMBER(SEARCH(\"{}\",{}{}))",
                escape_excel_text(label),
                col_name,
                excel_row
            );
            let conditional_format = ConditionalFormatFormula::new()
                .set_rule(rule.as_str())
                .set_format(&success_format);

            worksheet
                .add_conditional_format(
                    first_data_row,
                    col_index,
                    last_row,
                    col_index,
                    &conditional_format,
                )
                .map_err(xlsx_error)?;
        }

        for label in &column.danger_label_list {
            let danger_format = style_by_name("danger", theme)
                .ok_or_else(|| invalid_conditional_style("missing theme style: danger"))?;
            let rule = format!(
                "=ISNUMBER(SEARCH(\"{}\",{}{}))",
                escape_excel_text(label),
                col_name,
                excel_row
            );
            let conditional_format = ConditionalFormatFormula::new()
                .set_rule(rule.as_str())
                .set_format(&danger_format);

            worksheet
                .add_conditional_format(
                    first_data_row,
                    col_index,
                    last_row,
                    col_index,
                    &conditional_format,
                )
                .map_err(xlsx_error)?;
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::process::Command;
    use std::time::{SystemTime, UNIX_EPOCH};

    use serde_json::json;

    use super::write_workbook;
    use crate::payload::RenderWorkbookPayload;

    fn render_zip(payload: RenderWorkbookPayload) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        write_workbook(&payload, &mut output).unwrap();
        output.into_inner()
    }

    fn write_temp_file(bytes: &[u8]) -> PathBuf {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!("xlsx-renderer-test-{}.xlsx", unique));
        fs::write(&path, bytes).unwrap();
        path
    }

    fn read_zip_entry(path: &std::path::Path, entry: &str) -> String {
        let output = Command::new("unzip")
            .args(["-p", path.to_str().unwrap(), entry])
            .output()
            .unwrap();

        assert!(
            output.status.success(),
            "Impossible de lire {} via unzip: {}",
            entry,
            String::from_utf8_lossy(&output.stderr)
        );

        String::from_utf8(output.stdout).unwrap()
    }

    fn base_payload(rows: serde_json::Value) -> RenderWorkbookPayload {
        serde_json::from_value(json!({
            "theme": {
                "styles": {
                    "header": { "bold": true, "fontColor": "ffffff", "backgroundColor": "017074", "border": "thin", "verticalAlign": "middle", "textWrap": true },
                    "cell": { "border": "thin" },
                    "currency": { "border": "thin", "numFormat": "#,##0.00 [$EUR];-#,##0.00 [$EUR];-" },
                    "success": { "backgroundColor": "85ca8b" },
                    "danger": { "backgroundColor": "f8ceb7" },
                    "title": { "bold": true }
                }
            },
            "workbook": {
                "sheetList": [{
                    "sheetName": "Export",
                    "columns": [
                        { "key": "name", "title": "Name", "width": 20 },
                        { "key": "amount", "title": "Amount", "width": 12, "dataStyle": "currency", "successLabelList": ["OK"], "dangerLabelList": ["KO"] },
                        { "key": "link", "title": "Link", "width": 18 }
                    ],
                    "rows": rows
                }]
            },
            "target": { "kind": "buffer" },
            "memoryMode": "constant-memory"
        }))
        .unwrap()
    }

    #[test]
    fn writes_basic_workbook_zip_entries() {
        let payload = base_payload(json!([
            { "values": { "name": "Alice", "amount": 12.5, "link": "" } }
        ]));

        let bytes = render_zip(payload);
        assert!(bytes.starts_with(b"PK"));

        let path = write_temp_file(&bytes);
        let workbook_xml = read_zip_entry(&path, "xl/workbook.xml");
        let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
        fs::remove_file(&path).unwrap();

        assert!(workbook_xml.contains("Export"));
        assert!(sheet_xml.contains("<sheetData>"));
        assert!(sheet_xml.contains("Alice"));
    }

    #[test]
    fn writes_merge_formula_link_and_conditional_formatting() {
        let payload: RenderWorkbookPayload = serde_json::from_value(json!({
            "theme": {
                "styles": {
                    "header": { "bold": true, "fontColor": "ffffff", "backgroundColor": "017074", "border": "thin", "verticalAlign": "middle", "textWrap": true },
                    "cell": { "border": "thin" },
                    "success": { "backgroundColor": "85ca8b" },
                    "danger": { "backgroundColor": "f8ceb7" },
                    "title": { "bold": true }
                }
            },
            "workbook": {
                "sheetList": [{
                    "sheetName": "Checks",
                    "columns": [
                        { "key": "title", "title": "Title", "width": 20 },
                        { "key": "calc", "title": "Calc", "width": 12 },
                        { "key": "url", "title": "Url", "width": 20, "successLabelList": ["ready"], "dangerLabelList": ["blocked"] }
                    ],
                    "mergeRangeList": [
                        { "startRow": 3, "startCol": 0, "endRow": 3, "endCol": 1, "value": "Merged", "styleList": ["title"] }
                    ],
                    "conditionalFormatList": [
                        { "startRow": 1, "startCol": 2, "endRow": 2, "endCol": 2, "formula": "=SEARCH(\"ready\",C2)", "style": "success" }
                    ],
                    "rows": [
                        { "values": {
                            "title": "row-1",
                            "calc": { "type": "formula", "formula": "=1+1" },
                            "url": { "type": "link", "url": "https://example.com", "label": "example", "styleList": ["cell"] }
                        }},
                        { "values": {
                            "title": "row-2",
                            "calc": { "type": "date", "value": "2024-01-15" },
                            "url": "ready"
                        }}
                    ]
                }]
            },
            "target": { "kind": "buffer" },
            "memoryMode": "standard"
        }))
        .unwrap();

        let bytes = render_zip(payload);
        let path = write_temp_file(&bytes);
        let sheet_xml = read_zip_entry(&path, "xl/worksheets/sheet1.xml");
        let rels_xml = read_zip_entry(&path, "xl/worksheets/_rels/sheet1.xml.rels");
        fs::remove_file(&path).unwrap();

        assert!(sheet_xml.contains("<mergeCells"));
        assert!(sheet_xml.contains("<f>1+1</f>"));
        assert!(sheet_xml.contains("<conditionalFormatting"));
        assert!(sheet_xml.contains("SEARCH("));
        assert!(sheet_xml.contains("ready"));
        assert!(rels_xml.contains("https://example.com"));
    }
}
