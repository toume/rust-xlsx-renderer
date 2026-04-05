use rust_xlsxwriter::ExcelDateTime;

pub fn extract_string(value: Option<&serde_json::Value>) -> Option<String> {
    match value {
        Some(serde_json::Value::String(value)) => Some(value.clone()),
        Some(serde_json::Value::Number(value)) => Some(value.to_string()),
        Some(serde_json::Value::Bool(value)) => Some(value.to_string()),
        _ => None,
    }
}

pub fn extract_number(value: Option<&serde_json::Value>) -> Option<f64> {
    match value {
        Some(serde_json::Value::Number(value)) => value.as_f64(),
        Some(serde_json::Value::String(value)) => value.parse::<f64>().ok(),
        _ => None,
    }
}

pub fn parse_iso_date(value: &str) -> Option<ExcelDateTime> {
    let date = value.get(0..10)?;
    let mut parts = date.split('-');
    let year = parts.next()?.parse::<u16>().ok()?;
    let month = parts.next()?.parse::<u8>().ok()?;
    let day = parts.next()?.parse::<u8>().ok()?;
    ExcelDateTime::from_ymd(year, month, day).ok()
}

pub fn excel_column_name(mut column: u16) -> String {
    let mut name = String::new();

    loop {
        let remainder = column % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        if column < 26 {
            break;
        }
        column = (column / 26) - 1;
    }

    name
}

pub fn escape_excel_text(value: &str) -> String {
    value.replace('"', "\"\"")
}

#[cfg(test)]
mod tests {
    use super::{
        escape_excel_text, excel_column_name, extract_number, extract_string, parse_iso_date,
    };
    use serde_json::json;

    #[test]
    fn computes_excel_column_names() {
        assert_eq!(excel_column_name(0), "A");
        assert_eq!(excel_column_name(25), "Z");
        assert_eq!(excel_column_name(26), "AA");
        assert_eq!(excel_column_name(27), "AB");
        assert_eq!(excel_column_name(701), "ZZ");
        assert_eq!(excel_column_name(702), "AAA");
    }

    #[test]
    fn escapes_excel_quotes() {
        assert_eq!(escape_excel_text("foo"), "foo");
        assert_eq!(escape_excel_text("a\"b"), "a\"\"b");
    }

    #[test]
    fn extracts_string_and_number_values() {
        assert_eq!(extract_string(Some(&json!("abc"))), Some("abc".to_string()));
        assert_eq!(extract_string(Some(&json!(12.5))), Some("12.5".to_string()));
        assert_eq!(extract_number(Some(&json!(12.5))), Some(12.5));
        assert_eq!(extract_number(Some(&json!("12.5"))), Some(12.5));
        assert_eq!(extract_number(Some(&json!("nope"))), None);
    }

    #[test]
    fn parses_iso_dates_from_prefix() {
        assert!(parse_iso_date("2024-02-29").is_some());
        assert!(parse_iso_date("2024-02-29T12:30:00Z").is_some());
        assert!(parse_iso_date("2024/02/29").is_none());
    }
}
