mod error;
mod memory_mode;
mod payload;
mod styles;
mod utils;
mod writer;

use std::fs::{metadata, symlink_metadata};
use std::io::Cursor;
use std::path::{Path, PathBuf};

use napi::bindgen_prelude::{AsyncTask, Buffer};
use napi::{Env, Result, Task};
use napi_derive::napi;
use tempfile::{Builder, NamedTempFile};

use crate::error::{invalid_payload, invalid_target, io_error};
use crate::payload::{RenderTargetPayload, RenderWorkbookPayload};
use crate::writer::write_workbook;

#[napi(object)]
pub struct RenderWorkbookFileResult {
    pub kind: String,
    pub path: String,
    pub bytes: u32,
}

pub struct RenderWorkbookToBufferTask {
    payload: RenderWorkbookPayload,
}

pub struct RenderWorkbookToFileTask {
    payload: RenderWorkbookPayload,
}

#[napi]
pub fn render_workbook_to_buffer(
    payload_json: String,
) -> Result<AsyncTask<RenderWorkbookToBufferTask>> {
    Ok(AsyncTask::new(RenderWorkbookToBufferTask {
        payload: parse_payload(&payload_json)?,
    }))
}

#[napi]
pub fn render_workbook_to_file(
    payload_json: String,
) -> Result<AsyncTask<RenderWorkbookToFileTask>> {
    Ok(AsyncTask::new(RenderWorkbookToFileTask {
        payload: parse_payload(&payload_json)?,
    }))
}

impl Task for RenderWorkbookToBufferTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        let mut output = Cursor::new(Vec::new());
        write_workbook(&self.payload, &mut output)?;
        Ok(output.into_inner())
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output.into())
    }
}

impl Task for RenderWorkbookToFileTask {
    type Output = RenderWorkbookFileResult;
    type JsValue = RenderWorkbookFileResult;

    fn compute(&mut self) -> Result<Self::Output> {
        let path = match &self.payload.target {
            RenderTargetPayload::File { path } => path.clone(),
            RenderTargetPayload::Buffer => {
                return Err(invalid_target(
                    "file rendering requires target.kind to be 'file'",
                ))
            }
        };

        let target_path = PathBuf::from(&path);
        let mut file = create_output_file(&target_path)?;
        write_workbook(&self.payload, &mut file)?;
        file.as_file_mut().sync_all().map_err(io_error)?;
        file.persist(&target_path)
            .map_err(|error| io_error(error.error))?;
        let file_size = metadata(&target_path).map_err(io_error)?.len();

        Ok(RenderWorkbookFileResult {
            kind: "file".to_string(),
            path,
            bytes: u32::try_from(file_size).unwrap_or(u32::MAX),
        })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

fn create_output_file(target_path: &Path) -> Result<NamedTempFile> {
    if target_path.file_name().is_none() {
        return Err(invalid_target("file target path must include a file name"));
    }

    if let Ok(metadata) = symlink_metadata(target_path) {
        if metadata.file_type().is_symlink() {
            return Err(invalid_target(
                "file target path must not be a symbolic link",
            ));
        }
    }

    let parent = target_path.parent().unwrap_or_else(|| Path::new("."));

    if let Ok(metadata) = symlink_metadata(parent) {
        if metadata.file_type().is_symlink() {
            return Err(invalid_target(
                "file target parent directory must not be a symbolic link",
            ));
        }

        if !metadata.is_dir() {
            return Err(invalid_target("file target parent must be a directory"));
        }
    } else {
        return Err(invalid_target(
            "file target parent directory does not exist",
        ));
    }

    let file_name = target_path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("workbook.xlsx");

    Builder::new()
        .prefix(&format!(".{}.", file_name))
        .suffix(".tmp")
        .tempfile_in(parent)
        .map_err(io_error)
}

fn parse_payload(payload_json: &str) -> Result<RenderWorkbookPayload> {
    serde_json::from_str(payload_json)
        .map_err(|error| invalid_payload(format!("invalid xlsx payload: {}", error)))
}

#[cfg(test)]
mod tests {
    use std::fs;

    use serde_json::json;

    use super::{create_output_file, parse_payload, RenderWorkbookToFileTask};
    use crate::payload::RenderWorkbookPayload;
    use napi::Task;

    #[test]
    fn parses_buffer_payload() {
        let payload = parse_payload(
            r#"{"theme":{"styles":{}},"workbook":{"sheetList":[{"sheetName":"A","columns":[],"rows":[]}]},"target":{"kind":"buffer"}}"#,
        )
        .unwrap();

        assert_eq!(payload.workbook.sheet_list.len(), 1);
    }

    #[test]
    fn rejects_invalid_json() {
        let error = parse_payload("{").unwrap_err();
        assert!(error.to_string().contains("INVALID_PAYLOAD"));
    }

    #[test]
    fn rejects_missing_parent_directory_for_file_targets() {
        let payload: RenderWorkbookPayload = serde_json::from_value(json!({
            "theme": { "styles": {} },
            "workbook": {
                "sheetList": [{
                    "sheetName": "A",
                    "columns": [{ "key": "name", "title": "Name", "width": 12 }],
                    "rows": [{ "values": { "name": "Alice" } }]
                }]
            },
            "target": {
                "kind": "file",
                "path": "./definitely-missing-dir/output.xlsx"
            }
        }))
        .unwrap();

        let mut task = RenderWorkbookToFileTask { payload };
        let error = task.compute().err().unwrap();

        assert!(error
            .to_string()
            .contains("file target parent directory does not exist"));
    }

    #[cfg(unix)]
    #[test]
    fn rejects_symbolic_link_file_targets() {
        use std::os::unix::fs::symlink;

        let temp_dir = tempfile::tempdir().unwrap();
        let target_file = temp_dir.path().join("real.xlsx");
        fs::write(&target_file, b"old").unwrap();
        let symlink_path = temp_dir.path().join("link.xlsx");
        symlink(&target_file, &symlink_path).unwrap();

        let error = create_output_file(&symlink_path).unwrap_err();

        assert!(error
            .to_string()
            .contains("file target path must not be a symbolic link"));
    }
}
