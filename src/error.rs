use napi::{Error, Status};
use rust_xlsxwriter::XlsxError;
use serde::Serialize;

#[derive(Clone, Copy)]
pub enum XlsxRendererErrorCode {
    InvalidPayload,
    InvalidMemoryMode,
    InvalidTarget,
    InvalidConditionalStyle,
    InvalidCell,
    InvalidCellType,
    InvalidLink,
    InvalidNumber,
    Io,
    XlsxWriteFailed,
}

impl XlsxRendererErrorCode {
    fn as_str(self) -> &'static str {
        match self {
            Self::InvalidPayload => "INVALID_PAYLOAD",
            Self::InvalidMemoryMode => "INVALID_MEMORY_MODE",
            Self::InvalidTarget => "INVALID_TARGET",
            Self::InvalidConditionalStyle => "INVALID_CONDITIONAL_STYLE",
            Self::InvalidCell => "INVALID_CELL",
            Self::InvalidCellType => "INVALID_CELL_TYPE",
            Self::InvalidLink => "INVALID_LINK",
            Self::InvalidNumber => "INVALID_NUMBER",
            Self::Io => "IO_ERROR",
            Self::XlsxWriteFailed => "XLSX_WRITE_FAILED",
        }
    }
}

#[derive(Serialize)]
struct NativeErrorPayload<'a> {
    code: &'a str,
    message: &'a str,
}

pub fn native_error(
    status: Status,
    code: XlsxRendererErrorCode,
    message: impl Into<String>,
) -> Error {
    let message = message.into();
    let payload = NativeErrorPayload {
        code: code.as_str(),
        message: &message,
    };

    let serialized = serde_json::to_string(&payload)
        .unwrap_or_else(|_| format!(r#"{{"code":"{}","message":"{}"}}"#, code.as_str(), message));

    Error::new(status, serialized)
}

pub fn invalid_payload(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidPayload,
        message,
    )
}

pub fn invalid_memory_mode(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidMemoryMode,
        message,
    )
}

pub fn invalid_target(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidTarget,
        message,
    )
}

pub fn invalid_conditional_style(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidConditionalStyle,
        message,
    )
}

pub fn invalid_cell(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidCell,
        message,
    )
}

pub fn invalid_cell_type(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidCellType,
        message,
    )
}

pub fn invalid_link(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidLink,
        message,
    )
}

pub fn invalid_number(message: impl Into<String>) -> Error {
    native_error(
        Status::InvalidArg,
        XlsxRendererErrorCode::InvalidNumber,
        message,
    )
}

pub fn xlsx_error(error: XlsxError) -> Error {
    native_error(
        Status::GenericFailure,
        XlsxRendererErrorCode::XlsxWriteFailed,
        format!("xlsx writer failed: {}", error),
    )
}

pub fn io_error(error: std::io::Error) -> Error {
    native_error(
        Status::GenericFailure,
        XlsxRendererErrorCode::Io,
        format!("io operation failed: {}", error),
    )
}
