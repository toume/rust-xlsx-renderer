use napi::Result;

use crate::error::invalid_memory_mode;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MemoryMode {
    Standard,
    LowMemory,
    ConstantMemory,
}

impl MemoryMode {
    pub fn parse(value: Option<&str>) -> Result<Self> {
        match value.unwrap_or("constant-memory") {
            "constant-memory" => Ok(Self::ConstantMemory),
            "low-memory" => Ok(Self::LowMemory),
            "standard" => Ok(Self::Standard),
            invalid => Err(invalid_memory_mode(format!(
                "unsupported xlsx memory mode: {}",
                invalid
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::MemoryMode;

    #[test]
    fn default_is_constant_memory() {
        let mode = MemoryMode::parse(None).unwrap();
        assert_eq!(mode, MemoryMode::ConstantMemory);
    }

    #[test]
    fn parses_all_supported_values() {
        assert_eq!(
            MemoryMode::parse(Some("constant-memory")).unwrap(),
            MemoryMode::ConstantMemory
        );
        assert_eq!(
            MemoryMode::parse(Some("low-memory")).unwrap(),
            MemoryMode::LowMemory
        );
        assert_eq!(
            MemoryMode::parse(Some("standard")).unwrap(),
            MemoryMode::Standard
        );
    }

    #[test]
    fn rejects_unknown_values() {
        let error = MemoryMode::parse(Some("fast-and-loose")).unwrap_err();
        assert!(error.to_string().contains("INVALID_MEMORY_MODE"));
    }
}
