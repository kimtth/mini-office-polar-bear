use thiserror::Error;

#[derive(Debug, Error)]
pub enum PboError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("XML parsing error: {0}")]
    Xml(String),

    #[error("ZIP archive error: {0}")]
    Zip(String),

    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),

    #[error("Invalid document structure: {0}")]
    InvalidStructure(String),
}

pub type PboResult<T> = Result<T, PboError>;
