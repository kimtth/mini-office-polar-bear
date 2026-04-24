//! Internal document models — the canonical editing representation.
//! OOXML crates import/export to these models rather than being the live editing format.

pub mod calc;
pub mod slides;
pub mod writer;

use crate::PboResult;

/// Trait for documents that can be loaded from and saved to bytes.
pub trait Document: Sized {
    fn from_bytes(data: &[u8]) -> PboResult<Self>;
    fn to_bytes(&self) -> PboResult<Vec<u8>>;
}
