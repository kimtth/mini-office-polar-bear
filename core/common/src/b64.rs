//! Tiny base64 facade so downstream crates don't have to import `base64` directly.

use base64::Engine;
use base64::engine::general_purpose::STANDARD;

pub fn encode(bytes: &[u8]) -> String {
    STANDARD.encode(bytes)
}

pub fn decode(s: &str) -> Result<Vec<u8>, base64::DecodeError> {
    STANDARD.decode(s.as_bytes())
}
