//! OPC (Open Packaging Conventions) support for OOXML files.
//!
//! All OOXML formats (.docx, .pptx, .xlsx) are ZIP archives containing XML parts.
//! This crate provides shared utilities for reading and writing those packages.

pub mod content_types;
pub mod package;
pub mod relationship;
pub mod xml;
pub mod xmlio;

pub use content_types::{ContentTypes, DefaultType, OverrideType};
pub use package::OpcPackage;
pub use relationship::{Relationship, Relationships};
