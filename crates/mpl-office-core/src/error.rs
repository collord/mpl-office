//! Error types for the converter.

use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("SVG parse error: {0}")]
    Parse(String),

    #[error("SVG path parse error: {0}")]
    Path(String),

    #[error("invalid SVG: {0}")]
    InvalidSvg(String),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("XML attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
}

pub type Result<T> = std::result::Result<T, Error>;
