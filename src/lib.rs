#![forbid(unsafe_code)]

pub mod error;
pub mod update;
pub mod xml;
pub mod xsdtypes;
pub mod shared;
#[cfg(any(test, feature = "pptx"))]
pub mod pptx;
#[cfg(any(test, feature = "docx"))]
pub mod docx;

extern crate strum;
#[macro_use]
extern crate strum_macros;
