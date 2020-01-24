#![forbid(unsafe_code)]

#[cfg(any(test, feature = "docx"))]
pub mod docx;
pub mod error;
#[cfg(any(test, feature = "pptx"))]
pub mod pptx;
pub mod shared;
pub mod update;
pub mod xml;
pub mod xsdtypes;

extern crate strum;
#[macro_use]
extern crate strum_macros;
