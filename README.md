# oox-rs

A library to deserialize Open Office Xml files in Rust.

[![Latest version](https://img.shields.io/crates/v/oox.svg)](https://crates.io/crates/oox)
[![Documentation](https://docs.rs/oox/badge.svg)](https://docs.rs/oox)

## Overview

oox-rs is a low level deserializer for Microsoft's OfficeOpen XML file formats. It's still WIP, so expect API breaking changes.

The Office Open XML file formats are described by the [ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm).
The types represented in this library are generated from the Transitional XML Schema's, which is described in
[ECMA-376 4th edition Part 4](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%204%20-%20Transitional%20Migration%20Features.zip).

Documentation is generated from the "Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf" file, found in [ECMA-376 4th edition Part 1](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip)

## Usage

Oox-rs has no high level interface to access objects and their attributes in a page/slide yet. It only provides access
to the elements in various xml files as in-memory objects. If you want to import a document into your application you
should look up the documentations above. A high level interface is planned.

### Load a docx file

```Rust
use oox::docx::package::Package as DocxPackage;

let package = DocxPackage::from_file(&PathBuf::from("path/to/example/file.docx")).unwrap();
let main_document = package.main_document.as_ref().unwrap(); // Access the main document
```