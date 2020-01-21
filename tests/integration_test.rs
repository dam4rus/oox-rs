extern crate oox;

use oox::docx::package::Package;
use std::path::PathBuf;

#[test]
#[ignore]
fn test_package_load() {
    //simple_logger::init().unwrap();

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = manifest_dir.join("tests/sample.docx");

    let package = Package::from_file(&sample_docx_file).unwrap();
    assert!(package.app_info.is_some());
    assert!(package.core.is_some());
    assert!(package.main_document.is_some());
    assert_eq!(package.main_document_relationships.len(), 14);
    assert!(package.styles.is_some());
    assert!(package.footnotes.is_some());
    assert!(package.numbering.is_some());
    assert!(package.settings.is_some());
    assert_eq!(package.medias.len(), 4);
    assert_eq!(package.themes.len(), 1);
    package.themes.get("theme1").unwrap();
}

#[test]
#[ignore]
fn test_package_resolve_default_style() {
    //simple_logger::init().unwrap();

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = manifest_dir.join("tests/sample.docx");

    let package = Package::from_file(&sample_docx_file).unwrap();
    /*let def_style = */
    package.resolve_document_default_style().unwrap();
    // TODO(kalmar.robert) Write real unit test
    //println!("{:?}", def_style);
}

#[test]
#[ignore]
fn test_package_get_main_document_theme() {
    //simple_logger::init().unwrap();

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = manifest_dir.join("tests/sample.docx");

    let package = Package::from_file(&sample_docx_file).unwrap();
    /*let theme = */
    package.get_main_document_theme().unwrap();
    // TODO(kalmar.robert) Write real unit test
    //println!("{:?}", theme);
}
