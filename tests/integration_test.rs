#![cfg(feature = "docx")]
extern crate oox;

use oox::{
    docx::package::Package as DocxPackage,
    pptx::package::Package as PptxPackage,
    shared::drawingml::coordsys::{Point2D, PositiveSize2D},
};
use std::path::PathBuf;

#[test]
fn test_docx_package_load() {
    //simple_logger::init().unwrap();

    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = manifest_dir.join("tests/sample.docx");

    let package = DocxPackage::from_file(&sample_docx_file).unwrap();

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
fn test_pptx_package_load() {
    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_pptx_file = manifest_dir.join("tests/sample.pptx");

    let document = PptxPackage::from_file(&sample_pptx_file).unwrap();
    let mut slides = document.slides();
    {
        let first_slide = slides.next().unwrap();
        let sptree = &first_slide.common_slide_data.shape_tree;
        assert_eq!(sptree.non_visual_props.drawing_props.id, 1);
        let transform = sptree.group_shape_props.transform.as_ref().unwrap();
        assert_eq!(*transform.offset.as_ref().unwrap(), Point2D::new(0, 0));
        assert_eq!(*transform.extents.as_ref().unwrap(), PositiveSize2D::new(0, 0));
        assert_eq!(*transform.child_offset.as_ref().unwrap(), Point2D::new(0, 0));
        assert_eq!(*transform.child_extents.as_ref().unwrap(), PositiveSize2D::new(0, 0));
        assert_eq!(sptree.shape_array.len(), 2);
    }

    {
        let second_slide = slides.next().unwrap();
        let sptree = &second_slide.common_slide_data.shape_tree;
        assert_eq!(sptree.non_visual_props.drawing_props.id, 1);
        let transform = sptree.group_shape_props.transform.as_ref().unwrap();
        assert_eq!(*transform.offset.as_ref().unwrap(), Point2D::new(0, 0));
        assert_eq!(*transform.extents.as_ref().unwrap(), PositiveSize2D::new(0, 0));
        assert_eq!(*transform.child_offset.as_ref().unwrap(), Point2D::new(0, 0));
        assert_eq!(*transform.child_extents.as_ref().unwrap(), PositiveSize2D::new(0, 0));
        assert_eq!(sptree.shape_array.len(), 2);
    }

    assert_eq!(slides.next().is_none(), true);
}
