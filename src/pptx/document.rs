use super::pml::{
    presentation::Presentation,
    slides::{Slide, SlideLayout, SlideMaster},
};
use log::info;
use crate::shared::{
    docprops::{AppInfo, Core},
    drawingml::sharedstylesheet::OfficeStyleSheet,
    relationship::{Relationship, relationships_from_zip_file},
};
use std::collections::HashMap;
use std::fs::File;
use std::path::{Path, PathBuf};
use zip::ZipArchive;

#[derive(Debug, Clone, PartialEq)]
pub struct PPTXDocument {
    pub file_path: PathBuf,
    pub app: Option<Box<AppInfo>>,
    pub core: Option<Box<Core>>,
    pub presentation: Option<Box<Presentation>>,
    pub theme_map: HashMap<PathBuf, Box<OfficeStyleSheet>>,
    pub slide_master_map: HashMap<PathBuf, Box<SlideMaster>>,
    pub slide_layout_map: HashMap<PathBuf, Box<SlideLayout>>,
    pub slide_map: HashMap<PathBuf, Box<Slide>>,
    pub slide_master_rels_map: HashMap<PathBuf, Vec<Relationship>>,
    pub slide_layout_rels_map: HashMap<PathBuf, Vec<Relationship>>,
    pub slide_rels_map: HashMap<PathBuf, Vec<Relationship>>,
    pub medias: Vec<PathBuf>,
}

impl PPTXDocument {
    pub fn from_file(pptx_path: &Path) -> Result<Self, Box<dyn (::std::error::Error)>> {
        let pptx_file = File::open(&pptx_path)?;
        let mut zipper = ZipArchive::new(&pptx_file)?;

        info!("parsing docProps/app.xml");
        let app = AppInfo::from_zip(&mut zipper).map(|val| val.into()).ok();
        info!("parsing docProps/core.xml");
        let core = Core::from_zip(&mut zipper).map(|val| val.into()).ok();
        info!("parsing ppt/presentation.xml");
        let presentation = Presentation::from_zip(&mut zipper).map(|val| val.into()).ok();
        let mut theme_map = HashMap::new();
        let mut slide_master_map = HashMap::new();
        let mut slide_layout_map = HashMap::new();
        let mut slide_map = HashMap::new();
        let mut slide_master_rels_map = HashMap::new();
        let mut slide_layout_rels_map = HashMap::new();
        let mut slide_rels_map = HashMap::new();
        let mut medias = Vec::new();

        for i in 0..zipper.len() {
            let mut zip_file = zipper.by_index(i)?;

            let file_path = PathBuf::from(zip_file.name());
            if file_path.starts_with("ppt/theme") {
                info!("parsing theme file: {}", zip_file.name());
                theme_map.insert(file_path, Box::new(OfficeStyleSheet::from_zip_file(&mut zip_file)?));
            } else if file_path.starts_with("ppt/slideMasters/_rels") {
                if file_path.extension().unwrap_or_default() != "rels" {
                    continue;
                }

                info!("parsing slide master relationship file: {}", zip_file.name());
                slide_master_rels_map.insert(
                    file_path,
                    relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slideMasters") {
                if file_path.extension().unwrap_or_default() != "xml" {
                    continue;
                }

                info!("parsing slide master file: {}", zip_file.name());
                slide_master_map.insert(file_path, Box::new(SlideMaster::from_zip_file(&mut zip_file)?));
            } else if file_path.starts_with("ppt/slideLayouts/_rels") {
                if file_path.extension().unwrap_or_default() != "rels" {
                    continue;
                }

                info!("parsing slide layout relationship file: {}", zip_file.name());
                slide_layout_rels_map.insert(
                    file_path,
                    relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slideLayouts") {
                if file_path.extension().unwrap_or_default() != "xml" {
                    continue;
                }

                info!("parsing slide layout file: {}", zip_file.name());
                slide_layout_map.insert(file_path, Box::new(SlideLayout::from_zip_file(&mut zip_file)?));
            } else if file_path.starts_with("ppt/slides/_rels") {
                if file_path.extension().unwrap_or_default() != "rels" {
                    continue;
                }

                info!("parsing slide relationship file: {}", zip_file.name());
                slide_rels_map.insert(
                    file_path,
                    relationships_from_zip_file(&mut zip_file)?,
                );
            } else if file_path.starts_with("ppt/slides") {
                if file_path.extension().unwrap_or_default() != "xml" {
                    continue;
                }

                info!("parsing slide file: {}", zip_file.name());
                slide_map.insert(file_path, Box::new(Slide::from_zip_file(&mut zip_file)?));
            } else if file_path.starts_with("ppt/media") {
                medias.push(file_path);
            }
        }

        Ok(Self {
            file_path: PathBuf::from(pptx_path),
            app,
            core,
            presentation,
            theme_map,
            slide_master_map,
            slide_layout_map,
            slide_map,
            slide_master_rels_map,
            slide_layout_rels_map,
            slide_rels_map,
            medias,
        })
    }

    pub fn slides(&self) -> Slides {
        Slides::new(&self.slide_map)
    }
}
#[derive(Debug, Clone)]
pub struct Slides<'a> {
    slide_map: &'a HashMap<PathBuf, Box<Slide>>,
    current_page_num: usize,
}

impl<'a> Slides<'a> {
    pub fn new(slide_map: &'a HashMap<PathBuf, Box<Slide>>) -> Self {
        Self {
            slide_map,
            current_page_num: 1,
        }
    }
}

impl<'a> Iterator for Slides<'a> {
    type Item = &'a Slide;

    fn next(&mut self) -> Option<Self::Item> {
        for i in self.current_page_num..=self.slide_map.len() {
            let opt_slide = self.slide_map.get(&PathBuf::from(format!("ppt/slides/slide{}.xml", i)));
            self.current_page_num += 1;
            if let Some(slide) = opt_slide {
                return Some(slide);
            }
        }

        None
    }
}

#[cfg(test)]
#[test]
fn test_sample_pptx() {
    use crate::shared::drawingml::coordsys::{Point2D, PositiveSize2D};

    let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_pptx_path = test_dir.join("tests/sample.pptx");

    let document = PPTXDocument::from_file(&sample_pptx_path).unwrap();
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
