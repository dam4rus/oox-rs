use super::{
    core::{NonVisualDrawingProps, NonVisualPictureProperties, ShapeProperties},
    shapeprops::BlipFillProperties,
};
use crate::{error::MissingChildNodeError, xml::XmlNode};

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Debug, Clone, PartialEq)]
pub struct PictureNonVisual {
    pub non_visual_drawing_props: NonVisualDrawingProps,
    pub non_visual_picture_props: NonVisualPictureProperties,
}

impl PictureNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_drawing_props = None;
        let mut non_visual_picture_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => non_visual_drawing_props = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvPicPr" => {
                    non_visual_picture_props = Some(NonVisualPictureProperties::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let non_visual_drawing_props =
            non_visual_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;

        let non_visual_picture_props =
            non_visual_picture_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPicPr"))?;

        Ok(Self {
            non_visual_drawing_props,
            non_visual_picture_props,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Picture {
    pub non_visual_props: PictureNonVisual,
    pub blip_fill_props: BlipFillProperties,
    pub shape_props: ShapeProperties,
}

impl Picture {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut blip_fill_props = None;
        let mut shape_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvPicPr" => non_visual_props = Some(PictureNonVisual::from_xml_element(child_node)?),
                "blipFill" => blip_fill_props = Some(BlipFillProperties::from_xml_element(child_node)?),
                "spPr" => shape_props = Some(ShapeProperties::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPicPr"))?;

        let blip_fill_props =
            blip_fill_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "blipFill"))?;

        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;

        Ok(Self {
            non_visual_props,
            blip_fill_props,
            shape_props,
        })
    }
}
