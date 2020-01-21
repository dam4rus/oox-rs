use crate::error::MissingAttributeError;
use crate::xml::XmlNode;
use std::{io::Read, str::FromStr};
use zip::read::ZipFile;

pub type RelationshipId = String;

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

pub const THEME_RELATION_TYPE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

impl Relationship {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut rel_type = None;
        let mut target = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "Id" => id = Some(value.clone()),
                "Type" => rel_type = Some(value.clone()),
                "Target" => target = Some(value.clone()),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "Id"))?;
        let rel_type = rel_type.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "Type"))?;
        let target = target.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "Target"))?;

        Ok(Self { id, rel_type, target })
    }
}

pub fn relationships_from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Vec<Relationship>> {
    let mut xml_string = String::new();
    zip_file.read_to_string(&mut xml_string)?;
    let xml_node = XmlNode::from_str(xml_string.as_str())?;
    let mut relationships = Vec::new();

    for child_node in &xml_node.child_nodes {
        relationships.push(Relationship::from_xml_element(child_node)?);
    }

    Ok(relationships)
}
