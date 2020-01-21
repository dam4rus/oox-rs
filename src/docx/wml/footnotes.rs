use super::{document::BlockLevelElts, simpletypes::DecimalNumber};
use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError},
    xml::XmlNode,
    xsdtypes::XsdChoice,
};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum FtnEdnType {
    #[strum(serialize = "normal")]
    Normal,
    #[strum(serialize = "separator")]
    Separator,
    #[strum(serialize = "continuationSeparator")]
    ContinuationSeparator,
    #[strum(serialize = "continuationNotice")]
    ContinuationNotice,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FtnEdn {
    pub ftn_edn_type: Option<FtnEdnType>,
    pub id: DecimalNumber,
    pub block_level_elements: Vec<BlockLevelElts>,
}

impl FtnEdn {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut ftn_edn_type = None;
        let mut id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => ftn_edn_type = Some(value.parse()?),
                "w:id" => id = Some(value.parse()?),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:id"))?;

        let block_level_elements = xml_node
            .child_nodes
            .iter()
            .filter_map(BlockLevelElts::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        if !block_level_elements.is_empty() {
            Ok(Self {
                ftn_edn_type,
                id,
                block_level_elements,
            })
        } else {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "BlockLevelElts",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Footnotes(pub Vec<FtnEdn>);

impl Footnotes {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let footnotes = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "footnote")
            .map(FtnEdn::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(footnotes))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::document::{ContentBlockContent, P};
    use std::str::FromStr;

    impl Footnotes {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                FtnEdn::test_xml("w:footnote"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![FtnEdn::test_instance()])
        }
    }

    #[test]
    pub fn test_footnotes_from_xml() {
        let xml = Footnotes::test_xml("w:footnotes");
        assert_eq!(
            Footnotes::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Footnotes::test_instance(),
        );
    }

    impl FtnEdn {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="normal" w:id="1">
                {}
            </{node_name}>"#,
                P::test_xml("w:p"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ftn_edn_type: Some(FtnEdnType::Normal),
                id: 1,
                block_level_elements: vec![BlockLevelElts::Chunk(ContentBlockContent::Paragraph(Box::new(
                    P::test_instance(),
                )))],
            }
        }
    }

    #[test]
    pub fn test_ftn_edn_from_xml() {
        let xml = FtnEdn::test_xml("w:ftnEdn");
        assert_eq!(
            FtnEdn::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FtnEdn::test_instance(),
        );
    }
}
