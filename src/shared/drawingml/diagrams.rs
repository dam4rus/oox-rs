use super::{
    core::LineProperties,
    shapeprops::{EffectProperties, FillProperties},
};
use crate::{
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct BackgroundFormatting {
    pub fill: Option<FillProperties>,
    pub effect: Option<EffectProperties>,
}

impl BackgroundFormatting {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                let node_name = child_node.local_name();
                if FillProperties::is_choice_member(node_name) {
                    instance.fill = Some(FillProperties::from_xml_element(child_node)?);
                } else if EffectProperties::is_choice_member(node_name) {
                    instance.effect = Some(EffectProperties::from_xml_element(child_node)?);
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WholeE2oFormatting {
    pub line: Option<LineProperties>,
    pub effect: Option<EffectProperties>,
}

impl WholeE2oFormatting {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "ln" => instance.line = Some(LineProperties::from_xml_element(child_node)?),
                    node_name if EffectProperties::is_choice_member(node_name) => {
                        instance.effect = Some(EffectProperties::from_xml_element(child_node)?)
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    impl BackgroundFormatting {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name}><noFill /></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                fill: Some(FillProperties::NoFill),
                effect: None,
            }
        }
    }

    #[test]
    pub fn test_background_formatting_from_xml() {
        let xml = BackgroundFormatting::test_xml("backgroundFormatting");
        assert_eq!(
            BackgroundFormatting::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            BackgroundFormatting::test_instance(),
        );
    }

    impl WholeE2oFormatting {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name}><ln /></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                line: Some(Default::default()),
                effect: None,
            }
        }
    }

    #[test]
    pub fn test_whole_e2o_formatting_from_xml() {
        let xml = WholeE2oFormatting::test_xml("wholeE2oFormatting");
        assert_eq!(
            WholeE2oFormatting::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WholeE2oFormatting::test_instance(),
        );
    }
}
