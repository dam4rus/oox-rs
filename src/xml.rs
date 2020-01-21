use crate::error::{InvalidXmlError, ParseBoolError};
use quick_xml::{
    events::{BytesStart, Event},
    Reader,
};
use std::{
    collections::HashMap,
    fmt::{Display, Formatter},
    io::Read,
    str::FromStr,
};
use zip::read::ZipFile;

/// Represents an implementation independent xml node
#[derive(Debug, Clone, PartialEq)]
pub struct XmlNode {
    pub name: String,
    pub child_nodes: Vec<XmlNode>,
    pub attributes: HashMap<String, String>,
    pub text: Option<String>,
}

impl Display for XmlNode {
    fn fmt(&self, f: &mut Formatter<'_>) -> ::std::fmt::Result {
        write!(f, "name: {}", self.name)
    }
}

impl XmlNode {
    pub fn new<T: Into<String>>(name: T) -> Self {
        Self {
            name: name.into(),
            child_nodes: Vec::new(),
            attributes: HashMap::new(),
            text: None,
        }
    }

    pub fn local_name(&self) -> &str {
        match self.name.find(':') {
            Some(idx) => self.name.split_at(idx + 1).1,
            None => self.name.as_str(),
        }
    }

    fn from_quick_xml_element(xml_element: &BytesStart<'_>) -> Result<Self, ::std::str::Utf8Error> {
        let name = ::std::str::from_utf8(xml_element.name())?;
        let mut node = Self::new(name);

        for attr in xml_element.attributes() {
            if let Ok(a) = attr {
                let key_str = ::std::str::from_utf8(&a.key)?;
                let value_str = ::std::str::from_utf8(&a.value)?;
                node.attributes.insert(String::from(key_str), String::from(value_str));
            }
        }

        Ok(node)
    }

    fn parse_child_elements(
        xml_node: &mut Self,
        xml_element: &BytesStart<'_>,
        xml_reader: &mut Reader<&[u8]>,
    ) -> Result<Vec<Self>, ::std::str::Utf8Error> {
        let mut child_nodes = Vec::new();

        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(Event::Start(ref element)) => {
                    let mut node = Self::from_quick_xml_element(element)?;
                    node.child_nodes = Self::parse_child_elements(&mut node, element, xml_reader)?;
                    child_nodes.push(node);
                }
                Ok(Event::Text(text)) => {
                    xml_node.text = text.unescape_and_decode(xml_reader).ok();
                }
                Ok(Event::Empty(ref element)) => {
                    let node = Self::from_quick_xml_element(element)?;
                    child_nodes.push(node);
                }
                Ok(Event::End(ref element)) => {
                    if element.name() == xml_element.name() {
                        break;
                    }
                }
                Ok(Event::Eof) => {
                    break;
                }
                _ => (),
            }

            buffer.clear();
        }

        Ok(child_nodes)
    }
}

impl FromStr for XmlNode {
    type Err = InvalidXmlError;

    fn from_str(xml_string: &str) -> Result<Self, Self::Err> {
        let mut xml_reader = Reader::from_str(xml_string.as_ref());
        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(Event::Start(ref element)) => {
                    let mut root_node = Self::from_quick_xml_element(element).map_err(|_| InvalidXmlError {})?;
                    root_node.child_nodes = Self::parse_child_elements(&mut root_node, element, &mut xml_reader)
                        .map_err(|_| InvalidXmlError {})?;
                    return Ok(root_node);
                }
                Ok(Event::Eof) => break,
                _ => (),
            }

            buffer.clear();
        }

        Err(InvalidXmlError {})
    }
}

pub fn parse_xml_bool<T: AsRef<str>>(value: T) -> Result<bool, ParseBoolError> {
    match value.as_ref() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(ParseBoolError::new(String::from(value.as_ref()))),
    }
}

pub fn zip_file_to_xml_node(zip_file: &mut ZipFile) -> Result<XmlNode, Box<dyn std::error::Error>> {
    let mut xml_string = String::new();
    zip_file.read_to_string(&mut xml_string)?;
    XmlNode::from_str(xml_string.as_str()).map_err(Into::into)
}

#[cfg(test)]
mod tests {
    use super::XmlNode;
    use std::str::FromStr;

    #[test]
    fn test_xml_parser() {
        use std::fs::File;
        use std::io::Read;
        use std::path::PathBuf;

        let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        let sample_xml_file = test_dir.join("tests/presentation.xml");
        let mut file = File::open(sample_xml_file).expect("Sample xml file not found");

        let mut file_content = String::new();
        file.read_to_string(&mut file_content)
            .expect("Failed to read sample xml file to string");

        let root_node = XmlNode::from_str(file_content.as_str()).expect("Couldn't create XmlNode from string");
        assert_eq!(root_node.name, "p:presentation");
        assert_eq!(
            root_node.attributes.get("xmlns:a").unwrap(),
            "http://schemas.openxmlformats.org/drawingml/2006/main"
        );

        assert_eq!(root_node.child_nodes[0].name, "p:sldMasterIdLst");
        assert_eq!(root_node.child_nodes[1].name, "p:sldIdLst");
        assert_eq!(root_node.child_nodes[2].name, "p:sldSz");
        assert_eq!(root_node.child_nodes[3].name, "p:notesSz");
        assert_eq!(root_node.child_nodes[4].name, "p:custDataLst");
        assert_eq!(root_node.child_nodes[5].name, "p:defaultTextStyle");
        assert_eq!(root_node.child_nodes[0].child_nodes[0].name, "p:sldMasterId");

        let slide_id_0_node = &root_node.child_nodes[1].child_nodes[0];
        assert_eq!(slide_id_0_node.name, "p:sldId");
        assert_eq!(slide_id_0_node.attributes.get("id").unwrap(), "256");
        assert_eq!(slide_id_0_node.attributes.get("r:id").unwrap(), "rId2");

        assert_eq!(root_node.child_nodes[1].child_nodes[1].name, "p:sldId");

        let lvl1_ppr_defrpr_node = &root_node.child_nodes[5].child_nodes[1].child_nodes[0];
        assert_eq!(lvl1_ppr_defrpr_node.attributes.get("sz").unwrap(), "1800");
        assert_eq!(lvl1_ppr_defrpr_node.attributes.get("kern").unwrap(), "1200");
    }
}
