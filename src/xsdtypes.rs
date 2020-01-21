use super::{error::NotGroupMemberError, xml::XmlNode};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub trait XsdType
where
    Self: Sized,
{
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self>;
}

pub trait XsdChoice: XsdType {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool;

    /// Tries to parse an XmlNode as a choice member.
    /// None is returned if the XmlNode is not a member of the choice element (implementors should return a boxed
    /// NotGroupMemberError), otherwise Some is returned with the Result of from_xml_element.
    fn try_from_xml_element(xml_node: &XmlNode) -> Option<Result<Self>> {
        match Self::from_xml_element(xml_node) {
            Ok(val) => Some(Ok(val)),
            Err(err) => match err.downcast::<NotGroupMemberError>() {
                Ok(_) => None,
                Err(err) => Some(Err(err)),
            },
        }
    }
}
