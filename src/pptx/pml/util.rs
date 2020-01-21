use crate::{error::MissingAttributeError, xml::XmlNode};

pub(crate) trait XmlNodeExt {
    // It's a common pattern throughout the OpenOffice XML file format that a simple type is wrapped in a complex type
    // with a single attribute called `val`. This is a small wrapper function to reduce the boiler plate for such
    // complex types
    fn get_val_attribute(&self) -> std::result::Result<&String, MissingAttributeError>;
}

impl XmlNodeExt for XmlNode {
    fn get_val_attribute(&self) -> std::result::Result<&String, MissingAttributeError> {
        self.attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(self.name.clone(), "val"))
    }
}
