use crate::{error::MissingAttributeError, xml::XmlNode};
use std::{error::Error, str::FromStr};

pub(crate) trait XmlNodeExt {
    // It's a common pattern throughout the OpenOffice XML file format that a simple type is wrapped in a complex type
    // with a single attribute called `val`. This is a small wrapper function to reduce the boiler plate for such
    // complex types
    fn get_val_attribute(&self) -> Result<&String, MissingAttributeError>;

    fn parse_val_attribute<T>(&self) -> Result<T, Box<dyn Error>>
    where
        T: FromStr,
        T::Err: Error + 'static,
    {
        self.get_val_attribute()
            .map_err(Box::<dyn Error>::from)
            .and_then(|value| value.parse().map_err(Box::from))
    }
}

impl XmlNodeExt for XmlNode {
    fn get_val_attribute(&self) -> Result<&String, MissingAttributeError> {
        self.attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(self.name.clone(), "val"))
    }
}
