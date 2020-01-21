use std::{
    error::Error,
    fmt::{Display, Formatter, Result},
    num::ParseIntError,
};

/// An error indicating that an xml element doesn't have an attribute that's marked as required in the schema
#[derive(Debug, Clone, PartialEq)]
pub struct MissingAttributeError {
    pub node_name: String,
    pub attr: &'static str,
}

impl MissingAttributeError {
    pub fn new<T>(node_name: T, attr: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            attr,
        }
    }
}

impl Display for MissingAttributeError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Xml element '{}' is missing a required attribute: {}",
            self.node_name, self.attr
        )
    }
}

impl Error for MissingAttributeError {
    fn description(&self) -> &str {
        "Missing required attribute"
    }
}

/// An error indicating that an xml element doesn't have a child node that's marked as required in the schema
#[derive(Debug, Clone, PartialEq)]
pub struct MissingChildNodeError {
    pub node_name: String,
    pub child_node: &'static str,
}

impl MissingChildNodeError {
    pub fn new<T>(node_name: T, child_node: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            child_node,
        }
    }
}

impl Display for MissingChildNodeError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Xml element '{}' is missing a required child element: {}",
            self.node_name, self.child_node
        )
    }
}

impl Error for MissingChildNodeError {
    fn description(&self) -> &str {
        "Xml element missing required child element"
    }
}

/// An error indicating that an xml element is not a member of a given element group
#[derive(Debug, Clone, PartialEq)]
pub struct NotGroupMemberError {
    node_name: String,
    group: &'static str,
}

impl NotGroupMemberError {
    pub fn new<T>(node_name: T, group: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            group,
        }
    }
}

impl Display for NotGroupMemberError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "XmlNode '{}' is not a member of {} group",
            self.node_name, self.group
        )
    }
}

impl Error for NotGroupMemberError {
    fn description(&self) -> &str {
        "Xml element is not a group member error"
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MaxOccurs {
    Value(u32),
    Unbounded,
}

impl Display for MaxOccurs {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        match self {
            MaxOccurs::Value(val) => write!(f, "{}", val),
            MaxOccurs::Unbounded => write!(f, "unbounded"),
        }
    }
}

/// An error indicating that the xml element violates either minOccurs or maxOccurs of the schema
#[derive(Debug, Clone, PartialEq)]
pub struct LimitViolationError {
    node_name: String,
    violating_node_name: &'static str,
    min_occurs: u32,
    max_occurs: MaxOccurs,
    occurs: u32,
}

impl LimitViolationError {
    pub fn new<T>(
        node_name: T,
        violating_node_name: &'static str,
        min_occurs: u32,
        max_occurs: MaxOccurs,
        occurs: u32,
    ) -> Self
    where
        T: Into<String>,
    {
        LimitViolationError {
            node_name: node_name.into(),
            violating_node_name,
            min_occurs,
            max_occurs,
            occurs,
        }
    }
}

impl Display for LimitViolationError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Element {} violates the limits of occurance in element: {}. minOccurs: {}, maxOccurs: {}, occurance: {}",
            self.node_name, self.violating_node_name, self.min_occurs, self.max_occurs, self.occurs,
        )
    }
}

impl Error for LimitViolationError {
    fn description(&self) -> &str {
        "Occurance limit violation"
    }
}

/// An error indicating that the parsed xml document is invalid
#[derive(Debug, Clone, Copy, Default)]
pub struct InvalidXmlError {}

impl Display for InvalidXmlError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Invalid xml document")
    }
}

impl Error for InvalidXmlError {
    fn description(&self) -> &str {
        "Invalid xml error"
    }
}

/// Error indicating that an xml element's attribute is not a valid bool value
/// Valid bool values are: true, false, 0, 1
#[derive(Debug, Clone, PartialEq)]
pub struct ParseBoolError {
    pub attr_value: String,
}

impl ParseBoolError {
    pub fn new<T>(attr_value: T) -> Self
    where
        T: Into<String>,
    {
        Self {
            attr_value: attr_value.into(),
        }
    }
}

impl Display for ParseBoolError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Xml attribute is not a valid bool value: {}", self.attr_value)
    }
}

impl Error for ParseBoolError {
    fn description(&self) -> &str {
        "Xml attribute is not a valid bool value"
    }
}

/// Error indicating that a string cannot be converted to an enum type
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ParseEnumError {
    enum_name: &'static str,
}

impl ParseEnumError {
    pub fn new(enum_name: &'static str) -> Self {
        Self { enum_name }
    }
}

impl Display for ParseEnumError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Cannot convert string to {}", self.enum_name)
    }
}

impl Error for ParseEnumError {
    fn description(&self) -> &str {
        "Cannot convert string to enum"
    }
}

/// Error indicating that parsing an AdjCoordinate or AdjAngle has failed
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AdjustParseError {}

impl Display for AdjustParseError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "AdjCoordinate or AdjAngle parse error")
    }
}

impl Error for AdjustParseError {
    fn description(&self) -> &str {
        "Adjust parse error"
    }
}

/// Error indicating that parsing a str as HexColorRGB has failed
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParseHexColorRGBError {
    Parse(ParseIntError),
    InvalidLength(StringLengthMismatch),
}

impl Display for ParseHexColorRGBError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            ParseHexColorRGBError::Parse(ref err) => err.fmt(f),
            ParseHexColorRGBError::InvalidLength(ref mismatch) => write!(
                f,
                "length of string should be {} but {} is provided",
                mismatch.required, mismatch.provided
            ),
        }
    }
}

impl From<ParseIntError> for ParseHexColorRGBError {
    fn from(v: ParseIntError) -> Self {
        ParseHexColorRGBError::Parse(v)
    }
}

impl Error for ParseHexColorRGBError {}

/// Struct used to describe invalid length errors
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StringLengthMismatch {
    pub required: usize,
    pub provided: usize,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PatternRestrictionError {
    NoMatch,
}

impl Display for PatternRestrictionError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "string doesn't match pattern")
    }
}

impl Error for PatternRestrictionError {}

#[derive(Debug, Clone, PartialEq)]
pub enum ParseHexColorError {
    Enum(ParseEnumError),
    HexColorRGB(ParseHexColorRGBError),
}

impl Display for ParseHexColorError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            ParseHexColorError::Enum(ref e) => e.fmt(f),
            ParseHexColorError::HexColorRGB(ref e) => e.fmt(f),
        }
    }
}

impl Error for ParseHexColorError {}

impl From<ParseEnumError> for ParseHexColorError {
    fn from(v: ParseEnumError) -> Self {
        ParseHexColorError::Enum(v)
    }
}

impl From<ParseHexColorRGBError> for ParseHexColorError {
    fn from(v: ParseHexColorRGBError) -> Self {
        ParseHexColorError::HexColorRGB(v)
    }
}
