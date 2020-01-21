use super::{
    drawing::{Anchor, Inline},
    simpletypes::{
        parse_on_off_xml_element, parse_text_scale_percent, DateTime, DecimalNumber, EightPointMeasure, FFHelpTextVal,
        FFName, FFStatusTextVal, LongHexNumber, MacroName, PointMeasure, ShortHexNumber, TextScale, UcharHexNumber,
        UnqualifiedPercentage, UnsignedDecimalNumber,
    },
    table::Tbl,
    util::XmlNodeExt,
};
use log::info;
use crate::{
    shared::{
        drawingml::simpletypes::{parse_hex_color_rgb, HexColorRGB},
        relationship::RelationshipId,
        sharedtypes::{
            CalendarType, ConformanceClass, Lang, OnOff, Percentage, PositiveUniversalMeasure, TwipsMeasure,
            UniversalMeasure, VerticalAlignRun, XAlign, XmlName, YAlign,
        },
    },
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError, ParseHexColorError},
    update::{update_options, Update},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::str::FromStr;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Charset {
    pub value: Option<UcharHexNumber>,
    pub character_set: Option<String>,
}

impl Charset {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Charset");

        let mut instance: Charset = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => instance.value = Some(UcharHexNumber::from_str_radix(value, 16)?),
                "w:characterSet" => instance.character_set = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum DecimalNumberOrPercent {
    Decimal(UnqualifiedPercentage),
    Percentage(Percentage),
}

impl FromStr for DecimalNumberOrPercent {
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        if let Ok(value) = s.parse::<UnqualifiedPercentage>() {
            Ok(DecimalNumberOrPercent::Decimal(value))
        } else {
            Ok(DecimalNumberOrPercent::Percentage(s.parse()?))
        }
    }
}

// pub enum TextScale {
//     Percent(TextScalePercent),
//     Decimal(TextScaleDecimal),
// }

#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum ThemeColor {
    #[strum(serialize = "dark1")]
    Dark1,
    #[strum(serialize = "light1")]
    Light1,
    #[strum(serialize = "dark2")]
    Dark2,
    #[strum(serialize = "light2")]
    Light2,
    #[strum(serialize = "accent1")]
    Accent1,
    #[strum(serialize = "accent2")]
    Accent2,
    #[strum(serialize = "accent3")]
    Accent3,
    #[strum(serialize = "accent4")]
    Accent4,
    #[strum(serialize = "accent5")]
    Accent5,
    #[strum(serialize = "accent6")]
    Accent6,
    #[strum(serialize = "hyperlink")]
    Hyperlink,
    #[strum(serialize = "followedHyperlink")]
    FollowedHyperlink,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "background1")]
    Background1,
    #[strum(serialize = "text1")]
    Text1,
    #[strum(serialize = "background2")]
    Background2,
    #[strum(serialize = "text2")]
    Text2,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum HighlightColor {
    #[strum(serialize = "black")]
    Black,
    #[strum(serialize = "blue")]
    Blue,
    #[strum(serialize = "cyan")]
    Cyan,
    #[strum(serialize = "green")]
    Green,
    #[strum(serialize = "magenta")]
    Magenta,
    #[strum(serialize = "red")]
    Red,
    #[strum(serialize = "yellow")]
    Yellow,
    #[strum(serialize = "white")]
    White,
    #[strum(serialize = "darkBlue")]
    DarkBlue,
    #[strum(serialize = "darkCyan")]
    DarkCyan,
    #[strum(serialize = "darkGreen")]
    DarkGreen,
    #[strum(serialize = "darkMagenta")]
    DarkMagenta,
    #[strum(serialize = "darkRed")]
    DarkRed,
    #[strum(serialize = "darkYellow")]
    DarkYellow,
    #[strum(serialize = "darkGray")]
    DarkGray,
    #[strum(serialize = "lightGray")]
    LightGray,
    #[strum(serialize = "none")]
    None,
}

impl HighlightColor {
    pub fn to_rgb(self) -> Option<[u8; 3]> {
        match self {
            HighlightColor::Black => Some([0, 0, 0]),
            HighlightColor::Blue => Some([0, 0, 0xff]),
            HighlightColor::Cyan => Some([0, 0xff, 0xff]),
            HighlightColor::Green => Some([0, 0xff, 0]),
            HighlightColor::Magenta => Some([0xff, 0, 0xff]),
            HighlightColor::Red => Some([0xff, 0, 0]),
            HighlightColor::Yellow => Some([0xff, 0xff, 0]),
            HighlightColor::White => Some([0xff, 0xff, 0xff]),
            HighlightColor::DarkBlue => Some([0, 0, 0x8b]),
            HighlightColor::DarkCyan => Some([0, 0x8b, 0x8b]),
            HighlightColor::DarkGreen => Some([0, 0x64, 0]),
            HighlightColor::DarkMagenta => Some([0x80, 0, 0x80]),
            HighlightColor::DarkRed => Some([0x8b, 0, 0]),
            HighlightColor::DarkYellow => Some([0x80, 0x80, 0]),
            HighlightColor::DarkGray => Some([0xa9, 0xa9, 0xa9]),
            HighlightColor::LightGray => Some([0xd3, 0xd3, 0xd3]),
            HighlightColor::None => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HexColor {
    Auto,
    RGB(HexColorRGB),
}

impl FromStr for HexColor {
    type Err = ParseHexColorError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "auto" => Ok(HexColor::Auto),
            _ => Ok(HexColor::RGB(parse_hex_color_rgb(s)?)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SignedTwipsMeasure {
    Decimal(i32),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for SignedTwipsMeasure {
    // TODO custom error type
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        // TODO maybe use TryFrom instead?
        if let Ok(value) = s.parse::<i32>() {
            Ok(SignedTwipsMeasure::Decimal(value))
        } else {
            Ok(SignedTwipsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl SignedTwipsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HpsMeasure {
    Decimal(u64),
    UniversalMeasure(PositiveUniversalMeasure),
}

impl FromStr for HpsMeasure {
    type Err = Box<dyn ::std::error::Error>;

    fn from_str(s: &str) -> Result<Self> {
        if let Ok(value) = s.parse::<u64>() {
            Ok(HpsMeasure::Decimal(value))
        } else {
            Ok(HpsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl HpsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SignedHpsMeasure {
    Decimal(i32),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for SignedHpsMeasure {
    // TODO custom error type
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        // TODO maybe use TryFrom instead?
        if let Ok(value) = s.parse::<i32>() {
            Ok(SignedHpsMeasure::Decimal(value))
        } else {
            Ok(SignedHpsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl SignedHpsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Color {
    pub value: HexColor,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
}

impl Color {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Color");

        let mut value = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:themeColor" => theme_color = Some(attr_value.parse()?),
                "w:themeTint" => theme_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "w:themeShade" => theme_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self {
            value,
            theme_color,
            theme_tint,
            theme_shade,
        })
    }
}

impl Update for Color {
    fn update_with(self, other: Self) -> Self {
        Self {
            value: other.value,
            theme_color: other.theme_color.or(self.theme_color),
            theme_tint: other.theme_tint.or(self.theme_tint),
            theme_shade: other.theme_shade.or(self.theme_shade),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ProofErrType {
    #[strum(serialize = "spellStart")]
    SpellingStart,
    #[strum(serialize = "spellEnd")]
    SpellingEnd,
    #[strum(serialize = "gramStart")]
    GrammarStart,
    #[strum(serialize = "gramEnd")]
    GrammarEnd,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ProofErr {
    pub error_type: ProofErrType,
}

impl ProofErr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ProofErr");

        let type_attr = xml_node
            .attributes
            .get("w:type")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self {
            error_type: type_attr.parse()?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum EdGrp {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "everyone")]
    Everyone,
    #[strum(serialize = "administrators")]
    Administrators,
    #[strum(serialize = "contributors")]
    Contributors,
    #[strum(serialize = "editors")]
    Editors,
    #[strum(serialize = "owners")]
    Owners,
    #[strum(serialize = "current")]
    Current,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum DisplacedByCustomXml {
    #[strum(serialize = "next")]
    Next,
    #[strum(serialize = "prev")]
    Prev,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Perm {
    pub id: String,
    pub displaced_by_custom_xml: Option<DisplacedByCustomXml>,
}

impl Perm {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Perm");

        let mut id = None;
        let mut displaced_by_custom_xml = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:id" => id = Some(value.clone()),
                "w:displacedByCustomXml" => displaced_by_custom_xml = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            id: id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?,
            displaced_by_custom_xml,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PermStart {
    pub permission: Perm,
    pub editor_group: Option<EdGrp>,
    pub editor: Option<String>,
    pub first_column: Option<DecimalNumber>,
    pub last_column: Option<DecimalNumber>,
}

impl PermStart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PermStart");

        let permission = Perm::from_xml_element(xml_node)?;
        let mut editor_group = None;
        let mut editor = None;
        let mut first_column = None;
        let mut last_column = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:edGrp" => editor_group = Some(value.parse()?),
                "w:ed" => editor = Some(value.clone()),
                "w:colFirst" => first_column = Some(value.parse()?),
                "w:colLast" => last_column = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            permission,
            editor_group,
            editor,
            first_column,
            last_column,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Markup {
    pub id: DecimalNumber,
}

impl Markup {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Markup");

        let id_attr = xml_node
            .attributes
            .get("w:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        Ok(Self { id: id_attr.parse()? })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MarkupRange {
    pub base: Markup,
    pub displaced_by_custom_xml: Option<DisplacedByCustomXml>,
}

impl MarkupRange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing MarkupRange");

        let base = Markup::from_xml_element(xml_node)?;
        let displaced_by_custom_xml = xml_node
            .attributes
            .get("w:displacedByCustomXml")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self {
            base,
            displaced_by_custom_xml,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BookmarkRange {
    pub base: MarkupRange,
    pub first_column: Option<DecimalNumber>,
    pub last_column: Option<DecimalNumber>,
}

impl BookmarkRange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing BookmarkRange");

        let base = MarkupRange::from_xml_element(xml_node)?;
        let first_column = xml_node
            .attributes
            .get("w:colFirst")
            .map(|value| value.parse())
            .transpose()?;

        let last_column = xml_node
            .attributes
            .get("w:colLast")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self {
            base,
            first_column,
            last_column,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Bookmark {
    pub base: BookmarkRange,
    pub name: String,
}

impl Bookmark {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Bookmark");

        let base = BookmarkRange::from_xml_element(xml_node)?;
        let name = xml_node
            .attributes
            .get("w:name")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?
            .clone();

        Ok(Self { base, name })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct MoveBookmark {
    pub base: Bookmark,
    pub author: String,
    pub date: DateTime,
}

impl MoveBookmark {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing MoveBookmark");

        let base = Bookmark::from_xml_element(xml_node)?;
        let author = xml_node
            .attributes
            .get("w:author")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?
            .clone();

        let date = xml_node
            .attributes
            .get("w:date")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "date"))?
            .clone();

        Ok(Self { base, author, date })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TrackChange {
    pub base: Markup,
    pub author: String,
    pub date: Option<DateTime>,
}

impl TrackChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TrackChange");

        let base = Markup::from_xml_element(xml_node)?;
        let author = xml_node
            .attributes
            .get("w:author")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?
            .clone();

        let date = xml_node.attributes.get("w:date").cloned();

        Ok(Self { base, author, date })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Attr {
    pub uri: String,
    pub name: String,
    pub value: String,
}

impl Attr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Attr");

        let mut uri = None;
        let mut name = None;
        let mut value = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(attr_value.clone()),
                "w:name" => name = Some(attr_value.clone()),
                "w:val" => value = Some(attr_value.clone()),
                _ => (),
            }
        }

        Ok(Self {
            uri: uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?,
            name: name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?,
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlPr {
    pub placeholder: Option<String>,
    pub attributes: Vec<Attr>,
}

impl CustomXmlPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CustomXmlPr");

        let mut placeholder = None;
        let mut attributes = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "placeholder" => placeholder = Some(child_node.get_val_attribute()?.clone()),
                "attr" => attributes.push(Attr::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            placeholder,
            attributes,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SimpleField {
    pub paragraph_contents: Vec<PContent>,
    pub field_codes: String,
    pub field_lock: Option<OnOff>,
    pub dirty: Option<OnOff>,
}

impl SimpleField {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SimpleField");

        let mut field_codes = None;
        let mut field_lock = None;
        let mut dirty = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:instr" => field_codes = Some(value.clone()),
                "w:fldLock" => field_lock = Some(parse_xml_bool(value)?),
                "w:dirty" => dirty = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let paragraph_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(PContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        let field_codes = field_codes.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "instr"))?;

        Ok(Self {
            field_codes,
            field_lock,
            dirty,
            paragraph_contents,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Hyperlink {
    pub paragraph_contents: Vec<PContent>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub document_location: Option<String>,
    pub history: Option<OnOff>,
    pub anchor: Option<String>,
    pub rel_id: Option<RelationshipId>,
}

impl Hyperlink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Hyperlink");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:tgtFrame" => instance.target_frame = Some(value.clone()),
                "w:tooltip" => instance.tooltip = Some(value.clone()),
                "w:docLocation" => instance.document_location = Some(value.clone()),
                "w:history" => instance.history = Some(parse_xml_bool(value)?),
                "w:anchor" => instance.anchor = Some(value.clone()),
                "r:id" => instance.rel_id = Some(value.clone()),
                _ => (),
            }
        }

        instance.paragraph_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(PContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Rel {
    pub rel_id: RelationshipId,
}

impl Rel {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Rel");

        let rel_id = xml_node
            .attributes
            .get("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?
            .clone();

        Ok(Self { rel_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum PContent {
    ContentRunContent(Box<ContentRunContent>),
    SimpleField(SimpleField),
    Hyperlink(Hyperlink),
    SubDocument(Rel),
}

impl XsdType for PContent {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PContent");

        match xml_node.local_name() {
            "fldSimple" => Ok(PContent::SimpleField(SimpleField::from_xml_element(xml_node)?)),
            "hyperlink" => Ok(PContent::Hyperlink(Hyperlink::from_xml_element(xml_node)?)),
            "subDoc" => Ok(PContent::SubDocument(Rel::from_xml_element(xml_node)?)),
            node_name if ContentRunContent::is_choice_member(node_name) => Ok(PContent::ContentRunContent(Box::new(
                ContentRunContent::from_xml_element(xml_node)?,
            ))),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "PContent"))),
        }
    }
}

impl XsdChoice for PContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "fldSimple" | "hyperlink" | "subDoc" => true,
            _ => ContentRunContent::is_choice_member(&node_name),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlRun {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub paragraph_contents: Vec<PContent>,

    pub uri: String,
    pub element: String,
}

impl CustomXmlRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CustomXmlRun");

        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(value.clone()),
                "w:element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let mut custom_xml_properties = None;
        let mut paragraph_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name if PContent::is_choice_member(node_name) => {
                    paragraph_contents.push(PContent::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let uri = uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?;
        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;
        Ok(Self {
            custom_xml_properties,
            paragraph_contents,
            uri,
            element,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SmartTagPr {
    pub attributes: Vec<Attr>,
}

impl SmartTagPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SmartTagPr");

        let mut attributes = Vec::new();
        for child_node in &xml_node.child_nodes {
            if child_node.local_name() == "attr" {
                attributes.push(Attr::from_xml_element(child_node)?);
            }
        }

        Ok(Self { attributes })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SmartTagRun {
    pub smart_tag_properties: Option<SmartTagPr>,
    pub paragraph_contents: Vec<PContent>,
    pub uri: String,
    pub element: String,
}

impl SmartTagRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SmartTagRun");

        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(value.clone()),
                "w:element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let mut smart_tag_properties = None;
        let mut paragraph_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "smartTagPr" => smart_tag_properties = Some(SmartTagPr::from_xml_element(child_node)?),
                node_name if PContent::is_choice_member(node_name) => {
                    paragraph_contents.push(PContent::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let uri = uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?;
        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        Ok(Self {
            uri,
            element,
            smart_tag_properties,
            paragraph_contents,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Hint {
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "eastAsia")]
    EastAsia,
    #[strum(serialize = "cs")]
    ComplexScript,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Theme {
    #[strum(serialize = "majorEastAsia")]
    MajorEastAsia,
    #[strum(serialize = "majorBidi")]
    MajorBidirectional,
    #[strum(serialize = "majorAscii")]
    MajorAscii,
    #[strum(serialize = "majorHAnsi")]
    MajorHighAnsi,
    #[strum(serialize = "minorEastAsia")]
    MinorEastAsia,
    #[strum(serialize = "minorBidi")]
    MinorBidirectional,
    #[strum(serialize = "minorAscii")]
    MinorAscii,
    #[strum(serialize = "minorHAnsi")]
    MinorHighAnsi,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Fonts {
    pub hint: Option<Hint>,
    pub ascii: Option<String>,
    pub high_ansi: Option<String>,
    pub east_asia: Option<String>,
    pub complex_script: Option<String>,
    pub ascii_theme: Option<Theme>,
    pub high_ansi_theme: Option<Theme>,
    pub east_asia_theme: Option<Theme>,
    pub complex_script_theme: Option<Theme>,
}

impl Fonts {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Fonts");

        let mut instance: Fonts = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:hint" => instance.hint = Some(value.parse()?),
                "w:ascii" => instance.ascii = Some(value.clone()),
                "w:hAnsi" => instance.high_ansi = Some(value.clone()),
                "w:eastAsia" => instance.east_asia = Some(value.clone()),
                "w:cs" => instance.complex_script = Some(value.clone()),
                "w:asciiTheme" => instance.ascii_theme = Some(value.parse()?),
                "w:hAnsiTheme" => instance.high_ansi_theme = Some(value.parse()?),
                "w:eastAsiaTheme" => instance.east_asia_theme = Some(value.parse()?),
                "w:cstheme" => instance.complex_script_theme = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for Fonts {
    fn update_with(self, other: Self) -> Self {
        Self {
            hint: other.hint.or(self.hint),
            ascii: other.ascii.or(self.ascii),
            high_ansi: other.high_ansi.or(self.high_ansi),
            east_asia: other.east_asia.or(self.east_asia),
            complex_script: other.complex_script.or(self.complex_script),
            ascii_theme: other.ascii_theme.or(self.ascii_theme),
            high_ansi_theme: other.high_ansi_theme.or(self.high_ansi_theme),
            east_asia_theme: other.east_asia_theme.or(self.east_asia_theme),
            complex_script_theme: other.complex_script_theme.or(self.complex_script_theme),
        }
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum UnderlineType {
    #[strum(serialize = "single")]
    Single,
    #[strum(serialize = "words")]
    Words,
    #[strum(serialize = "double")]
    Double,
    #[strum(serialize = "thick")]
    Thick,
    #[strum(serialize = "dotted")]
    Dotted,
    #[strum(serialize = "dottedHeavy")]
    DottedHeavy,
    #[strum(serialize = "dash")]
    Dash,
    #[strum(serialize = "dashedHeavy")]
    DashedHeavy,
    #[strum(serialize = "dashLong")]
    DashLong,
    #[strum(serialize = "dashLongHeavy")]
    DashLongHeavy,
    #[strum(serialize = "dotDash")]
    DotDash,
    #[strum(serialize = "dashDotHeavy")]
    DashDotHeavy,
    #[strum(serialize = "dotDotDash")]
    DotDotDash,
    #[strum(serialize = "dashDotDotHeavy")]
    DashDotDotHeavy,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "wavyHeavy")]
    WavyHeavy,
    #[strum(serialize = "wavyDouble")]
    WavyDouble,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Underline {
    pub value: Option<UnderlineType>,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
}

impl Underline {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Underline");

        let mut instance: Underline = Default::default();
        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => instance.value = Some(attr_value.parse()?),
                "w:color" => instance.color = Some(attr_value.parse()?),
                "w:themeColor" => instance.theme_color = Some(attr_value.parse()?),
                "w:themeTint" => instance.theme_tint = Some(u8::from_str_radix(attr_value, 16)?),
                "w:themeShade" => instance.theme_shade = Some(u8::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for Underline {
    fn update_with(self, other: Self) -> Self {
        Self {
            value: other.value.or(self.value),
            color: other.color.or(self.color),
            theme_color: other.theme_color.or(self.theme_color),
            theme_tint: other.theme_tint.or(self.theme_tint),
            theme_shade: other.theme_shade.or(self.theme_shade),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TextEffect {
    #[strum(serialize = "blinkBackground")]
    BlinkBackground,
    #[strum(serialize = "lights")]
    Lights,
    #[strum(serialize = "antsBlack")]
    AntsBlack,
    #[strum(serialize = "antsRed")]
    AntsRed,
    #[strum(serialize = "shimmer")]
    Shimmer,
    #[strum(serialize = "sparkle")]
    Sparkle,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum BorderType {
    #[strum(serialize = "nil")]
    Nil,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "single")]
    Single,
    #[strum(serialize = "thick")]
    Thick,
    #[strum(serialize = "double")]
    Double,
    #[strum(serialize = "dotted")]
    Dotted,
    #[strum(serialize = "dashed")]
    Dashed,
    #[strum(serialize = "dotDash")]
    DotDash,
    #[strum(serialize = "dotDotDash")]
    DotDotDash,
    #[strum(serialize = "triple")]
    Triple,
    #[strum(serialize = "thinThickSmallGap")]
    ThinThickSmallGap,
    #[strum(serialize = "thickThinSmallGap")]
    ThickThinSmallGap,
    #[strum(serialize = "thinThickThinSmallGap")]
    ThinThickThinSmallGap,
    #[strum(serialize = "thinThickMediumGap")]
    ThinThickMediumGap,
    #[strum(serialize = "thickThinMediumGap")]
    ThickThinMediumGap,
    #[strum(serialize = "thinThickThinMediumGap")]
    ThinThickThinMediumGap,
    #[strum(serialize = "thinThickLargeGap")]
    ThinThickLargeGap,
    #[strum(serialize = "thickThinLargeGap")]
    ThickThinLargeGap,
    #[strum(serialize = "thinThickThinLargeGap")]
    ThinThickThinLargeGap,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "doubleWave")]
    DoubleWave,
    #[strum(serialize = "dashSmallGap")]
    DashSmallGap,
    #[strum(serialize = "dashDotStroked")]
    DashDotStroked,
    #[strum(serialize = "threeDEmboss")]
    ThreeDEmboss,
    #[strum(serialize = "threeDEngrave")]
    ThreeDEngrave,
    #[strum(serialize = "outset")]
    Outset,
    #[strum(serialize = "inset")]
    Inset,
    #[strum(serialize = "apples")]
    Apples,
    #[strum(serialize = "archedScallops")]
    ArchedScallops,
    #[strum(serialize = "babyPacifier")]
    BabyPacifier,
    #[strum(serialize = "babyRattle")]
    BabyRattle,
    #[strum(serialize = "balloons3Colors")]
    Balloons3Colors,
    #[strum(serialize = "balloonsHotAir")]
    BalloonsHotAir,
    #[strum(serialize = "basicBlackDashes")]
    BasicBlackDashes,
    #[strum(serialize = "basicBlackDots")]
    BasicBlackDots,
    #[strum(serialize = "basicBlackSquares")]
    BasicBlackSquares,
    #[strum(serialize = "basicThinLines")]
    BasicThinLines,
    #[strum(serialize = "basicWhiteDashes")]
    BasicWhiteDashes,
    #[strum(serialize = "basicWhiteDots")]
    BasicWhiteDots,
    #[strum(serialize = "basicWhiteSquares")]
    BasicWhiteSquares,
    #[strum(serialize = "basicWideInline")]
    BasicWideInline,
    #[strum(serialize = "basicWideMidline")]
    BasicWideMidline,
    #[strum(serialize = "basicWideOutline")]
    BasicWideOutline,
    #[strum(serialize = "bats")]
    Bats,
    #[strum(serialize = "birds")]
    Birds,
    #[strum(serialize = "birdsFlight")]
    BirdsFlight,
    #[strum(serialize = "cabins")]
    Cabins,
    #[strum(serialize = "cakeSlice")]
    CakeSlice,
    #[strum(serialize = "candyCorn")]
    CandyCorn,
    #[strum(serialize = "celticKnotwork")]
    CelticKnotwork,
    #[strum(serialize = "certificateBanner")]
    CertificateBanner,
    #[strum(serialize = "chainLink")]
    ChainLink,
    #[strum(serialize = "champagneBottle")]
    ChampagneBottle,
    #[strum(serialize = "checkedBarBlack")]
    CheckedBarBlack,
    #[strum(serialize = "checkedBarColor")]
    CheckedBarColor,
    #[strum(serialize = "checkered")]
    Checkered,
    #[strum(serialize = "christmasTree")]
    ChristmasTree,
    #[strum(serialize = "circlesLines")]
    CirclesLines,
    #[strum(serialize = "circlesRectangles")]
    CirclesRectangles,
    #[strum(serialize = "classicalWave")]
    ClassicalWave,
    #[strum(serialize = "clocks")]
    Clocks,
    #[strum(serialize = "compass")]
    Compass,
    #[strum(serialize = "confetti")]
    Confetti,
    #[strum(serialize = "confettiGrays")]
    ConfettiGrays,
    #[strum(serialize = "confettiOutline")]
    ConfettiOutline,
    #[strum(serialize = "confettiStreamers")]
    ConfettiStreamers,
    #[strum(serialize = "confettiWhite")]
    ConfettiWhite,
    #[strum(serialize = "cornerTriangles")]
    CornerTriangles,
    #[strum(serialize = "couponCutoutDashes")]
    CouponCutoutDashes,
    #[strum(serialize = "couponCutoutDots")]
    CouponCutoutDots,
    #[strum(serialize = "crazyMaze")]
    CrazyMaze,
    #[strum(serialize = "creaturesButterfly")]
    CreaturesButterfly,
    #[strum(serialize = "creaturesFish")]
    CreaturesFish,
    #[strum(serialize = "creaturesInsects")]
    CreaturesInsects,
    #[strum(serialize = "creaturesLadyBug")]
    CreaturesLadyBug,
    #[strum(serialize = "crossStitch")]
    CrossStitch,
    #[strum(serialize = "cup")]
    Cup,
    #[strum(serialize = "decoArch")]
    DecoArch,
    #[strum(serialize = "decoArchColor")]
    DecoArchColor,
    #[strum(serialize = "decoBlocks")]
    DecoBlocks,
    #[strum(serialize = "diamondsGray")]
    DiamondsGray,
    #[strum(serialize = "doubleD")]
    DoubleD,
    #[strum(serialize = "doubleDiamonds")]
    DoubleDiamonds,
    #[strum(serialize = "earth1")]
    Earth1,
    #[strum(serialize = "earth2")]
    Earth2,
    #[strum(serialize = "earth3")]
    Earth3,
    #[strum(serialize = "eclipsingSquares1")]
    EclipsingSquares1,
    #[strum(serialize = "eclipsingSquares2")]
    EclipsingSquares2,
    #[strum(serialize = "eggsBlack")]
    EggsBlack,
    #[strum(serialize = "fans")]
    Fans,
    #[strum(serialize = "film")]
    Film,
    #[strum(serialize = "firecrackers")]
    Firecrackers,
    #[strum(serialize = "flowersBlockPrint")]
    FlowersBlockPrint,
    #[strum(serialize = "flowersDaisies")]
    FlowersDaisies,
    #[strum(serialize = "flowersModern1")]
    FlowersModern1,
    #[strum(serialize = "flowersModern2")]
    FlowersModern2,
    #[strum(serialize = "flowersPansy")]
    FlowersPansy,
    #[strum(serialize = "flowersRedRose")]
    FlowersRedRose,
    #[strum(serialize = "flowersRoses")]
    FlowersRoses,
    #[strum(serialize = "flowersTeacup")]
    FlowersTeacup,
    #[strum(serialize = "flowersTiny")]
    FlowersTiny,
    #[strum(serialize = "gems")]
    Gems,
    #[strum(serialize = "gingerbreadMan")]
    GingerbreadMan,
    #[strum(serialize = "gradient")]
    Gradient,
    #[strum(serialize = "handmade1")]
    Handmade1,
    #[strum(serialize = "handmade2")]
    Handmade2,
    #[strum(serialize = "heartBalloon")]
    HeartBalloon,
    #[strum(serialize = "heartGray")]
    HeartGray,
    #[strum(serialize = "hearts")]
    Hearts,
    #[strum(serialize = "heebieJeebies")]
    HeebieJeebies,
    #[strum(serialize = "holly")]
    Holly,
    #[strum(serialize = "houseFunky")]
    HouseFunky,
    #[strum(serialize = "hypnotic")]
    Hypnotic,
    #[strum(serialize = "iceCreamCones")]
    IceCreamCones,
    #[strum(serialize = "lightBulb")]
    LightBulb,
    #[strum(serialize = "lightning1")]
    Lightning1,
    #[strum(serialize = "lightning2")]
    Lightning2,
    #[strum(serialize = "mapPins")]
    MapPins,
    #[strum(serialize = "mapleLeaf")]
    MapleLeaf,
    #[strum(serialize = "mapleMuffins")]
    MapleMuffins,
    #[strum(serialize = "marquee")]
    Marquee,
    #[strum(serialize = "marqueeToothed")]
    MarqueeToothed,
    #[strum(serialize = "moons")]
    Moons,
    #[strum(serialize = "mosaic")]
    Mosaic,
    #[strum(serialize = "musicNotes")]
    MusicNotes,
    #[strum(serialize = "northwest")]
    Northwest,
    #[strum(serialize = "ovals")]
    Ovals,
    #[strum(serialize = "packages")]
    Packages,
    #[strum(serialize = "palmsBlack")]
    PalmsBlack,
    #[strum(serialize = "palmsColor")]
    PalmsColor,
    #[strum(serialize = "paperClips")]
    PaperClips,
    #[strum(serialize = "papyrus")]
    Papyrus,
    #[strum(serialize = "partyFavor")]
    PartyFavor,
    #[strum(serialize = "partyGlass")]
    PartyGlass,
    #[strum(serialize = "pencils")]
    Pencils,
    #[strum(serialize = "people")]
    People,
    #[strum(serialize = "peopleWaving")]
    PeopleWaving,
    #[strum(serialize = "peopleHats")]
    PeopleHats,
    #[strum(serialize = "poinsettias")]
    Poinsettias,
    #[strum(serialize = "postageStamp")]
    PostageStamp,
    #[strum(serialize = "pumpkin1")]
    Pumpkin1,
    #[strum(serialize = "pushPinNote2")]
    PushPinNote2,
    #[strum(serialize = "pushPinNote1")]
    PushPinNote1,
    #[strum(serialize = "pyramids")]
    Pyramids,
    #[strum(serialize = "pyramidsAbove")]
    PyramidsAbove,
    #[strum(serialize = "quadrants")]
    Quadrants,
    #[strum(serialize = "rings")]
    Rings,
    #[strum(serialize = "safari")]
    Safari,
    #[strum(serialize = "sawtooth")]
    Sawtooth,
    #[strum(serialize = "sawtoothGray")]
    SawtoothGray,
    #[strum(serialize = "scaredCat")]
    ScaredCat,
    #[strum(serialize = "seattle")]
    Seattle,
    #[strum(serialize = "shadowedSquares")]
    ShadowedSquares,
    #[strum(serialize = "sharksTeeth")]
    SharksTeeth,
    #[strum(serialize = "shorebirdTracks")]
    ShorebirdTracks,
    #[strum(serialize = "skyrocket")]
    Skyrocket,
    #[strum(serialize = "snowflakeFancy")]
    SnowflakeFancy,
    #[strum(serialize = "snowflakes")]
    Snowflakes,
    #[strum(serialize = "sombrero")]
    Sombrero,
    #[strum(serialize = "southwest")]
    Southwest,
    #[strum(serialize = "stars")]
    Stars,
    #[strum(serialize = "starsTop")]
    StarsTop,
    #[strum(serialize = "stars3d")]
    Stars3d,
    #[strum(serialize = "starsBlack")]
    StarsBlack,
    #[strum(serialize = "starsShadowed")]
    StarsShadowed,
    #[strum(serialize = "sun")]
    Sun,
    #[strum(serialize = "swirligig")]
    Swirligig,
    #[strum(serialize = "tornPaper")]
    TornPaper,
    #[strum(serialize = "tornPaperBlack")]
    TornPaperBlack,
    #[strum(serialize = "trees")]
    Trees,
    #[strum(serialize = "triangleParty")]
    TriangleParty,
    #[strum(serialize = "triangles")]
    Triangles,
    #[strum(serialize = "triangle1")]
    Triangle1,
    #[strum(serialize = "triangle2")]
    Triangle2,
    #[strum(serialize = "triangleCircle1")]
    TriangleCircle1,
    #[strum(serialize = "triangleCircle2")]
    TriangleCircle2,
    #[strum(serialize = "shapes1")]
    Shapes1,
    #[strum(serialize = "shapes2")]
    Shapes2,
    #[strum(serialize = "twistedLines1")]
    TwistedLines1,
    #[strum(serialize = "twistedLines2")]
    TwistedLines2,
    #[strum(serialize = "vine")]
    Vine,
    #[strum(serialize = "waveline")]
    Waveline,
    #[strum(serialize = "weavingAngles")]
    WeavingAngles,
    #[strum(serialize = "weavingBraid")]
    WeavingBraid,
    #[strum(serialize = "weavingRibbon")]
    WeavingRibbon,
    #[strum(serialize = "weavingStrips")]
    WeavingStrips,
    #[strum(serialize = "whiteFlowers")]
    WhiteFlowers,
    #[strum(serialize = "woodwork")]
    Woodwork,
    #[strum(serialize = "xIllusions")]
    XIllusions,
    #[strum(serialize = "zanyTriangles")]
    ZanyTriangles,
    #[strum(serialize = "zigZag")]
    ZigZag,
    #[strum(serialize = "zigZagStitch")]
    ZigZagStitch,
    #[strum(serialize = "custom")]
    Custom,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Border {
    pub value: BorderType,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
    pub size: Option<EightPointMeasure>,
    pub spacing: Option<PointMeasure>,
    pub shadow: Option<OnOff>,
    pub frame: Option<OnOff>,
}

impl Border {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Border");

        let mut value = None;
        let mut color = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;
        let mut size = None;
        let mut spacing = None;
        let mut shadow = None;
        let mut frame = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:color" => color = Some(attr_value.parse()?),
                "w:themeColor" => theme_color = Some(attr_value.parse()?),
                "w:themeTint" => theme_tint = Some(u8::from_str_radix(attr_value, 16)?),
                "w:themeShade" => theme_shade = Some(u8::from_str_radix(attr_value, 16)?),
                "w:sz" => size = Some(attr_value.parse()?),
                "w:space" => spacing = Some(attr_value.parse()?),
                "w:shadow" => shadow = Some(parse_xml_bool(attr_value)?),
                "w:frame" => frame = Some(parse_xml_bool(attr_value)?),
                _ => (),
            }
        }

        Ok(Self {
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
            color,
            theme_color,
            theme_tint,
            theme_shade,
            size,
            spacing,
            shadow,
            frame,
        })
    }
}

impl Update for Border {
    fn update_with(self, other: Self) -> Self {
        Self {
            value: other.value,
            color: other.color.or(self.color),
            theme_color: other.theme_color.or(self.theme_color),
            theme_tint: other.theme_tint.or(self.theme_tint),
            theme_shade: other.theme_shade.or(self.theme_shade),
            size: other.size.or(self.size),
            spacing: other.spacing.or(self.spacing),
            shadow: other.shadow.or(self.shadow),
            frame: other.frame.or(self.frame),
        }
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ShdType {
    #[strum(serialize = "nil")]
    Nil,
    #[strum(serialize = "clear")]
    Clear,
    #[strum(serialize = "solid")]
    Solid,
    #[strum(serialize = "horzStripe")]
    HorizontalStripe,
    #[strum(serialize = "vertStripe")]
    VerticalStripe,
    #[strum(serialize = "reverseDiagStripe")]
    ReverseDiagonalStripe,
    #[strum(serialize = "diagStripe")]
    DiagonalStripe,
    #[strum(serialize = "horzCross")]
    HorizontalCross,
    #[strum(serialize = "diagCross")]
    DiagonalCross,
    #[strum(serialize = "thinHorzStripe")]
    ThinHorizontalStripe,
    #[strum(serialize = "thinVertStripe")]
    ThinVerticalStripe,
    #[strum(serialize = "thinReverseDiagStripe")]
    ThinReverseDiagonalStripe,
    #[strum(serialize = "thinDiagStripe")]
    ThinDiagonalStripe,
    #[strum(serialize = "thinHorzCross")]
    ThinHorizontalCross,
    #[strum(serialize = "thinDiagCross")]
    ThinDiagonalCross,
    #[strum(serialize = "pct5")]
    Percent5,
    #[strum(serialize = "pct10")]
    Percent10,
    #[strum(serialize = "pct12")]
    Percent12,
    #[strum(serialize = "pct15")]
    Percent15,
    #[strum(serialize = "pct20")]
    Percent20,
    #[strum(serialize = "pct25")]
    Percent25,
    #[strum(serialize = "pct30")]
    Percent30,
    #[strum(serialize = "pct35")]
    Percent35,
    #[strum(serialize = "pct37")]
    Percent37,
    #[strum(serialize = "pct40")]
    Percent40,
    #[strum(serialize = "pct45")]
    Percent45,
    #[strum(serialize = "pct50")]
    Percent50,
    #[strum(serialize = "pct55")]
    Percent55,
    #[strum(serialize = "pct60")]
    Percent60,
    #[strum(serialize = "pct62")]
    Percent62,
    #[strum(serialize = "pct65")]
    Percent65,
    #[strum(serialize = "pct70")]
    Percent70,
    #[strum(serialize = "pct75")]
    Percent75,
    #[strum(serialize = "pct80")]
    Percent80,
    #[strum(serialize = "pct85")]
    Percent85,
    #[strum(serialize = "pct87")]
    Percent87,
    #[strum(serialize = "pct90")]
    Percent90,
    #[strum(serialize = "pct95")]
    Percent95,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Shd {
    pub value: ShdType,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
    pub fill: Option<HexColor>,
    pub theme_fill: Option<ThemeColor>,
    pub theme_fill_tint: Option<UcharHexNumber>,
    pub theme_fill_shade: Option<UcharHexNumber>,
}

impl Shd {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Shd");

        let mut value = None;
        let mut color = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;
        let mut fill = None;
        let mut theme_fill = None;
        let mut theme_fill_tint = None;
        let mut theme_fill_shade = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:color" => color = Some(attr_value.parse()?),
                "w:themeColor" => theme_color = Some(attr_value.parse()?),
                "w:themeTint" => theme_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "w:themeShade" => theme_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "w:fill" => fill = Some(attr_value.parse()?),
                "w:themeFill" => theme_fill = Some(attr_value.parse()?),
                "w:themeFillTint" => theme_fill_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "w:themeFillShade" => theme_fill_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "value"))?;
        Ok(Self {
            value,
            color,
            theme_color,
            theme_tint,
            theme_shade,
            fill,
            theme_fill,
            theme_fill_tint,
            theme_fill_shade,
        })
    }
}

impl Update for Shd {
    fn update_with(self, other: Self) -> Self {
        Self {
            value: other.value,
            color: other.color.or(self.color),
            theme_color: other.theme_color.or(self.theme_color),
            theme_tint: other.theme_tint.or(self.theme_tint),
            theme_shade: other.theme_shade.or(self.theme_shade),
            fill: other.fill.or(self.fill),
            theme_fill: other.theme_fill.or(self.theme_fill),
            theme_fill_tint: other.theme_fill_tint.or(self.theme_fill_tint),
            theme_fill_shade: other.theme_fill_shade.or(self.theme_fill_shade),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FitText {
    pub value: TwipsMeasure,
    pub id: Option<DecimalNumber>,
}

impl FitText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FitText");

        let mut value = None;
        let mut id = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:id" => id = Some(attr_value.parse()?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self { value, id })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Em {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "comma")]
    Comma,
    #[strum(serialize = "circle")]
    Circle,
    #[strum(serialize = "underDot")]
    UnderDot,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Language {
    pub value: Option<Lang>,
    pub east_asia: Option<Lang>,
    pub bidirectional: Option<Lang>,
}

impl Language {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        info!("parsing Language");

        xml_node
            .attributes
            .iter()
            .fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:val" => instance.value = Some(value.clone()),
                    "w:eastAsia" => instance.east_asia = Some(value.clone()),
                    "w:bidi" => instance.bidirectional = Some(value.clone()),
                    _ => (),
                }

                instance
            })
    }
}

impl Update for Language {
    fn update_with(self, other: Self) -> Self {
        Self {
            value: other.value.or(self.value),
            east_asia: other.east_asia.or(self.east_asia),
            bidirectional: other.bidirectional.or(self.bidirectional),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum CombineBrackets {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "round")]
    Round,
    #[strum(serialize = "square")]
    Square,
    #[strum(serialize = "angle")]
    Angle,
    #[strum(serialize = "curly")]
    Curly,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct EastAsianLayout {
    pub id: Option<DecimalNumber>,
    pub combine: Option<OnOff>,
    pub combine_brackets: Option<CombineBrackets>,
    pub vertical: Option<OnOff>,
    pub vertical_compress: Option<OnOff>,
}

impl EastAsianLayout {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing EastAsianLayout");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:id" => instance.id = Some(value.parse()?),
                    "w:combine" => instance.combine = Some(parse_xml_bool(value)?),
                    "w:combineBrackets" => instance.combine_brackets = Some(value.parse()?),
                    "w:vert" => instance.vertical = Some(parse_xml_bool(value)?),
                    "w:vertCompress" => instance.vertical_compress = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

impl Update for EastAsianLayout {
    fn update_with(self, other: Self) -> Self {
        Self {
            id: other.id.or(self.id),
            combine: other.combine.or(self.combine),
            combine_brackets: other.combine_brackets.or(self.combine_brackets),
            vertical: other.vertical.or(self.vertical),
            vertical_compress: other.vertical_compress.or(self.vertical_compress),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RPrBase {
    RunStyle(String),
    RunFonts(Fonts),
    Bold(OnOff),
    ComplexScriptBold(OnOff),
    Italic(OnOff),
    ComplexScriptItalic(OnOff),
    Capitals(OnOff),
    SmallCapitals(OnOff),
    Strikethrough(OnOff),
    DoubleStrikethrough(OnOff),
    Outline(OnOff),
    Shadow(OnOff),
    Emboss(OnOff),
    Imprint(OnOff),
    NoProofing(OnOff),
    SnapToGrid(OnOff),
    Vanish(OnOff),
    WebHidden(OnOff),
    Color(Color),
    Spacing(SignedTwipsMeasure),
    Width(TextScale),
    Kerning(HpsMeasure),
    Position(SignedHpsMeasure),
    FontSize(HpsMeasure),
    ComplexScriptFontSize(HpsMeasure),
    Highlight(HighlightColor),
    Underline(Underline),
    Effect(TextEffect),
    Border(Border),
    Shading(Shd),
    FitText(FitText),
    VerticalAlignment(VerticalAlignRun),
    Rtl(OnOff),
    ComplexScript(OnOff),
    EmphasisMark(Em),
    Language(Language),
    EastAsianLayout(EastAsianLayout),
    SpecialVanish(OnOff),
    OMath(OnOff),
}

impl XsdType for RPrBase {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RPrBase");

        match xml_node.local_name() {
            "rStyle" => Ok(RPrBase::RunStyle(xml_node.get_val_attribute()?.clone())),
            "rFonts" => Ok(RPrBase::RunFonts(Fonts::from_xml_element(xml_node)?)),
            "b" => Ok(RPrBase::Bold(parse_on_off_xml_element(xml_node)?)),
            "bCs" => Ok(RPrBase::ComplexScriptBold(parse_on_off_xml_element(xml_node)?)),
            "i" => Ok(RPrBase::Italic(parse_on_off_xml_element(xml_node)?)),
            "iCs" => Ok(RPrBase::ComplexScriptItalic(parse_on_off_xml_element(xml_node)?)),
            "caps" => Ok(RPrBase::Capitals(parse_on_off_xml_element(xml_node)?)),
            "smallCaps" => Ok(RPrBase::SmallCapitals(parse_on_off_xml_element(xml_node)?)),
            "strike" => Ok(RPrBase::Strikethrough(parse_on_off_xml_element(xml_node)?)),
            "dstrike" => Ok(RPrBase::DoubleStrikethrough(parse_on_off_xml_element(xml_node)?)),
            "outline" => Ok(RPrBase::Outline(parse_on_off_xml_element(xml_node)?)),
            "shadow" => Ok(RPrBase::Shadow(parse_on_off_xml_element(xml_node)?)),
            "emboss" => Ok(RPrBase::Emboss(parse_on_off_xml_element(xml_node)?)),
            "imprint" => Ok(RPrBase::Imprint(parse_on_off_xml_element(xml_node)?)),
            "noProof" => Ok(RPrBase::NoProofing(parse_on_off_xml_element(xml_node)?)),
            "snapToGrid" => Ok(RPrBase::SnapToGrid(parse_on_off_xml_element(xml_node)?)),
            "vanish" => Ok(RPrBase::Vanish(parse_on_off_xml_element(xml_node)?)),
            "webHidden" => Ok(RPrBase::WebHidden(parse_on_off_xml_element(xml_node)?)),
            "color" => Ok(RPrBase::Color(Color::from_xml_element(xml_node)?)),
            "spacing" => Ok(RPrBase::Spacing(SignedTwipsMeasure::from_xml_element(xml_node)?)),
            "w" => {
                let val = xml_node
                    .attributes
                    .get("w:val")
                    .map(|val| parse_text_scale_percent(val))
                    .transpose()?
                    .unwrap_or(100.0);
                Ok(RPrBase::Width(val))
            }
            "kern" => Ok(RPrBase::Kerning(HpsMeasure::from_xml_element(xml_node)?)),
            "position" => Ok(RPrBase::Position(SignedHpsMeasure::from_xml_element(xml_node)?)),
            "sz" => Ok(RPrBase::FontSize(HpsMeasure::from_xml_element(xml_node)?)),
            "szCs" => Ok(RPrBase::ComplexScriptFontSize(HpsMeasure::from_xml_element(xml_node)?)),
            "highlight" => Ok(RPrBase::Highlight(xml_node.get_val_attribute()?.parse()?)),
            "u" => Ok(RPrBase::Underline(Underline::from_xml_element(xml_node)?)),
            "effect" => Ok(RPrBase::Effect(xml_node.get_val_attribute()?.parse()?)),
            "bdr" => Ok(RPrBase::Border(Border::from_xml_element(xml_node)?)),
            "shd" => Ok(RPrBase::Shading(Shd::from_xml_element(xml_node)?)),
            "fitText" => Ok(RPrBase::FitText(FitText::from_xml_element(xml_node)?)),
            "vertAlign" => Ok(RPrBase::VerticalAlignment(xml_node.get_val_attribute()?.parse()?)),
            "rtl" => Ok(RPrBase::Rtl(parse_on_off_xml_element(xml_node)?)),
            "cs" => Ok(RPrBase::ComplexScript(parse_on_off_xml_element(xml_node)?)),
            "em" => Ok(RPrBase::EmphasisMark(xml_node.get_val_attribute()?.parse()?)),
            "lang" => Ok(RPrBase::Language(Language::from_xml_element(xml_node))),
            "eastAsianLayout" => Ok(RPrBase::EastAsianLayout(EastAsianLayout::from_xml_element(xml_node)?)),
            "specVanish" => Ok(RPrBase::SpecialVanish(parse_on_off_xml_element(xml_node)?)),
            "oMath" => Ok(RPrBase::OMath(parse_on_off_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "RPrBase"))),
        }
    }
}

impl XsdChoice for RPrBase {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "rStyle" | "rFonts" | "b" | "bCs" | "i" | "iCs" | "caps" | "smallCaps" | "strike" | "dstrike"
            | "outline" | "shadow" | "emboss" | "imprint" | "noProof" | "snapToGrid" | "vanish" | "webHidden"
            | "color" | "spacing" | "w" | "kern" | "position" | "sz" | "szCs" | "highlight" | "u" | "effect"
            | "bdr" | "shd" | "fitText" | "vertAlign" | "rtl" | "cs" | "em" | "lang" | "eastAsianLayout"
            | "specVanish" | "oMath" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPrOriginal {
    pub r_pr_bases: Vec<RPrBase>,
}

impl RPrOriginal {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RPrOriginal");

        let r_pr_bases = xml_node
            .child_nodes
            .iter()
            .filter_map(RPrBase::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { r_pr_bases })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RPrChange {
    pub base: TrackChange,
    pub run_properties: RPrOriginal,
}

impl RPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let run_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rPr").into())
            .and_then(RPrOriginal::from_xml_element)?;

        Ok(Self { base, run_properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPr {
    pub r_pr_bases: Vec<RPrBase>,
    pub run_properties_change: Option<RPrChange>,
}

impl RPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RPr");

        let mut instance: RPr = Default::default();
        for child_node in &xml_node.child_nodes {
            let child_node_name = child_node.local_name();
            if RPrBase::is_choice_member(child_node_name) {
                instance.r_pr_bases.push(RPrBase::from_xml_element(child_node)?);
            } else if child_node_name == "rPrChange" {
                instance.run_properties_change = Some(RPrChange::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct SdtListItem {
    pub display_text: String,
    pub value: String,
}

impl SdtListItem {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtListItem");

        let display_text = xml_node
            .attributes
            .get("w:displayText")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "displayText"))?
            .clone();

        let value = xml_node
            .attributes
            .get("w:value")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "value"))?
            .clone();

        Ok(Self { display_text, value })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtComboBox {
    pub list_items: Vec<SdtListItem>,
    pub last_value: Option<String>,
}

impl SdtComboBox {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtComboBox");

        let last_value = xml_node.attributes.get("w:lastValue").cloned();

        let list_items = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "listItem")
            .map(SdtListItem::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum SdtDateMappingType {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "date")]
    Date,
    #[strum(serialize = "dateTime")]
    DateTime,
}

impl SdtDateMappingType {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        info!("parsing SdtDateMappingType");

        Ok(xml_node.attributes.get("w:val").map(|val| val.parse()).transpose()?)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDate {
    pub date_format: Option<String>,
    pub language_id: Option<Lang>,
    pub store_mapped_data_as: Option<SdtDateMappingType>,
    pub calendar: Option<CalendarType>,

    pub full_date: Option<DateTime>,
}

impl SdtDate {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtDate");

        let mut instance: Self = Default::default();
        instance.full_date = xml_node.attributes.get("w:fullDate").cloned();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "dateFormat" => instance.date_format = Some(child_node.get_val_attribute()?.clone()),
                "lid" => instance.language_id = Some(child_node.get_val_attribute()?.clone()),
                "storeMappedDataAs" => {
                    instance.store_mapped_data_as = SdtDateMappingType::from_xml_element(child_node)?
                }
                "calendar" => {
                    instance.calendar = child_node.attributes.get("w:val").map(|val| val.parse()).transpose()?;
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDocPart {
    pub doc_part_gallery: Option<String>,
    pub doc_part_category: Option<String>,
    pub doc_part_unique: Option<OnOff>,
}

impl SdtDocPart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtDocPart");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "docPartGallery" => instance.doc_part_gallery = Some(child_node.get_val_attribute()?.clone()),
                "docPartCategory" => instance.doc_part_category = Some(child_node.get_val_attribute()?.clone()),
                "docPartUnique" => instance.doc_part_unique = Some(parse_on_off_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDropDownList {
    pub list_items: Vec<SdtListItem>,
    pub last_value: Option<String>,
}

impl SdtDropDownList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtDropDownList");

        let last_value = xml_node.attributes.get("w:lastValue").cloned();

        let list_items = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "listItem")
            .map(SdtListItem::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SdtText {
    pub is_multi_line: OnOff,
}

impl SdtText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtText");

        let is_multi_line_attr = xml_node
            .attributes
            .get("w:multiLine")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "multiLine"))?;

        Ok(Self {
            is_multi_line: parse_xml_bool(is_multi_line_attr)?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum SdtPrChoice {
    Equation,
    ComboBox(SdtComboBox),
    Date(SdtDate),
    DocumentPartObject(SdtDocPart),
    DocumentPartList(SdtDocPart),
    DropDownList(SdtDropDownList),
    Picture,
    RichText,
    Text(SdtText),
    Citation,
    Group,
    Bibliography,
}

impl SdtPrChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "equation" | "comboBox" | "date" | "docPartObj" | "docPartList" | "dropDownList" | "picture"
            | "richText" | "text" | "citation" | "group" | "bibliography" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtPrChoice");

        match xml_node.local_name() {
            "equation" => Ok(SdtPrChoice::Equation),
            "comboBox" => Ok(SdtPrChoice::ComboBox(SdtComboBox::from_xml_element(xml_node)?)),
            "date" => Ok(SdtPrChoice::Date(SdtDate::from_xml_element(xml_node)?)),
            "docPartObj" => Ok(SdtPrChoice::DocumentPartObject(SdtDocPart::from_xml_element(xml_node)?)),
            "docPartList" => Ok(SdtPrChoice::DocumentPartList(SdtDocPart::from_xml_element(xml_node)?)),
            "dropDownList" => Ok(SdtPrChoice::DropDownList(SdtDropDownList::from_xml_element(xml_node)?)),
            "picture" => Ok(SdtPrChoice::Picture),
            "richText" => Ok(SdtPrChoice::RichText),
            "text" => Ok(SdtPrChoice::Text(SdtText::from_xml_element(xml_node)?)),
            "citation" => Ok(SdtPrChoice::Citation),
            "group" => Ok(SdtPrChoice::Group),
            "bibliography" => Ok(SdtPrChoice::Bibliography),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "SdtPrChoice"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Lock {
    #[strum(serialize = "sdtLocked")]
    SdtLocked,
    #[strum(serialize = "contentLocked")]
    ContentLocked,
    #[strum(serialize = "unlocked")]
    Unlocked,
    #[strum(serialize = "sdtContentLocked")]
    SdtContentLocked,
}

impl Lock {
    pub fn from_xml_element(xml_node: &XmlNode) -> std::result::Result<Option<Self>, strum::ParseError> {
        info!("parsing Lock");

        xml_node.attributes.get("w:val").map(|val| val.parse()).transpose()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Placeholder {
    pub document_part: String,
}

impl Placeholder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Placeholder");

        let document_part = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "docPart")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPart"))?
            .get_val_attribute()?
            .clone();

        Ok(Self { document_part })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataBinding {
    pub prefix_mappings: Option<String>,
    pub xpath: String,
    pub store_item_id: String,
}

impl DataBinding {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing DataBinding");

        let mut prefix_mappings = None;
        let mut xpath = None;
        let mut store_item_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:prefixMappings" => prefix_mappings = Some(value.clone()),
                "w:xpath" => xpath = Some(value.clone()),
                "w:storeItemID" => store_item_id = Some(value.clone()),
                _ => (),
            }
        }

        let xpath = xpath.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "xpath"))?;
        let store_item_id =
            store_item_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "storeItemId"))?;

        Ok(Self {
            prefix_mappings,
            xpath,
            store_item_id,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtPr {
    pub run_properties: Option<RPr>,
    pub alias: Option<String>,
    pub tag: Option<String>,
    pub id: Option<DecimalNumber>,
    pub lock: Option<Lock>,
    pub placeholder: Option<Placeholder>,
    pub temporary: Option<OnOff>,
    pub showing_placeholder_header: Option<OnOff>,
    pub data_binding: Option<DataBinding>,
    pub label: Option<DecimalNumber>,
    pub tab_index: Option<UnsignedDecimalNumber>,
    pub control_choice: Option<SdtPrChoice>,
}

impl SdtPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtPr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                "alias" => instance.alias = Some(child_node.get_val_attribute()?.clone()),
                "tag" => instance.tag = Some(child_node.get_val_attribute()?.clone()),
                "id" => instance.id = Some(child_node.get_val_attribute()?.parse()?),
                "lock" => instance.lock = child_node.attributes.get("w:val").map(|val| val.parse()).transpose()?,
                "placeholder" => instance.placeholder = Some(Placeholder::from_xml_element(child_node)?),
                "temporary" => instance.temporary = Some(parse_on_off_xml_element(child_node)?),
                "showingPlcHdr" => instance.showing_placeholder_header = Some(parse_on_off_xml_element(child_node)?),
                "dataBinding" => instance.data_binding = Some(DataBinding::from_xml_element(child_node)?),
                "label" => instance.label = Some(child_node.get_val_attribute()?.parse()?),
                "tabIndex" => instance.tab_index = Some(child_node.get_val_attribute()?.parse()?),
                node_name if SdtPrChoice::is_choice_member(node_name) => {
                    instance.control_choice = Some(SdtPrChoice::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtEndPr {
    pub run_properties_vec: Vec<RPr>,
}

impl SdtEndPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtEndPr");

        let run_properties_vec = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "rPr")
            .map(RPr::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { run_properties_vec })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentRun {
    pub p_contents: Vec<PContent>,
}

impl SdtContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtContentRun");

        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(PContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtRun {
    pub sdt_properties: Option<SdtPr>,
    pub sdt_end_properties: Option<SdtEndPr>,
    pub sdt_content: Option<SdtContentRun>,
}

impl SdtRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtRun");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sdtPr" => instance.sdt_properties = Some(SdtPr::from_xml_element(child_node)?),
                "sdtEndPr" => instance.sdt_end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                "sdtContent" => instance.sdt_content = Some(SdtContentRun::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Direction {
    #[strum(serialize = "ltr")]
    LeftToRight,
    #[strum(serialize = "rtl")]
    RightToLeft,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DirContentRun {
    pub p_contents: Vec<PContent>,
    pub value: Option<Direction>,
}

impl DirContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing DirContentRun");

        let value = xml_node.attributes.get("w:val").map(|val| val.parse()).transpose()?;

        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(PContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents, value })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct BdoContentRun {
    pub p_contents: Vec<PContent>,
    pub value: Option<Direction>,
}

impl BdoContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing BdoContentRun");

        let value = xml_node.attributes.get("w:val").map(|val| val.parse()).transpose()?;

        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(PContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents, value })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum BrType {
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "column")]
    Column,
    #[strum(serialzie = "textWrapping")]
    TextWrapping,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum BrClear {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "all")]
    All,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Br {
    pub break_type: Option<BrType>,
    pub clear: Option<BrClear>,
}

impl Br {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Br");

        let mut instance: Self = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => instance.break_type = Some(value.parse()?),
                "w:clear" => instance.clear = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Text {
    pub text: String,
    pub xml_space: Option<String>, // default or preserve
}

impl Text {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Text");

        let xml_space = xml_node.attributes.get("xml:space").cloned();

        let text = xml_node
            .text
            .as_ref()
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
            .clone();

        Ok(Self { text, xml_space })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Sym {
    pub font: Option<String>,
    pub character: Option<ShortHexNumber>,
}

impl Sym {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Sym");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:font" => instance.font = Some(value.clone()),
                "w:char" => instance.character = Some(ShortHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Control {
    pub name: Option<String>,
    pub shapeid: Option<String>,
    pub rel_id: Option<RelationshipId>,
}

impl Control {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:name" => instance.name = Some(value.clone()),
                "w:shapeid" => instance.shapeid = Some(value.clone()),
                "r:id" => instance.rel_id = Some(value.clone()),
                _ => (),
            }
        }

        instance
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ObjectDrawAspect {
    #[strum(serialize = "content")]
    Content,
    #[strum(serialize = "icon")]
    Icon,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ObjectEmbed {
    pub draw_aspect: Option<ObjectDrawAspect>,
    pub rel_id: RelationshipId,
    pub application_id: Option<String>,
    pub shape_id: Option<String>,
    pub field_codes: Option<String>,
}

impl ObjectEmbed {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ObjectEmbed");

        let mut draw_aspect = None;
        let mut rel_id = None;
        let mut application_id = None;
        let mut shape_id = None;
        let mut field_codes = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:drawAspect" => draw_aspect = Some(value.parse()?),
                "r:id" => rel_id = Some(value.clone()),
                "w:progId" => application_id = Some(value.clone()),
                "w:shapeId" => shape_id = Some(value.clone()),
                "w:fieldCodes" => field_codes = Some(value.clone()),
                _ => (),
            }
        }

        let rel_id = rel_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self {
            draw_aspect,
            rel_id,
            application_id,
            shape_id,
            field_codes,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ObjectUpdateMode {
    #[strum(serialize = "always")]
    Always,
    #[strum(serialize = "onCall")]
    OnCall,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ObjectLink {
    pub base: ObjectEmbed,
    pub update_mode: ObjectUpdateMode,
    pub locked_field: Option<OnOff>,
}

impl ObjectLink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ObjectLink");

        let base = ObjectEmbed::from_xml_element(xml_node)?;
        let mut update_mode = None;
        let mut locked_field = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:updateMode" => update_mode = Some(value.parse()?),
                "w:lockedField" => locked_field = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let update_mode = update_mode.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "updateMode"))?;

        Ok(Self {
            base,
            update_mode,
            locked_field,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ObjectChoice {
    Control(Control),
    ObjectLink(ObjectLink),
    ObjectEmbed(ObjectEmbed),
    Movie(Rel),
}

impl ObjectChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "control" | "objectLink" | "objectEmbed" | "movie" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ObjectChoice");

        match xml_node.local_name() {
            "control" => Ok(ObjectChoice::Control(Control::from_xml_element(xml_node))),
            "objectLink" => Ok(ObjectChoice::ObjectLink(ObjectLink::from_xml_element(xml_node)?)),
            "objectEmbed" => Ok(ObjectChoice::ObjectEmbed(ObjectEmbed::from_xml_element(xml_node)?)),
            "movie" => Ok(ObjectChoice::Movie(Rel::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ObjectChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DrawingChoice {
    Anchor(Anchor),
    Inline(Inline),
}

impl XsdType for DrawingChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "anchor" => Ok(DrawingChoice::Anchor(Anchor::from_xml_element(xml_node)?)),
            "inline" => Ok(DrawingChoice::Inline(Inline::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "DrawingChoice",
            ))),
        }
    }
}

impl XsdChoice for DrawingChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "anchor" | "inline" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Drawing(pub Vec<DrawingChoice>);

impl Drawing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Drawing");

        let anchor_or_inline_vec = xml_node
            .child_nodes
            .iter()
            .filter_map(DrawingChoice::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(anchor_or_inline_vec))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Object {
    pub drawing: Option<Drawing>,
    pub choice: Option<ObjectChoice>,
    pub original_image_width: Option<TwipsMeasure>,
    pub original_image_height: Option<TwipsMeasure>,
}

impl Object {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Object");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:dxaOrig" => instance.original_image_width = Some(value.parse()?),
                "w:dyaOrig" => instance.original_image_height = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "drawing" => instance.drawing = Some(Drawing::from_xml_element(child_node)?),
                node_name if ObjectChoice::is_choice_member(node_name) => {
                    instance.choice = Some(ObjectChoice::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum InfoTextType {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "autoText")]
    AutoText,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFHelpText {
    pub info_text_type: Option<InfoTextType>,
    pub value: Option<FFHelpTextVal>,
}

impl FFHelpText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFHelpText");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => instance.info_text_type = Some(value.parse()?),
                "w:val" => instance.value = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFStatusText {
    pub info_text_type: Option<InfoTextType>,
    pub value: Option<FFStatusTextVal>,
}

impl FFStatusText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFStatusText");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => instance.info_text_type = Some(value.parse()?),
                "w:val" => instance.value = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FFCheckBoxSizeChoice {
    Explicit(HpsMeasure),
    Auto(OnOff),
}

impl FFCheckBoxSizeChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "size" | "sizeAuto" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFCheckBoxSizeChoice");

        match xml_node.local_name() {
            "size" => Ok(FFCheckBoxSizeChoice::Explicit(HpsMeasure::from_xml_element(xml_node)?)),
            "sizeAuto" => Ok(FFCheckBoxSizeChoice::Auto(parse_on_off_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "FFCheckBoxSizeChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FFCheckBox {
    pub size: FFCheckBoxSizeChoice,
    pub is_default: Option<OnOff>,
    pub is_checked: Option<OnOff>,
}

impl FFCheckBox {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFCheckBox");

        let mut size = None;
        let mut is_default = None;
        let mut is_checked = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                node_name if FFCheckBoxSizeChoice::is_choice_member(node_name) => {
                    size = Some(FFCheckBoxSizeChoice::from_xml_element(child_node)?)
                }
                "default" => is_default = Some(parse_on_off_xml_element(child_node)?),
                "checked" => is_checked = Some(parse_on_off_xml_element(child_node)?),
                _ => (),
            }
        }

        let size = size.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "size|sizeAuto"))?;

        Ok(Self {
            size,
            is_default,
            is_checked,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFDDList {
    pub result: Option<DecimalNumber>,
    pub default: Option<DecimalNumber>,
    pub list_entries: Vec<String>,
}

impl FFDDList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFDDList");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "result" => instance.result = Some(child_node.get_val_attribute()?.parse()?),
                "default" => instance.default = Some(child_node.get_val_attribute()?.parse()?),
                "listEntry" => instance.list_entries.push(child_node.get_val_attribute()?.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum FFTextType {
    #[strum(serialize = "regular")]
    Regular,
    #[strum(serialize = "number")]
    Number,
    #[strum(serialize = "date")]
    Date,
    #[strum(serialize = "currentTime")]
    CurrentTime,
    #[strum(serialize = "currentDate")]
    CurrentDate,
    #[strum(serialize = "calculated")]
    Calculated,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFTextInput {
    pub text_type: Option<FFTextType>,
    pub default: Option<String>,
    pub max_length: Option<DecimalNumber>,
    pub format: Option<String>,
}

impl FFTextInput {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FFTextInput");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "type" => instance.text_type = Some(child_node.get_val_attribute()?.parse()?),
                "default" => instance.default = Some(child_node.get_val_attribute()?.clone()),
                "maxLength" => instance.max_length = Some(child_node.get_val_attribute()?.parse()?),
                "format" => instance.format = Some(child_node.get_val_attribute()?.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FFData {
    Name(FFName),
    Label(DecimalNumber),
    TabIndex(UnsignedDecimalNumber),
    Enabled(OnOff),
    RecalculateOnExit(OnOff),
    EntryMacro(MacroName),
    ExitMacro(MacroName),
    HelpText(FFHelpText),
    StatusText(FFStatusText),
    CheckBox(FFCheckBox),
    DropDownList(FFDDList),
    TextInput(FFTextInput),
}

impl XsdType for FFData {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "name" => Ok(FFData::Name(xml_node.get_val_attribute()?.clone())),
            "label" => Ok(FFData::Label(xml_node.get_val_attribute()?.parse()?)),
            "tabIndex" => Ok(FFData::TabIndex(xml_node.get_val_attribute()?.parse()?)),
            "enabled" => Ok(FFData::Enabled(parse_on_off_xml_element(xml_node)?)),
            "calcOnExit" => Ok(FFData::RecalculateOnExit(parse_on_off_xml_element(xml_node)?)),
            "entryMacro" => Ok(FFData::EntryMacro(xml_node.get_val_attribute()?.clone())),
            "exitMacro" => Ok(FFData::ExitMacro(xml_node.get_val_attribute()?.clone())),
            "helpText" => Ok(FFData::HelpText(FFHelpText::from_xml_element(xml_node)?)),
            "statusText" => Ok(FFData::StatusText(FFStatusText::from_xml_element(xml_node)?)),
            "checkBox" => Ok(FFData::CheckBox(FFCheckBox::from_xml_element(xml_node)?)),
            "ddList" => Ok(FFData::DropDownList(FFDDList::from_xml_element(xml_node)?)),
            "textInput" => Ok(FFData::TextInput(FFTextInput::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "FFData"))),
        }
    }
}

impl XsdChoice for FFData {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "name" | "label" | "tabIndex" | "enabled" | "calcOnExit" | "entryMacro" | "exitMacro" | "helpText"
            | "statusText" | "checkBox" | "ddList" | "textInput" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum FldCharType {
    #[strum(serialize = "begin")]
    Begin,
    #[strum(serialize = "separate")]
    Separate,
    #[strum(serialize = "end")]
    End,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FldChar {
    pub form_field_properties: Option<FFData>,
    pub field_char_type: FldCharType,
    pub field_lock: Option<OnOff>,
    pub dirty: Option<OnOff>,
}

impl FldChar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FldChar");

        let mut field_char_type = None;
        let mut field_lock = None;
        let mut dirty = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:fldCharType" => field_char_type = Some(value.parse()?),
                "w:fldLock" => field_lock = Some(parse_xml_bool(value)?),
                "w:dirty" => dirty = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let form_field_properties = xml_node
            .child_nodes
            .iter()
            .find_map(|child_node| FFData::try_from_xml_element(child_node))
            .transpose()?;

        let field_char_type =
            field_char_type.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "fldCharType"))?;

        Ok(Self {
            form_field_properties,
            field_char_type,
            field_lock,
            dirty,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum RubyAlign {
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "distributeLetter")]
    DistributeLetter,
    #[strum(serialize = "distributeSpace")]
    DistributeSpace,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "rightVertical")]
    RightVertical,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RubyPr {
    pub ruby_align: RubyAlign,
    pub hps: HpsMeasure,
    pub hps_raise: HpsMeasure,
    pub hps_base_text: HpsMeasure,
    pub language_id: Lang,
    pub dirty: Option<OnOff>,
}

impl RubyPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RubyPr");

        let mut ruby_align = None;
        let mut hps = None;
        let mut hps_raise = None;
        let mut hps_base_text = None;
        let mut language_id = None;
        let mut dirty = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rubyAlign" => ruby_align = Some(child_node.get_val_attribute()?.parse()?),
                "hps" => hps = Some(child_node.get_val_attribute()?.parse()?),
                "hpsRaise" => hps_raise = Some(child_node.get_val_attribute()?.parse()?),
                "hpsBaseText" => hps_base_text = Some(child_node.get_val_attribute()?.parse()?),
                "lid" => language_id = Some(child_node.get_val_attribute()?.clone()),
                "dirty" => dirty = Some(parse_on_off_xml_element(child_node)?),
                _ => (),
            }
        }

        let ruby_align = ruby_align.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyAlign"))?;
        let hps = hps.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hps"))?;
        let hps_raise = hps_raise.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hpsRaise"))?;
        let hps_base_text =
            hps_base_text.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hpsBaseText"))?;
        let language_id = language_id.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lid"))?;

        Ok(Self {
            ruby_align,
            hps,
            hps_raise,
            hps_base_text,
            language_id,
            dirty,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RubyContentChoice {
    Run(R),
    RunLevelElement(RunLevelElts),
}

impl XsdType for RubyContentChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "r" => Ok(RubyContentChoice::Run(R::from_xml_element(xml_node)?)),
            node_name if RunLevelElts::is_choice_member(node_name) => Ok(RubyContentChoice::RunLevelElement(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RubyContentChoice",
            ))),
        }
    }
}

impl XsdChoice for RubyContentChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "r" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RubyContent {
    pub ruby_contents: Vec<RubyContentChoice>,
}

impl RubyContent {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RubyContent");

        let ruby_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(RubyContentChoice::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { ruby_contents })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Ruby {
    pub ruby_properties: RubyPr,
    pub ruby_content: RubyContent,
    pub ruby_base: RubyContent,
}

impl Ruby {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Ruby");

        let mut ruby_properties = None;
        let mut ruby_content = None;
        let mut ruby_base = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rubyPr" => ruby_properties = Some(RubyPr::from_xml_element(child_node)?),
                "rt" => ruby_content = Some(RubyContent::from_xml_element(child_node)?),
                "rubyBase" => ruby_base = Some(RubyContent::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let ruby_properties =
            ruby_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyPr"))?;
        let ruby_content = ruby_content.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rt"))?;
        let ruby_base = ruby_base.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyBase"))?;

        Ok(Self {
            ruby_properties,
            ruby_content,
            ruby_base,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FtnEdnRef {
    pub custom_mark_follows: Option<OnOff>,
    pub id: DecimalNumber,
}

impl FtnEdnRef {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FtnEdnRef");

        let mut custom_mark_follows = None;
        let mut id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:customMarkFollows" => custom_mark_follows = Some(parse_xml_bool(value)?),
                "w:id" => id = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            custom_mark_follows,
            id: id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PTabAlignment {
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "right")]
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PTabRelativeTo {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "indent")]
    Indent,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PTabLeader {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "underscore")]
    Underscore,
    #[strum(serialize = "middleDot")]
    MiddleDot,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PTab {
    pub alignment: PTabAlignment,
    pub relative_to: PTabRelativeTo,
    pub leader: PTabLeader,
}

impl PTab {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PTab");

        let mut alignment = None;
        let mut relative_to = None;
        let mut leader = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:alignment" => alignment = Some(value.parse()?),
                "w:relativeTo" => relative_to = Some(value.parse()?),
                "w:leader" => leader = Some(value.parse()?),
                _ => (),
            }
        }

        let alignment = alignment.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "alignment"))?;
        let relative_to = relative_to.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "relativeTo"))?;
        let leader = leader.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "leader"))?;

        Ok(Self {
            alignment,
            relative_to,
            leader,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunInnerContent {
    Break(Br),
    Text(Text),
    ContentPart(Rel),
    DeletedText(Text),
    InstructionText(Text),
    DeletedInstructionText(Text),
    NonBreakingHyphen,
    OptionalHypen,
    ShortDayFormat,
    ShortMonthFormat,
    ShortYearFormat,
    LongDayFormat,
    LongMonthFormat,
    LongYearFormat,
    AnnorationReferenceMark,
    FootnoteReferenceMark,
    EndnoteReferenceMark,
    Separator,
    ContinuationSeparator,
    Symbol(Sym),
    PageNum,
    CarriageReturn,
    Tab,
    Object(Object),
    FieldCharacter(FldChar),
    Ruby(Ruby),
    FootnoteReference(FtnEdnRef),
    EndnoteReference(FtnEdnRef),
    CommentReference(Markup),
    Drawing(Drawing),
    PositionTab(PTab),
    LastRenderedPageBreak,
}

impl RunInnerContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "br"
            | "t"
            | "contentPart"
            | "delText"
            | "instrText"
            | "delInstrText"
            | "noBreakHyphen"
            | "softHyphen"
            | "dayShort"
            | "monthShort"
            | "yearShort"
            | "dayLong"
            | "monthLong"
            | "yearLong"
            | "annotationRef"
            | "footnoteRef"
            | "endnoteRef"
            | "separator"
            | "continuationSeparator"
            | "sym"
            | "pgNum"
            | "cr"
            | "tab"
            | "object"
            | "fldChar"
            | "ruby"
            | "footnoteReference"
            | "endnoteReference"
            | "commentReference"
            | "drawing"
            | "ptab"
            | "lastRenderedPageBreak" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RunInnerContent");

        match xml_node.local_name() {
            "br" => Ok(RunInnerContent::Break(Br::from_xml_element(xml_node)?)),
            "t" => Ok(RunInnerContent::Text(Text::from_xml_element(xml_node)?)),
            "contentPart" => Ok(RunInnerContent::ContentPart(Rel::from_xml_element(xml_node)?)),
            "delText" => Ok(RunInnerContent::DeletedText(Text::from_xml_element(xml_node)?)),
            "instrText" => Ok(RunInnerContent::InstructionText(Text::from_xml_element(xml_node)?)),
            "delInstrText" => Ok(RunInnerContent::DeletedInstructionText(Text::from_xml_element(
                xml_node,
            )?)),
            "noBreakHyphen" => Ok(RunInnerContent::NonBreakingHyphen),
            "softHyphen" => Ok(RunInnerContent::OptionalHypen),
            "dayShort" => Ok(RunInnerContent::ShortDayFormat),
            "monthShort" => Ok(RunInnerContent::ShortMonthFormat),
            "yearShort" => Ok(RunInnerContent::ShortYearFormat),
            "dayLong" => Ok(RunInnerContent::LongDayFormat),
            "monthLong" => Ok(RunInnerContent::LongMonthFormat),
            "yearLong" => Ok(RunInnerContent::LongYearFormat),
            "annotationRef" => Ok(RunInnerContent::AnnorationReferenceMark),
            "footnoteRef" => Ok(RunInnerContent::FootnoteReferenceMark),
            "endnoteRef" => Ok(RunInnerContent::EndnoteReferenceMark),
            "separator" => Ok(RunInnerContent::Separator),
            "continuationSeparator" => Ok(RunInnerContent::ContinuationSeparator),
            "sym" => Ok(RunInnerContent::Symbol(Sym::from_xml_element(xml_node)?)),
            "pgNum" => Ok(RunInnerContent::PageNum),
            "cr" => Ok(RunInnerContent::CarriageReturn),
            "tab" => Ok(RunInnerContent::Tab),
            "object" => Ok(RunInnerContent::Object(Object::from_xml_element(xml_node)?)),
            "fldChar" => Ok(RunInnerContent::FieldCharacter(FldChar::from_xml_element(xml_node)?)),
            "ruby" => Ok(RunInnerContent::Ruby(Ruby::from_xml_element(xml_node)?)),
            "footnoteReference" => Ok(RunInnerContent::FootnoteReference(FtnEdnRef::from_xml_element(
                xml_node,
            )?)),
            "endnoteReference" => Ok(RunInnerContent::EndnoteReference(FtnEdnRef::from_xml_element(
                xml_node,
            )?)),
            "commentReference" => Ok(RunInnerContent::CommentReference(Markup::from_xml_element(xml_node)?)),
            "drawing" => Ok(RunInnerContent::Drawing(Drawing::from_xml_element(xml_node)?)),
            "ptab" => Ok(RunInnerContent::PositionTab(PTab::from_xml_element(xml_node)?)),
            "lastRenderedPageBreak" => Ok(RunInnerContent::LastRenderedPageBreak),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunInnerContent",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct R {
    pub run_properties: Option<RPr>,
    pub run_inner_contents: Vec<RunInnerContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
}

impl R {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing R");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                node_name if RunInnerContent::is_choice_member(node_name) => instance
                    .run_inner_contents
                    .push(RunInnerContent::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ContentRunContent {
    CustomXml(CustomXmlRun),
    SmartTag(SmartTagRun),
    Sdt(Box<SdtRun>),
    Bidirectional(DirContentRun),
    BidirectionalOverride(BdoContentRun),
    Run(R),
    RunLevelElements(RunLevelElts),
}

impl ContentRunContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "customXml" | "smartTag" | "sdt" | "dir" | "bdo" | "r" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ContentRunContent");

        match xml_node.local_name() {
            "customXml" => Ok(ContentRunContent::CustomXml(CustomXmlRun::from_xml_element(xml_node)?)),
            "smartTag" => Ok(ContentRunContent::SmartTag(SmartTagRun::from_xml_element(xml_node)?)),
            "sdt" => Ok(ContentRunContent::Sdt(Box::new(SdtRun::from_xml_element(xml_node)?))),
            "dir" => Ok(ContentRunContent::Bidirectional(DirContentRun::from_xml_element(
                xml_node,
            )?)),
            "bdo" => Ok(ContentRunContent::BidirectionalOverride(
                BdoContentRun::from_xml_element(xml_node)?,
            )),
            "r" => Ok(ContentRunContent::Run(R::from_xml_element(xml_node)?)),
            node_name if RunLevelElts::is_choice_member(node_name) => Ok(ContentRunContent::RunLevelElements(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentRunContent",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunTrackChangeChoice {
    ContentRunContent(ContentRunContent),
    // TODO
    // OMathMathElements(OMathMathElements),
}

impl XsdType for RunTrackChangeChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let local_name = xml_node.local_name();
        if ContentRunContent::is_choice_member(local_name) {
            Ok(RunTrackChangeChoice::ContentRunContent(
                ContentRunContent::from_xml_element(xml_node)?,
            ))
        } else {
            Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunTrackChangeChoice",
            )))
        }
    }
}

impl XsdChoice for RunTrackChangeChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        ContentRunContent::is_choice_member(node_name)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RunTrackChange {
    pub base: TrackChange,
    pub choices: Vec<RunTrackChangeChoice>,
}

impl RunTrackChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RunTrackChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let choices = xml_node
            .child_nodes
            .iter()
            .filter_map(RunTrackChangeChoice::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { base, choices })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RangeMarkupElements {
    BookmarkStart(Bookmark),
    BookmarkEnd(MarkupRange),
    MoveFromRangeStart(MoveBookmark),
    MoveFromRangeEnd(MarkupRange),
    MoveToRangeStart(MoveBookmark),
    MoveToRangeEnd(MarkupRange),
    CommentRangeStart(MarkupRange),
    CommentRangeEnd(MarkupRange),
    CustomXmlInsertRangeStart(TrackChange),
    CustomXmlInsertRangeEnd(Markup),
    CustomXmlDeleteRangeStart(TrackChange),
    CustomXmlDeleteRangeEnd(Markup),
    CustomXmlMoveFromRangeStart(TrackChange),
    CustomXmlMoveFromRangeEnd(Markup),
    CustomXmlMoveToRangeStart(TrackChange),
    CustomXmlMoveToRangeEnd(Markup),
}

impl RangeMarkupElements {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "bookmarkStart"
            | "bookmarkEnd"
            | "moveFromRangeStart"
            | "moveFromRangeEnd"
            | "moveToRangeStart"
            | "moveToRangeEnd"
            | "commentRangeStart"
            | "commentRangeEnd"
            | "customXmlInsRangeStart"
            | "customXmlInsRangeEnd"
            | "customXmlDelRangeStart"
            | "customXmlDelRangeEnd"
            | "customXmlMoveFromRangeStart"
            | "customXmlMoveFromRangeEnd"
            | "customXmlMoveToRangeStart"
            | "customXmlMoveToRangeEnd" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RangeMarkupElements");

        match xml_node.local_name() {
            "bookmarkStart" => Ok(RangeMarkupElements::BookmarkStart(Bookmark::from_xml_element(
                xml_node,
            )?)),
            "bookmarkEnd" => Ok(RangeMarkupElements::BookmarkEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "moveFromRangeStart" => Ok(RangeMarkupElements::MoveFromRangeStart(MoveBookmark::from_xml_element(
                xml_node,
            )?)),
            "moveFromRangeEnd" => Ok(RangeMarkupElements::MoveFromRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "moveToRangeStart" => Ok(RangeMarkupElements::MoveToRangeStart(MoveBookmark::from_xml_element(
                xml_node,
            )?)),
            "moveToRangeEnd" => Ok(RangeMarkupElements::MoveToRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "commentRangeStart" => Ok(RangeMarkupElements::CommentRangeStart(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "commentRangeEnd" => Ok(RangeMarkupElements::CommentRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "customXmlInsRangeStart" => Ok(RangeMarkupElements::CustomXmlInsertRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlInsRangeEnd" => Ok(RangeMarkupElements::CustomXmlInsertRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            "customXmlDelRangeStart" => Ok(RangeMarkupElements::CustomXmlDeleteRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlDelRangeEnd" => Ok(RangeMarkupElements::CustomXmlDeleteRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            "customXmlMoveFromRangeStart" => Ok(RangeMarkupElements::CustomXmlMoveFromRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlMoveFromRangeEnd" => Ok(RangeMarkupElements::CustomXmlMoveFromRangeEnd(
                Markup::from_xml_element(xml_node)?,
            )),
            "customXmlMoveToRangeStart" => Ok(RangeMarkupElements::CustomXmlMoveToRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlMoveToRangeEnd" => Ok(RangeMarkupElements::CustomXmlMoveToRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RangeMarkupElements",
            ))),
        }
    }
}

// TODO
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MathContent {
    // OMathParagraph(OMathParagraph),
// OMath(OMath),
}

impl MathContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "oMathPara" | "oMath" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunLevelElts {
    ProofError(ProofErr),
    PermissionStart(PermStart),
    PermissionEnd(Perm),
    RangeMarkupElements(RangeMarkupElements),
    Insert(RunTrackChange),
    Delete(RunTrackChange),
    MoveFrom(RunTrackChange),
    MoveTo(RunTrackChange),
    MathContent(MathContent),
}

impl RunLevelElts {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "proofErr" | "permStart" | "permEnd" | "ins" | "del" | "moveFrom" | "moveTo" => true,
            _ => RangeMarkupElements::is_choice_member(&node_name) || MathContent::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing RunLevelElts");

        let local_name = xml_node.local_name();
        match local_name {
            "proofErr" => Ok(RunLevelElts::ProofError(ProofErr::from_xml_element(xml_node)?)),
            "permStart" => Ok(RunLevelElts::PermissionStart(PermStart::from_xml_element(xml_node)?)),
            "permEnd" => Ok(RunLevelElts::PermissionEnd(Perm::from_xml_element(xml_node)?)),
            "ins" => Ok(RunLevelElts::Insert(RunTrackChange::from_xml_element(xml_node)?)),
            "del" => Ok(RunLevelElts::Delete(RunTrackChange::from_xml_element(xml_node)?)),
            "moveFrom" => Ok(RunLevelElts::MoveFrom(RunTrackChange::from_xml_element(xml_node)?)),
            "moveTo" => Ok(RunLevelElts::MoveTo(RunTrackChange::from_xml_element(xml_node)?)),
            _ if RangeMarkupElements::is_choice_member(local_name) => Ok(RunLevelElts::RangeMarkupElements(
                RangeMarkupElements::from_xml_element(xml_node)?,
            )),
            // TODO MathContent
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunLevelElts",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlBlock {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub block_contents: Vec<ContentBlockContent>,
    pub uri: Option<String>,
    pub element: XmlName,
}

impl CustomXmlBlock {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CustomXmlBlock");

        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(value.clone()),
                "w:element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let mut custom_xml_properties = None;
        let mut block_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name if ContentBlockContent::is_choice_member(node_name) => {
                    block_contents.push(ContentBlockContent::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        Ok(Self {
            custom_xml_properties,
            block_contents,
            uri,
            element,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentBlock {
    pub block_contents: Vec<ContentBlockContent>,
}

impl SdtContentBlock {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtContentBlock");

        let block_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(ContentBlockContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { block_contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtBlock {
    pub sdt_properties: Option<SdtPr>,
    pub sdt_end_properties: Option<SdtEndPr>,
    pub sdt_content: Option<SdtContentBlock>,
}

impl SdtBlock {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtBlock");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sdtPr" => instance.sdt_properties = Some(SdtPr::from_xml_element(child_node)?),
                "sdtEndPr" => instance.sdt_end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                "sdtContent" => instance.sdt_content = Some(SdtContentBlock::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum DropCap {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "drop")]
    Drop,
    #[strum(serialize = "margin")]
    Margin,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum HeightRule {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "exact")]
    Exact,
    #[strum(serialize = "atLeast")]
    AtLeast,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Wrap {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "notBeside")]
    NotBeside,
    #[strum(serialize = "around")]
    Around,
    #[strum(serialize = "tight")]
    Tight,
    #[strum(serialize = "through")]
    Throught,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum VAnchor {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum HAnchor {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct FramePr {
    pub drop_cap: Option<DropCap>,
    pub lines: Option<DecimalNumber>,
    pub width: Option<TwipsMeasure>,
    pub height: Option<TwipsMeasure>,
    pub vertical_space: Option<TwipsMeasure>,
    pub horizontal_space: Option<TwipsMeasure>,
    pub wrap: Option<Wrap>,
    pub horizontal_anchor: Option<HAnchor>,
    pub vertical_anchor: Option<VAnchor>,
    pub x: Option<SignedTwipsMeasure>,
    pub x_align: Option<XAlign>,
    pub y: Option<SignedTwipsMeasure>,
    pub y_align: Option<YAlign>,
    pub height_rule: Option<HeightRule>,
    pub anchor_lock: Option<OnOff>,
}

impl FramePr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FramePr");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:dropCap" => instance.drop_cap = Some(value.parse()?),
                "w:lines" => instance.lines = Some(value.parse()?),
                "w:w" => instance.width = Some(value.parse()?),
                "w:h" => instance.height = Some(value.parse()?),
                "w:vSpace" => instance.vertical_space = Some(value.parse()?),
                "w:hSpace" => instance.horizontal_space = Some(value.parse()?),
                "w:wrap" => instance.wrap = Some(value.parse()?),
                "w:hAnchor" => instance.horizontal_anchor = Some(value.parse()?),
                "w:vAnchor" => instance.vertical_anchor = Some(value.parse()?),
                "w:x" => instance.x = Some(value.parse()?),
                "w:xAlign" => instance.x_align = Some(value.parse()?),
                "w:y" => instance.y = Some(value.parse()?),
                "w:yAlign" => instance.y_align = Some(value.parse()?),
                "w:hRule" => instance.height_rule = Some(value.parse()?),
                "w:anchorLock" => instance.anchor_lock = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for FramePr {
    fn update_with(self, other: Self) -> Self {
        Self {
            drop_cap: other.drop_cap.or(self.drop_cap),
            lines: other.lines.or(self.lines),
            width: other.width.or(self.width),
            height: other.height.or(self.height),
            vertical_space: other.vertical_space.or(self.vertical_space),
            horizontal_space: other.horizontal_space.or(self.horizontal_space),
            wrap: other.wrap.or(self.wrap),
            horizontal_anchor: other.horizontal_anchor.or(self.horizontal_anchor),
            vertical_anchor: other.vertical_anchor.or(self.vertical_anchor),
            x: other.x.or(self.x),
            x_align: other.x_align.or(self.x_align),
            y: other.y.or(self.y),
            y_align: other.y_align.or(self.y_align),
            height_rule: other.height_rule.or(self.height_rule),
            anchor_lock: other.anchor_lock.or(self.anchor_lock),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct NumPr {
    pub indent_level: Option<DecimalNumber>,
    pub numbering_id: Option<DecimalNumber>,
    pub inserted: Option<TrackChange>,
}

impl NumPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing NumPr");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "ilvl" => instance.indent_level = Some(child_node.get_val_attribute()?.parse()?),
                    "numId" => instance.numbering_id = Some(child_node.get_val_attribute()?.parse()?),
                    "ins" => instance.inserted = Some(TrackChange::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

impl Update for NumPr {
    fn update_with(self, other: Self) -> Self {
        Self {
            indent_level: other.indent_level.or(self.indent_level),
            numbering_id: other.numbering_id.or(self.numbering_id),
            inserted: other.inserted.or(self.inserted),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct PBdr {
    pub top: Option<Border>,
    pub left: Option<Border>,
    pub bottom: Option<Border>,
    pub right: Option<Border>,
    pub between: Option<Border>,
    pub bar: Option<Border>,
}

impl PBdr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PBdr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "left" => instance.left = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "right" => instance.right = Some(Border::from_xml_element(child_node)?),
                "between" => instance.between = Some(Border::from_xml_element(child_node)?),
                "bar" => instance.bar = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for PBdr {
    fn update_with(self, other: Self) -> Self {
        Self {
            top: update_options(self.top, other.top),
            left: update_options(self.left, other.left),
            bottom: update_options(self.bottom, other.bottom),
            right: update_options(self.right, other.right),
            between: update_options(self.between, other.between),
            bar: update_options(self.bar, other.bar),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TabJc {
    #[strum(serialize = "clear")]
    Clear,
    #[strum(serialize = "start")]
    Start,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "decimal")]
    Decimal,
    #[strum(serialize = "bar")]
    Bar,
    #[strum(serialize = "num")]
    Number,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TabTlc {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "underscore")]
    Underscore,
    #[strum(serialize = "heavy")]
    Heavy,
    #[strum(serialize = "middleDot")]
    MiddleDot,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TabStop {
    pub value: TabJc,
    pub leader: Option<TabTlc>,
    pub position: SignedTwipsMeasure,
}

impl TabStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TabStop");

        let mut value = None;
        let mut leader = None;
        let mut position = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:leader" => leader = Some(attr_value.parse()?),
                "w:pos" => position = Some(attr_value.parse()?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        let position = position.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "pos"))?;

        Ok(Self {
            value,
            leader,
            position,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Tabs(pub Vec<TabStop>);

impl Tabs {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Tabs");

        let tabs = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "tab")
            .map(TabStop::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        if tabs.is_empty() {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "tab",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        } else {
            Ok(Self(tabs))
        }
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum LineSpacingRule {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "exact")]
    Exact,
    #[strum(serialize = "atLeast")]
    AtLeast,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Spacing {
    pub before: Option<TwipsMeasure>,
    pub before_lines: Option<DecimalNumber>,
    pub before_autospacing: Option<OnOff>,
    pub after: Option<TwipsMeasure>,
    pub after_lines: Option<DecimalNumber>,
    pub after_autospacing: Option<OnOff>,
    pub line: Option<SignedTwipsMeasure>,
    pub line_rule: Option<LineSpacingRule>,
}

impl Spacing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Spacing");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:before" => instance.before = Some(value.parse()?),
                    "w:beforeLines" => instance.before_lines = Some(value.parse()?),
                    "w:beforeAutospacing" => instance.before_autospacing = Some(parse_xml_bool(value)?),
                    "w:after" => instance.after = Some(value.parse()?),
                    "w:afterLines" => instance.after_lines = Some(value.parse()?),
                    "w:afterAutospacing" => instance.after_autospacing = Some(parse_xml_bool(value)?),
                    "w:line" => instance.line = Some(value.parse()?),
                    "w:lineRule" => instance.line_rule = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

impl Update for Spacing {
    fn update_with(self, other: Self) -> Self {
        Self {
            before: other.before.or(self.before),
            before_lines: other.before_lines.or(self.before_lines),
            before_autospacing: other.before_autospacing.or(self.before_autospacing),
            after: other.after.or(self.after),
            after_lines: other.after_lines.or(self.after_lines),
            after_autospacing: other.after_autospacing.or(self.after_autospacing),
            line: other.line.or(self.line),
            line_rule: other.line_rule.or(self.line_rule),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Ind {
    pub start: Option<SignedTwipsMeasure>,
    pub start_chars: Option<DecimalNumber>,
    pub end: Option<SignedTwipsMeasure>,
    pub end_chars: Option<DecimalNumber>,
    // Deprecated
    pub left: Option<SignedTwipsMeasure>,
    // Deprecated
    pub left_chars: Option<DecimalNumber>,
    // Deprecated
    pub right: Option<SignedTwipsMeasure>,
    // Deprecated
    pub right_chars: Option<DecimalNumber>,
    pub hanging: Option<TwipsMeasure>,
    pub hanging_chars: Option<DecimalNumber>,
    pub first_line: Option<TwipsMeasure>,
    pub first_line_chars: Option<DecimalNumber>,
}

impl Ind {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Ind");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:start" => instance.start = Some(value.parse()?),
                "w:startChars" => instance.start_chars = Some(value.parse()?),
                "w:end" => instance.end = Some(value.parse()?),
                "w:endChars" => instance.end_chars = Some(value.parse()?),
                "w:left" => instance.left = Some(value.parse()?),
                "w:leftChars" => instance.left_chars = Some(value.parse()?),
                "w:right" => instance.right = Some(value.parse()?),
                "w:rightChars" => instance.right_chars = Some(value.parse()?),
                "w:hanging" => instance.hanging = Some(value.parse()?),
                "w:hangingChars" => instance.hanging_chars = Some(value.parse()?),
                "w:firstLine" => instance.first_line = Some(value.parse()?),
                "w:firstLineChars" => instance.first_line_chars = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for Ind {
    fn update_with(self, other: Self) -> Self {
        Self {
            start: other.start.or(self.start),
            start_chars: other.start_chars.or(self.start_chars),
            end: other.end.or(self.end),
            end_chars: other.end_chars.or(self.end_chars),
            left: other.left.or(self.left),
            left_chars: other.left_chars.or(self.left_chars),
            right: other.right.or(self.right),
            right_chars: other.right_chars.or(self.right_chars),
            hanging: other.hanging.or(self.hanging),
            hanging_chars: other.hanging_chars.or(self.hanging_chars),
            first_line: other.first_line.or(self.first_line),
            first_line_chars: other.first_line_chars.or(self.first_line_chars),
        }
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Jc {
    #[strum(serialize = "start")]
    Start,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "both")]
    Both,
    #[strum(serialize = "mediumKashida")]
    MediumKashida,
    #[strum(serialize = "distribute")]
    Distribute,
    #[strum(serialize = "numTab")]
    NumTab,
    #[strum(serialize = "highKashida")]
    HighKashida,
    #[strum(serialize = "lowKashida")]
    LowKashida,
    #[strum(serialize = "thaiDistribute")]
    ThaiDistribute,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TextDirection {
    #[strum(serialize = "tb")]
    TopToBottom,
    #[strum(serialize = "rl")]
    RightToLeft,
    #[strum(serialize = "lr")]
    LeftToRight,
    #[strum(serialize = "tbV")]
    TopToBottomRotated,
    #[strum(serialize = "rlV")]
    RightToLeftRotated,
    #[strum(serialize = "lrV")]
    LeftToRightRotated,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TextAlignment {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "baseline")]
    Baseline,
    #[strum(serialize = "bottom")]
    Bottom,
    #[strum(serialize = "auto")]
    Auto,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TextboxTightWrap {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "allLines")]
    AllLines,
    #[strum(serialize = "firstAndLastLine")]
    FirstAndLastLine,
    #[strum(serialize = "firstLineOnly")]
    FirstLineOnly,
    #[strum(serialize = "lastLineOnly")]
    LastLineOnly,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Cnf {
    pub first_row: Option<OnOff>,
    pub last_row: Option<OnOff>,
    pub first_column: Option<OnOff>,
    pub last_column: Option<OnOff>,
    pub odd_vertical_band: Option<OnOff>,
    pub even_vertical_band: Option<OnOff>,
    pub odd_horizontal_band: Option<OnOff>,
    pub even_horizontal_band: Option<OnOff>,
    pub first_row_first_column: Option<OnOff>,
    pub first_row_last_column: Option<OnOff>,
    pub last_row_first_column: Option<OnOff>,
    pub last_row_last_column: Option<OnOff>,
}

impl Cnf {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Cnf");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:firstRow" => instance.first_row = Some(parse_xml_bool(value)?),
                "w:lastRow" => instance.last_row = Some(parse_xml_bool(value)?),
                "w:firstColumn" => instance.first_column = Some(parse_xml_bool(value)?),
                "w:lastColumn" => instance.last_column = Some(parse_xml_bool(value)?),
                "w:oddVBand" => instance.odd_vertical_band = Some(parse_xml_bool(value)?),
                "w:evenVBand" => instance.even_vertical_band = Some(parse_xml_bool(value)?),
                "w:oddHBand" => instance.odd_horizontal_band = Some(parse_xml_bool(value)?),
                "w:evenHBand" => instance.even_horizontal_band = Some(parse_xml_bool(value)?),
                "w:firstRowFirstColumn" => instance.first_row_first_column = Some(parse_xml_bool(value)?),
                "w:firstRowLastColumn" => instance.first_row_last_column = Some(parse_xml_bool(value)?),
                "w:lastRowFirstColumn" => instance.last_row_first_column = Some(parse_xml_bool(value)?),
                "w:lastRowLastColumn" => instance.last_row_last_column = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

impl Update for Cnf {
    fn update_with(self, other: Self) -> Self {
        Self {
            first_row: other.first_column.or(self.first_row),
            last_row: other.last_row.or(self.last_row),
            first_column: other.first_column.or(self.first_column),
            last_column: other.last_column.or(self.last_column),
            odd_vertical_band: other.odd_vertical_band.or(self.odd_vertical_band),
            even_vertical_band: other.even_vertical_band.or(self.even_vertical_band),
            odd_horizontal_band: other.odd_horizontal_band.or(self.odd_horizontal_band),
            even_horizontal_band: other.even_horizontal_band.or(self.even_horizontal_band),
            first_row_first_column: other.first_row_first_column.or(self.first_row_first_column),
            first_row_last_column: other.first_row_last_column.or(self.first_row_last_column),
            last_row_first_column: other.last_row_first_column.or(self.last_row_first_column),
            last_row_last_column: other.last_row_last_column.or(self.last_row_last_column),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPrBase {
    pub style: Option<String>,
    pub keep_with_next: Option<OnOff>,
    pub keep_lines_on_one_page: Option<OnOff>,
    pub start_on_next_page: Option<OnOff>,
    pub frame_properties: Option<FramePr>,
    pub widow_control: Option<OnOff>,
    pub numbering_properties: Option<NumPr>,
    pub suppress_line_numbers: Option<OnOff>,
    pub borders: Option<PBdr>,
    pub shading: Option<Shd>,
    pub tabs: Option<Tabs>,
    pub suppress_auto_hyphens: Option<OnOff>,
    pub kinsoku: Option<OnOff>,
    pub word_wrapping: Option<OnOff>,
    pub overflow_punctuations: Option<OnOff>,
    pub top_line_punctuations: Option<OnOff>,
    pub auto_space_latin_and_east_asian: Option<OnOff>,
    pub auto_space_east_asian_and_numbers: Option<OnOff>,
    pub bidirectional: Option<OnOff>,
    pub adjust_right_indent: Option<OnOff>,
    pub snap_to_grid: Option<OnOff>,
    pub spacing: Option<Spacing>,
    pub indent: Option<Ind>,
    pub contextual_spacing: Option<OnOff>,
    pub mirror_indents: Option<OnOff>,
    pub suppress_overlapping: Option<OnOff>,
    pub alignment: Option<Jc>,
    pub text_direction: Option<TextDirection>,
    pub text_alignment: Option<TextAlignment>,
    pub textbox_tight_wrap: Option<TextboxTightWrap>,
    pub outline_level: Option<DecimalNumber>,
    pub div_id: Option<DecimalNumber>,
    pub conditional_formatting: Option<Cnf>,
}

impl PPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PPrBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "pStyle" => self.style = Some(xml_node.get_val_attribute()?.clone()),
            "keepNext" => self.keep_with_next = Some(parse_on_off_xml_element(xml_node)?),
            "keepLines" => self.keep_lines_on_one_page = Some(parse_on_off_xml_element(xml_node)?),
            "pageBreakBefore" => self.start_on_next_page = Some(parse_on_off_xml_element(xml_node)?),
            "framePr" => self.frame_properties = Some(FramePr::from_xml_element(xml_node)?),
            "widowControl" => self.widow_control = Some(parse_on_off_xml_element(xml_node)?),
            "numPr" => self.numbering_properties = Some(NumPr::from_xml_element(xml_node)?),
            "suppressLineNumbers" => self.suppress_line_numbers = Some(parse_on_off_xml_element(xml_node)?),
            "pBdr" => self.borders = Some(PBdr::from_xml_element(xml_node)?),
            "shd" => self.shading = Some(Shd::from_xml_element(xml_node)?),
            "tabs" => self.tabs = Some(Tabs::from_xml_element(xml_node)?),
            "suppressAutoHyphens" => self.suppress_auto_hyphens = Some(parse_on_off_xml_element(xml_node)?),
            "kinsoku" => self.kinsoku = Some(parse_on_off_xml_element(xml_node)?),
            "wordWrap" => self.word_wrapping = Some(parse_on_off_xml_element(xml_node)?),
            "overflowPunct" => self.overflow_punctuations = Some(parse_on_off_xml_element(xml_node)?),
            "topLinePunct" => self.top_line_punctuations = Some(parse_on_off_xml_element(xml_node)?),
            "autoSpaceDE" => self.auto_space_latin_and_east_asian = Some(parse_on_off_xml_element(xml_node)?),
            "autoSpaceDN" => self.auto_space_east_asian_and_numbers = Some(parse_on_off_xml_element(xml_node)?),
            "bidi" => self.bidirectional = Some(parse_on_off_xml_element(xml_node)?),
            "adjustRightInd" => self.adjust_right_indent = Some(parse_on_off_xml_element(xml_node)?),
            "snapToGrid" => self.snap_to_grid = Some(parse_on_off_xml_element(xml_node)?),
            "spacing" => self.spacing = Some(Spacing::from_xml_element(xml_node)?),
            "ind" => self.indent = Some(Ind::from_xml_element(xml_node)?),
            "contextualSpacing" => self.contextual_spacing = Some(parse_on_off_xml_element(xml_node)?),
            "mirrorIndents" => self.mirror_indents = Some(parse_on_off_xml_element(xml_node)?),
            "suppressOverlap" => self.suppress_overlapping = Some(parse_on_off_xml_element(xml_node)?),
            "jc" => self.alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "textDirection" => self.text_direction = Some(xml_node.get_val_attribute()?.parse()?),
            "textAlignment" => self.text_alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "textboxTightWrap" => self.textbox_tight_wrap = Some(xml_node.get_val_attribute()?.parse()?),
            "outlineLvl" => self.outline_level = Some(xml_node.get_val_attribute()?.parse()?),
            "divId" => self.div_id = Some(xml_node.get_val_attribute()?.parse()?),
            "cnfStyle" => self.conditional_formatting = Some(Cnf::from_xml_element(xml_node)?),
            _ => (),
        }

        Ok(self)
    }
}

impl Update for PPrBase {
    fn update_with(self, other: Self) -> Self {
        Self {
            style: other.style.or(self.style),
            keep_with_next: other.keep_with_next.or(self.keep_with_next),
            keep_lines_on_one_page: other.keep_lines_on_one_page.or(self.keep_lines_on_one_page),
            start_on_next_page: other.start_on_next_page.or(self.start_on_next_page),
            frame_properties: update_options(self.frame_properties, other.frame_properties),
            widow_control: other.widow_control.or(self.widow_control),
            numbering_properties: update_options(self.numbering_properties, other.numbering_properties),
            suppress_line_numbers: other.suppress_line_numbers.or(self.suppress_line_numbers),
            borders: update_options(self.borders, other.borders),
            shading: update_options(self.shading, other.shading),
            tabs: other.tabs.or(self.tabs),
            suppress_auto_hyphens: other.suppress_auto_hyphens.or(self.suppress_auto_hyphens),
            kinsoku: other.kinsoku.or(self.kinsoku),
            word_wrapping: other.word_wrapping.or(self.word_wrapping),
            overflow_punctuations: other.overflow_punctuations.or(self.overflow_punctuations),
            top_line_punctuations: other.top_line_punctuations.or(self.top_line_punctuations),
            auto_space_latin_and_east_asian: other
                .auto_space_latin_and_east_asian
                .or(self.auto_space_latin_and_east_asian),
            auto_space_east_asian_and_numbers: other
                .auto_space_east_asian_and_numbers
                .or(self.auto_space_east_asian_and_numbers),
            bidirectional: other.bidirectional.or(self.bidirectional),
            adjust_right_indent: other.adjust_right_indent.or(self.adjust_right_indent),
            snap_to_grid: other.snap_to_grid.or(self.snap_to_grid),
            spacing: update_options(self.spacing, other.spacing),
            indent: update_options(self.indent, other.indent),
            contextual_spacing: other.contextual_spacing.or(self.contextual_spacing),
            mirror_indents: other.mirror_indents.or(self.mirror_indents),
            suppress_overlapping: other.suppress_overlapping.or(self.suppress_overlapping),
            alignment: other.alignment.or(self.alignment),
            text_direction: other.text_direction.or(self.text_direction),
            text_alignment: other.text_alignment.or(self.text_alignment),
            textbox_tight_wrap: other.textbox_tight_wrap.or(self.textbox_tight_wrap),
            outline_level: other.outline_level.or(self.outline_level),
            div_id: other.div_id.or(self.div_id),
            conditional_formatting: update_options(self.conditional_formatting, other.conditional_formatting),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPrGeneral {
    pub base: PPrBase,
    pub change: Option<PPrChange>,
}

impl PPrGeneral {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing PPrGeneral");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "pPrChange" => instance.change = Some(PPrChange::from_xml_element(child_node)?),
                    _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPrTrackChanges {
    pub inserted: Option<TrackChange>,
    pub deleted: Option<TrackChange>,
    pub move_from: Option<TrackChange>,
    pub move_to: Option<TrackChange>,
}

impl ParaRPrTrackChanges {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "ins" => {
                instance.get_or_insert_with(Default::default).inserted = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "del" => {
                instance.get_or_insert_with(Default::default).deleted = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "moveFrom" => {
                instance.get_or_insert_with(Default::default).move_from =
                    Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "moveTo" => {
                instance.get_or_insert_with(Default::default).move_to = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPrOriginal {
    pub track_changes: Option<ParaRPrTrackChanges>,
    pub bases: Vec<RPrBase>,
}

impl ParaRPrOriginal {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ParaRPrOriginal");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            if ParaRPrTrackChanges::try_parse_group_node(&mut instance.track_changes, child_node)? {
                continue;
            }

            if RPrBase::is_choice_member(child_node.local_name()) {
                instance.bases.push(RPrBase::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ParaRPrChange {
    base: TrackChange,
    run_properties: ParaRPrOriginal,
}

impl ParaRPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ParaRPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let run_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rPr").into())
            .and_then(ParaRPrOriginal::from_xml_element)?;

        Ok(Self { base, run_properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPr {
    pub track_changes: Option<ParaRPrTrackChanges>,
    pub bases: Vec<RPrBase>,
    pub change: Option<ParaRPrChange>,
}

impl ParaRPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ParaRPr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            if ParaRPrTrackChanges::try_parse_group_node(&mut instance.track_changes, child_node)? {
                continue;
            }

            let local_name = child_node.local_name();
            if RPrBase::is_choice_member(local_name) {
                instance.bases.push(RPrBase::from_xml_element(child_node)?);
            } else if local_name == "rPrChange" {
                instance.change = Some(ParaRPrChange::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum HdrFtr {
    #[strum(serialize = "even")]
    Even,
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "first")]
    First,
}

#[derive(Debug, Clone, PartialEq)]
pub struct HdrFtrRef {
    pub base: Rel,
    pub header_footer_type: HdrFtr,
}

impl HdrFtrRef {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing HdrFtrRef");

        let base = Rel::from_xml_element(xml_node)?;
        let header_footer_type = xml_node
            .attributes
            .get("w:type")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?
            .parse()?;

        Ok(Self {
            base,
            header_footer_type,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum HdrFtrReferences {
    Header(HdrFtrRef),
    Footer(HdrFtrRef),
}

impl XsdType for HdrFtrReferences {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "headerReference" => Ok(HdrFtrReferences::Header(HdrFtrRef::from_xml_element(xml_node)?)),
            "footerReference" => Ok(HdrFtrReferences::Footer(HdrFtrRef::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "HdrFtrReferences",
            ))),
        }
    }
}

impl XsdChoice for HdrFtrReferences {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "headerReference" | "footerReference" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum FtnPos {
    #[strum(serialize = "pageBottom")]
    PageBottom,
    #[strum(serialize = "beneathText")]
    BeneathText,
    #[strum(serialize = "sectEnd")]
    SectionEnd,
    #[strum(serialize = "docEnd")]
    DocumentEnd,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum NumberFormat {
    #[strum(serialize = "decimal")]
    Decimal,
    #[strum(serialize = "upperRoman")]
    UpperRoman,
    #[strum(serialize = "lowerRoman")]
    LowerRoman,
    #[strum(serialize = "upperLetter")]
    UpperLetter,
    #[strum(serialize = "lowerLetter")]
    LowerLetter,
    #[strum(serialize = "ordinal")]
    Ordinal,
    #[strum(serialize = "cardinalText")]
    CardinalText,
    #[strum(serialize = "ordinalText")]
    OrdinalText,
    #[strum(serialize = "hex")]
    Hex,
    #[strum(serialize = "chicago")]
    Chicago,
    #[strum(serialize = "ideographDigital")]
    IdeographDigital,
    #[strum(serialize = "japaneseCounting")]
    JapaneseCounting,
    #[strum(serialize = "aiueo")]
    Aiueo,
    #[strum(serialize = "iroha")]
    Iroha,
    #[strum(serialize = "decimalFullWidth")]
    DecimalFullWidth,
    #[strum(serialize = "decimalHalfWidth")]
    DecimalHalfWidth,
    #[strum(serialize = "japaneseLegal")]
    JapaneseLegal,
    #[strum(serialize = "japaneseDigitalTenThousand")]
    JapaneseDigitalTenThousand,
    #[strum(serialize = "decimalEnclosedCircle")]
    DecimalEnclosedCircle,
    #[strum(serialize = "decimalFullWidth2")]
    DecimalFullWidth2,
    #[strum(serialize = "aiueoFullWidth")]
    AiueoFullWidth,
    #[strum(serialize = "irohaFullWidth")]
    IrohaFullWidth,
    #[strum(serialize = "decimalZero")]
    DecimalZero,
    #[strum(serialize = "bullet")]
    Bullet,
    #[strum(serialize = "ganada")]
    Ganada,
    #[strum(serialize = "chosung")]
    Chosung,
    #[strum(serialize = "decimalEnclosedFullstop")]
    DecimalEnclosedFullstop,
    #[strum(serialize = "decimalEnclosedParen")]
    DecimalEnclosedParen,
    #[strum(serialize = "decimalEnclosedCircleChinese")]
    DecimalEnclosedCircleChinese,
    #[strum(serialize = "ideographEnclosedCircle")]
    IdeographEnclosedCircle,
    #[strum(serialize = "ideographTraditional")]
    IdeographTraditional,
    #[strum(serialize = "ideographZodiac")]
    IdeographZodiac,
    #[strum(serialize = "ideographZodiacTraditional")]
    IdeographZodiacTraditional,
    #[strum(serialize = "taiwaneseCounting")]
    TaiwaneseCounting,
    #[strum(serialize = "ideographLegalTraditional")]
    IdeographLegalTraditional,
    #[strum(serialize = "taiwaneseCountingThousand")]
    TaiwaneseCountingThousand,
    #[strum(serialize = "taiwaneseDigital")]
    TaiwaneseDigital,
    #[strum(serialize = "chineseCounting")]
    ChineseCounting,
    #[strum(serialize = "chineseLegalSimplified")]
    ChineseLegalSimplified,
    #[strum(serialize = "chineseCountingThousand")]
    ChineseCountingThousand,
    #[strum(serialize = "koreanDigital")]
    KoreanDigital,
    #[strum(serialize = "koreanCounting")]
    KoreanCounting,
    #[strum(serialize = "koreanLegal")]
    KoreanLegal,
    #[strum(serialize = "koreanDigital2")]
    KoreanDigital2,
    #[strum(serialize = "vietnameseCounting")]
    VietnameseCounting,
    #[strum(serialize = "russianLower")]
    RussianLower,
    #[strum(serialize = "russianUpper")]
    RussianUpper,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "numberInDash")]
    NumberInDash,
    #[strum(serialize = "hebrew1")]
    Hebrew1,
    #[strum(serialize = "hebrew2")]
    Hebrew2,
    #[strum(serialize = "arabicAlpha")]
    ArabicAlpha,
    #[strum(serialize = "arabicAbjad")]
    ArabicAbjad,
    #[strum(serialize = "hindiVowels")]
    HindiVowels,
    #[strum(serialize = "hindiConsonants")]
    HindiConsonants,
    #[strum(serialize = "hindiNumbers")]
    HindiNumbers,
    #[strum(serialize = "hindiCounting")]
    HindiCounting,
    #[strum(serialize = "thaiLetters")]
    ThaiLetters,
    #[strum(serialize = "thaiNumbers")]
    ThaiNumbers,
    #[strum(serialize = "thaiCounting")]
    ThaiCounting,
    #[strum(serialize = "bahtText")]
    BahtText,
    #[strum(serialize = "dollarText")]
    DollarText,
    #[strum(serialize = "custom")]
    Custom,
}

#[derive(Debug, Clone, PartialEq)]
pub struct NumFmt {
    pub value: NumberFormat,
    pub format: Option<String>,
}

impl NumFmt {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing NumFmt");

        let mut value = None;
        let mut format = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:format" => format = Some(attr_value.clone()),
                _ => (),
            }
        }

        Ok(Self {
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
            format,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum RestartNumber {
    #[strum(serialize = "continuous")]
    Continuous,
    #[strum(serialize = "eachSect")]
    EachSection,
    #[strum(serialize = "eachPage")]
    EachPage,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct FtnEdnNumProps {
    pub numbering_start: Option<DecimalNumber>,
    pub numbering_restart: Option<RestartNumber>,
}

impl FtnEdnNumProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "numStart" => {
                instance.get_or_insert_with(Default::default).numbering_start =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "numRestart" => {
                instance.get_or_insert_with(Default::default).numbering_restart =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FtnProps {
    pub position: Option<FtnPos>,
    pub numbering_format: Option<NumFmt>,
    pub numbering_properties: Option<FtnEdnNumProps>,
}

impl FtnProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing FtnProps");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "pos" => self.position = Some(xml_node.get_val_attribute()?.parse()?),
            "numFmt" => self.numbering_format = Some(NumFmt::from_xml_element(xml_node)?),
            _ => {
                FtnEdnNumProps::try_parse_group_node(&mut self.numbering_properties, xml_node)?;
            }
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum EdnPos {
    #[strum(serialize = "sectEnd")]
    SectionEnd,
    #[strum(serialize = "docEnd")]
    DocumentEnd,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct EdnProps {
    pub position: Option<EdnPos>,
    pub numbering_format: Option<NumFmt>,
    pub numbering_properties: Option<FtnEdnNumProps>,
}

impl EdnProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing EdnProps");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "pos" => self.position = Some(xml_node.get_val_attribute()?.parse()?),
            "numFmt" => self.numbering_format = Some(NumFmt::from_xml_element(xml_node)?),
            _ => {
                FtnEdnNumProps::try_parse_group_node(&mut self.numbering_properties, xml_node)?;
            }
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum SectionMark {
    #[strum(serialize = "nextPage")]
    NextPage,
    #[strum(serialize = "nextColumn")]
    NextColumn,
    #[strum(serialize = "continuous")]
    Continuous,
    #[strum(serialize = "evenPage")]
    EvenPage,
    #[strum(serialize = "oddPage")]
    OddPage,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PageOrientation {
    #[strum(serialize = "portrait")]
    Portrait,
    #[strum(serialize = "landscape")]
    Landscape,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct PageSz {
    pub width: Option<TwipsMeasure>,
    pub height: Option<TwipsMeasure>,
    pub orientation: Option<PageOrientation>,
    pub code: Option<DecimalNumber>,
}

impl PageSz {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PageSz");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:w" => instance.width = Some(value.parse()?),
                "w:h" => instance.height = Some(value.parse()?),
                "w:orient" => instance.orientation = Some(value.parse()?),
                "w:code" => instance.code = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMar {
    pub top: SignedTwipsMeasure,
    pub right: TwipsMeasure,
    pub bottom: SignedTwipsMeasure,
    pub left: TwipsMeasure,
    pub header: TwipsMeasure,
    pub footer: TwipsMeasure,
    pub gutter: TwipsMeasure,
}

impl PageMar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PageMar");

        let mut top = None;
        let mut right = None;
        let mut bottom = None;
        let mut left = None;
        let mut header = None;
        let mut footer = None;
        let mut gutter = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:top" => top = Some(value.parse()?),
                "w:right" => right = Some(value.parse()?),
                "w:bottom" => bottom = Some(value.parse()?),
                "w:left" => left = Some(value.parse()?),
                "w:header" => header = Some(value.parse()?),
                "w:footer" => footer = Some(value.parse()?),
                "w:gutter" => gutter = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            top: top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "top"))?,
            right: right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "right"))?,
            bottom: bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bottom"))?,
            left: left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "left"))?,
            header: header.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "header"))?,
            footer: footer.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "footer"))?,
            gutter: gutter.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "gutter"))?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct PaperSource {
    pub first: Option<DecimalNumber>,
    pub other: Option<DecimalNumber>,
}

impl PaperSource {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PaperSource");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:first" => instance.first = Some(value.parse()?),
                "w:other" => instance.other = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PageBorder {
    pub base: Border,
    pub rel_id: Option<RelationshipId>,
}

impl PageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PageBorder");

        let base = Border::from_xml_element(xml_node)?;
        let rel_id = xml_node.attributes.get("r:id").map(|value| value.parse()).transpose()?;

        Ok(Self { base, rel_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TopPageBorder {
    pub base: PageBorder,
    pub top_left: Option<RelationshipId>,
    pub top_right: Option<RelationshipId>,
}

impl TopPageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TopPageBorder");

        let base = PageBorder::from_xml_element(xml_node)?;
        let top_left = xml_node.attributes.get("r:topLeft").cloned();
        let top_right = xml_node.attributes.get("r:topRight").cloned();

        Ok(Self {
            base,
            top_left,
            top_right,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct BottomPageBorder {
    pub base: PageBorder,
    pub bottom_left: Option<RelationshipId>,
    pub bottom_right: Option<RelationshipId>,
}

impl BottomPageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing BottomPageBorder");

        let base = PageBorder::from_xml_element(xml_node)?;
        let bottom_left = xml_node.attributes.get("r:bottomLeft").cloned();
        let bottom_right = xml_node.attributes.get("r:bottomRight").cloned();

        Ok(Self {
            base,
            bottom_left,
            bottom_right,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PageBorderZOrder {
    #[strum(serialize = "front")]
    Front,
    #[strum(serialize = "back")]
    Back,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PageBorderDisplay {
    #[strum(serialize = "allPages")]
    AllPages,
    #[strum(serialize = "firstPage")]
    FirstPage,
    #[strum(serialize = "notFirstPage")]
    NotFirstPage,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PageBorderOffset {
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "text")]
    Text,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageBorders {
    pub top: Option<TopPageBorder>,
    pub left: Option<PageBorder>,
    pub bottom: Option<BottomPageBorder>,
    pub right: Option<PageBorder>,
    pub z_order: Option<PageBorderZOrder>,
    pub display: Option<PageBorderDisplay>,
    pub offset_from: Option<PageBorderOffset>,
}

impl PageBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PageBorders");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:zOrder" => instance.z_order = Some(value.parse()?),
                "w:display" => instance.display = Some(value.parse()?),
                "w:offsetFrom" => instance.offset_from = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(TopPageBorder::from_xml_element(child_node)?),
                "left" => instance.left = Some(PageBorder::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(BottomPageBorder::from_xml_element(child_node)?),
                "right" => instance.right = Some(PageBorder::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum LineNumberRestart {
    #[strum(serialize = "newPage")]
    NewPage,
    #[strum(serialize = "newSection")]
    NewSection,
    #[strum(serialize = "continuous")]
    Continuous,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct LineNumber {
    pub count_by: Option<DecimalNumber>,
    pub start: Option<DecimalNumber>,
    pub distance: Option<TwipsMeasure>,
    pub restart: Option<LineNumberRestart>,
}

impl LineNumber {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing LineNumber");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:countBy" => instance.count_by = Some(value.parse()?),
                "w:start" => instance.start = Some(value.parse()?),
                "w:distance" => instance.distance = Some(value.parse()?),
                "w:restart" => instance.restart = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ChapterSep {
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "period")]
    Period,
    #[strum(serialize = "colon")]
    Color,
    #[strum(serialize = "emDash")]
    EmDash,
    #[strum(serialize = "enDash")]
    EnDash,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct PageNumber {
    pub format: Option<NumberFormat>,
    pub start: Option<DecimalNumber>,
    pub chapter_style: Option<DecimalNumber>,
    pub chapter_separator: Option<ChapterSep>,
}

impl PageNumber {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PageNumber");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:fmt" => instance.format = Some(value.parse()?),
                "w:start" => instance.start = Some(value.parse()?),
                "w:chapStyle" => instance.chapter_style = Some(value.parse()?),
                "w:chapSep" => instance.chapter_separator = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Column {
    pub width: Option<TwipsMeasure>,
    pub spacing: Option<TwipsMeasure>,
}

impl Column {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Column");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:w" => instance.width = Some(value.parse()?),
                "w:space" => instance.spacing = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Columns {
    pub columns: Vec<Column>,
    pub equal_width: Option<OnOff>,
    pub spacing: Option<TwipsMeasure>,
    pub number: Option<DecimalNumber>,
    pub separator: Option<OnOff>,
}

impl Columns {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Columns");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:equalWidth" => instance.equal_width = Some(parse_xml_bool(value)?),
                "w:space" => instance.spacing = Some(value.parse()?),
                "w:num" => instance.number = Some(value.parse()?),
                "w:sep" => instance.separator = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        instance.columns = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "col")
            .map(Column::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        match instance.columns.len() {
            0..=45 => Ok(instance),
            occurs => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "col",
                0,
                MaxOccurs::Value(45),
                occurs as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum VerticalJc {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "both")]
    Both,
    #[strum(serialize = "bottom")]
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum DocGridType {
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "lines")]
    Lines,
    #[strum(serialize = "linesAndChars")]
    LinesAndChars,
    #[strum(serialize = "snapToChars")]
    SnapToChars,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct DocGrid {
    pub doc_grid_type: Option<DocGridType>,
    pub line_pitch: Option<DecimalNumber>,
    pub char_spacing: Option<DecimalNumber>, // defaults to 0
}

impl DocGrid {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing DocGrid");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => instance.doc_grid_type = Some(value.parse()?),
                "w:linePitch" => instance.line_pitch = Some(value.parse()?),
                "w:charSpace" => instance.char_spacing = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPrContents {
    pub footnote_properties: Option<FtnProps>,
    pub endnote_properties: Option<EdnProps>,
    pub section_type: Option<SectionMark>,
    pub page_size: Option<PageSz>,
    pub page_margin: Option<PageMar>,
    pub paper_source: Option<PaperSource>,
    pub page_borders: Option<PageBorders>,
    pub line_number_type: Option<LineNumber>,
    pub page_number_type: Option<PageNumber>,
    pub columns: Option<Columns>,
    pub protect_form_fields: Option<OnOff>,
    pub vertical_align: Option<VerticalJc>,
    pub no_endnote: Option<OnOff>,
    pub title_page: Option<OnOff>,
    pub text_direction: Option<TextDirection>,
    pub bidirectional: Option<OnOff>,
    pub rtl_gutter: Option<OnOff>,
    pub document_grid: Option<DocGrid>,
    pub printer_settings: Option<Rel>,
}

impl SectPrContents {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "footnotePr" => {
                instance.get_or_insert_with(Default::default).footnote_properties =
                    Some(FtnProps::from_xml_element(xml_node)?);
                Ok(true)
            }
            "endnotePr" => {
                instance.get_or_insert_with(Default::default).endnote_properties =
                    Some(EdnProps::from_xml_element(xml_node)?);
                Ok(true)
            }
            "type" => {
                instance.get_or_insert_with(Default::default).section_type =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "pgSz" => {
                instance.get_or_insert_with(Default::default).page_size = Some(PageSz::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgMar" => {
                instance.get_or_insert_with(Default::default).page_margin = Some(PageMar::from_xml_element(xml_node)?);
                Ok(true)
            }
            "paperSrc" => {
                instance.get_or_insert_with(Default::default).paper_source =
                    Some(PaperSource::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgBorders" => {
                instance.get_or_insert_with(Default::default).page_borders =
                    Some(PageBorders::from_xml_element(xml_node)?);
                Ok(true)
            }
            "lnNumType" => {
                instance.get_or_insert_with(Default::default).line_number_type =
                    Some(LineNumber::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgNumType" => {
                instance.get_or_insert_with(Default::default).page_number_type =
                    Some(PageNumber::from_xml_element(xml_node)?);
                Ok(true)
            }
            "cols" => {
                instance.get_or_insert_with(Default::default).columns = Some(Columns::from_xml_element(xml_node)?);
                Ok(true)
            }
            "formProt" => {
                instance.get_or_insert_with(Default::default).protect_form_fields =
                    Some(parse_on_off_xml_element(xml_node)?);
                Ok(true)
            }
            "vAlign" => {
                instance.get_or_insert_with(Default::default).vertical_align =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "noEndnote" => {
                instance.get_or_insert_with(Default::default).no_endnote = Some(parse_on_off_xml_element(xml_node)?);
                Ok(true)
            }
            "titlePg" => {
                instance.get_or_insert_with(Default::default).title_page = Some(parse_on_off_xml_element(xml_node)?);
                Ok(true)
            }
            "textDirection" => {
                instance.get_or_insert_with(Default::default).text_direction =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "bidi" => {
                instance.get_or_insert_with(Default::default).bidirectional = Some(parse_on_off_xml_element(xml_node)?);
                Ok(true)
            }
            "rtlGutter" => {
                instance.get_or_insert_with(Default::default).rtl_gutter = Some(parse_on_off_xml_element(xml_node)?);
                Ok(true)
            }
            "docGrid" => {
                instance.get_or_insert_with(Default::default).document_grid =
                    Some(DocGrid::from_xml_element(xml_node)?);
                Ok(true)
            }
            "printerSettings" => {
                instance.get_or_insert_with(Default::default).printer_settings = Some(Rel::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct SectPrAttributes {
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub section_revision_id: Option<LongHexNumber>,
}

impl SectPrAttributes {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SectPrAttributes");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidSect" => instance.section_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPrBase {
    pub contents: Option<SectPrContents>,
    pub attributes: SectPrAttributes,
}

impl SectPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SectPrBase");

        Ok(Self {
            contents: SectPrContents::from_xml_element(xml_node)?,
            attributes: SectPrAttributes::from_xml_element(xml_node)?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SectPrChange {
    pub base: TrackChange,
    pub section_properties: Option<SectPrBase>,
}

impl SectPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SectPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let section_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "sectPr")
            .map(SectPrBase::from_xml_element)
            .transpose()?;

        Ok(Self {
            base,
            section_properties,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPr {
    pub header_footer_references: Vec<HdrFtrReferences>,
    pub contents: Option<SectPrContents>,
    pub change: Option<SectPrChange>,
    pub attributes: SectPrAttributes,
}

impl SectPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SectPr");

        let mut instance: Self = Default::default();

        instance.attributes = SectPrAttributes::from_xml_element(xml_node)?;

        for child_node in &xml_node.child_nodes {
            if let Some(result) = HdrFtrReferences::try_from_xml_element(child_node) {
                instance.header_footer_references.push(result?);
                continue;
            }

            if SectPrContents::try_parse_group_node(&mut instance.contents, child_node)? {
                continue;
            }

            if child_node.local_name() == "sectPrChange" {
                instance.change = Some(SectPrChange::from_xml_element(child_node)?);
            }
        }

        match instance.header_footer_references.len() {
            0..=6 => Ok(instance),
            occurs => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "headerReference|footerReference",
                0,
                MaxOccurs::Value(6),
                occurs as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PPrChange {
    pub base: TrackChange,
    pub properties: PPrBase,
}

impl PPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "pPr").into())
            .and_then(PPrBase::from_xml_element)?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPr {
    pub base: PPrBase,
    pub run_properties: Option<ParaRPr>,
    pub section_properties: Option<SectPr>,
    pub properties_change: Option<PPrChange>,
}

impl PPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing PPr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(ParaRPr::from_xml_element(child_node)?),
                "sectPr" => instance.section_properties = Some(SectPr::from_xml_element(child_node)?),
                "pPrChange" => instance.properties_change = Some(PPrChange::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct P {
    pub properties: Option<PPr>,
    pub contents: Vec<PContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub paragraph_revision_id: Option<LongHexNumber>,
    pub run_default_revision_id: Option<LongHexNumber>,
}

impl P {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing P");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidP" => instance.paragraph_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidRDefault" => instance.run_default_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "pPr" => instance.properties = Some(PPr::from_xml_element(child_node)?),
                node_name if PContent::is_choice_member(node_name) => {
                    instance.contents.push(PContent::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MeasurementOrPercent {
    DecimalOrPercent(DecimalNumberOrPercent),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for MeasurementOrPercent {
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        if let Ok(value) = s.parse::<DecimalNumberOrPercent>() {
            Ok(MeasurementOrPercent::DecimalOrPercent(value))
        } else {
            Ok(MeasurementOrPercent::UniversalMeasure(s.parse()?))
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ContentBlockContent {
    CustomXml(CustomXmlBlock),
    Sdt(Box<SdtBlock>),
    Paragraph(Box<P>),
    Table(Box<Tbl>),
    RunLevelElement(RunLevelElts),
}

impl XsdType for ContentBlockContent {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing ContentBlockContent");

        match xml_node.local_name() {
            "customXml" => Ok(ContentBlockContent::CustomXml(CustomXmlBlock::from_xml_element(
                xml_node,
            )?)),
            "sdt" => Ok(ContentBlockContent::Sdt(Box::new(SdtBlock::from_xml_element(
                xml_node,
            )?))),
            "p" => Ok(ContentBlockContent::Paragraph(Box::new(P::from_xml_element(xml_node)?))),
            "tbl" => Ok(ContentBlockContent::Table(Box::new(Tbl::from_xml_element(xml_node)?))),
            node_name if RunLevelElts::is_choice_member(&node_name) => Ok(ContentBlockContent::RunLevelElement(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentBlockContent",
            ))),
        }
    }
}

impl XsdChoice for ContentBlockContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "customXml" | "sdt" | "p" | "tbl" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct AltChunkPr {
    pub match_source: Option<OnOff>,
}

impl AltChunkPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing AltChunkPr");

        let match_source = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "matchSrc")
            .map(parse_on_off_xml_element)
            .transpose()?;

        Ok(Self { match_source })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct AltChunk {
    pub properties: Option<AltChunkPr>,
    pub rel_id: Option<RelationshipId>,
}

impl AltChunk {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing AltChunk");

        let rel_id = xml_node.attributes.get("r:id").cloned();

        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "altChunkPr")
            .map(AltChunkPr::from_xml_element)
            .transpose()?;

        Ok(Self { properties, rel_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum BlockLevelElts {
    Chunk(ContentBlockContent),
    AltChunk(AltChunk),
}

impl XsdType for BlockLevelElts {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing BlockLevelElts");

        match xml_node.local_name() {
            "altChunk" => Ok(BlockLevelElts::AltChunk(AltChunk::from_xml_element(xml_node)?)),
            node_name if ContentBlockContent::is_choice_member(node_name) => {
                Ok(BlockLevelElts::Chunk(ContentBlockContent::from_xml_element(xml_node)?))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "BlockLevelElts",
            ))),
        }
    }
}

impl XsdChoice for BlockLevelElts {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        node_name.as_ref() == "altChunk" || ContentBlockContent::is_choice_member(node_name)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Background {
    pub drawing: Option<Drawing>,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
}

impl Background {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Background");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:color" => instance.color = Some(value.parse()?),
                "w:themeColor" => instance.theme_color = Some(value.parse()?),
                "w:themeTint" => instance.theme_tint = Some(UcharHexNumber::from_str_radix(value, 16)?),
                "w:themeShade" => instance.theme_shade = Some(UcharHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        instance.drawing = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "drawing")
            .map(Drawing::from_xml_element)
            .transpose()?;

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocumentBase {
    pub background: Option<Background>,
}

impl DocumentBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing DocumentBase");

        let background = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "background")
            .map(Background::from_xml_element)
            .transpose()?;

        Ok(Self { background })
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        if xml_node.local_name() == "background" {
            self.background = Some(Background::from_xml_element(xml_node)?);
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Body {
    pub block_level_elements: Vec<BlockLevelElts>,
    pub section_properties: Option<SectPr>,
}

impl Body {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Body");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "sectPr" => instance.section_properties = Some(SectPr::from_xml_element(child_node)?),
                    node_name if BlockLevelElts::is_choice_member(node_name) => instance
                        .block_level_elements
                        .push(BlockLevelElts::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Document {
    pub base: DocumentBase,
    pub body: Option<Body>,
    pub conformance: Option<ConformanceClass>,
}

impl Document {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Document");

        let mut instance: Self = Default::default();

        instance.conformance = xml_node
            .attributes
            .get("w:conformance")
            .map(|value| value.parse())
            .transpose()?;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "body" => instance.body = Some(Body::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared::sharedtypes::UniversalMeasureUnit;
    use std::str::FromStr;

    #[test]
    pub fn test_parse_text_scale_percent() {
        assert_eq!(parse_text_scale_percent("100%").unwrap(), 100.0);
        assert_eq!(parse_text_scale_percent("600%").unwrap(), 600.0);
        assert_eq!(parse_text_scale_percent("333%").unwrap(), 333.0);
        assert_eq!(parse_text_scale_percent("0%").unwrap(), 0.0);
    }

    impl SignedTwipsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="123.456mm"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            SignedTwipsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

    #[test]
    pub fn test_signed_twips_measure_from_str() {
        assert_eq!(
            SignedTwipsMeasure::from_str("-123").unwrap(),
            SignedTwipsMeasure::Decimal(-123),
        );

        assert_eq!(
            SignedTwipsMeasure::from_str("123").unwrap(),
            SignedTwipsMeasure::Decimal(123),
        );

        assert_eq!(
            SignedTwipsMeasure::from_str("123mm").unwrap(),
            SignedTwipsMeasure::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
        );
    }

    #[test]
    pub fn test_signed_twips_measure_from_xml() {
        let xml = SignedTwipsMeasure::test_xml("signedTwipsMeasure");
        let signed_twips_measure =
            SignedTwipsMeasure::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(signed_twips_measure, SignedTwipsMeasure::test_instance());
    }

    impl HpsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="123.456mm"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

    #[test]
    pub fn test_hps_measure_from_str() {
        assert_eq!("123".parse::<HpsMeasure>().unwrap(), HpsMeasure::Decimal(123));
        assert_eq!(
            "123.456mm".parse::<HpsMeasure>().unwrap(),
            HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter)),
        );
    }

    #[test]
    pub fn test_hps_measure_from_xml() {
        let xml = HpsMeasure::test_xml("hpsMeasure");
        let hps_measure = HpsMeasure::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(hps_measure, HpsMeasure::test_instance());
    }

    impl SignedHpsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="123.456mm"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            SignedHpsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

    #[test]
    pub fn test_signed_hps_measure_from_str() {
        assert_eq!(
            SignedHpsMeasure::from_str("-123").unwrap(),
            SignedHpsMeasure::Decimal(-123),
        );

        assert_eq!(
            SignedHpsMeasure::from_str("123").unwrap(),
            SignedHpsMeasure::Decimal(123),
        );

        assert_eq!(
            SignedHpsMeasure::from_str("123mm").unwrap(),
            SignedHpsMeasure::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
        );
    }

    #[test]
    pub fn test_signed_hps_measure_from_xml() {
        let xml = SignedHpsMeasure::test_xml("signedHpsMeasure");
        let hps_measure = SignedHpsMeasure::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(hps_measure, SignedHpsMeasure::test_instance());
    }

    impl Color {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="ffffff" w:themeColor="accent1" w:themeTint="ff" w:themeShade="ff">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: HexColor::RGB([0xff, 0xff, 0xff]),
                theme_color: Some(ThemeColor::Accent1),
                theme_tint: Some(0xff),
                theme_shade: Some(0xff),
            }
        }
    }

    #[test]
    pub fn test_color_from_xml() {
        let xml = Color::test_xml("color");
        let color = Color::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(color, Color::test_instance());
    }

    impl ProofErr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="spellStart"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                error_type: ProofErrType::SpellingStart,
            }
        }
    }

    #[test]
    pub fn test_proof_err_from_xml() {
        let xml = ProofErr::test_xml("proofErr");
        let proof_err = ProofErr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(proof_err, ProofErr::test_instance());
    }

    impl Perm {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="Some id" w:displacedByCustomXml="next"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                id: String::from("Some id"),
                displaced_by_custom_xml: Some(DisplacedByCustomXml::Next),
            }
        }
    }

    #[test]
    pub fn test_perm_from_xml() {
        let xml = Perm::test_xml("perm");
        let perm = Perm::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(perm, Perm::test_instance());
    }

    impl PermStart {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:id="Some id" w:displacedByCustomXml="next" w:edGrp="everyone" w:ed="rfrostkalmar@gmail.com" w:colFirst="0" w:colLast="1">
        </{node_name}>"#, node_name=node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                permission: Perm::test_instance(),
                editor_group: Some(EdGrp::Everyone),
                editor: Some(String::from("rfrostkalmar@gmail.com")),
                first_column: Some(0),
                last_column: Some(1),
            }
        }
    }

    #[test]
    pub fn test_perm_start_from_xml() {
        let xml = PermStart::test_xml("permStart");
        let perm_start = PermStart::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(perm_start, PermStart::test_instance());
    }

    impl Markup {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:id="0"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { id: 0 }
        }
    }

    #[test]
    pub fn test_markup_from_xml() {
        let xml = Markup::test_xml("markup");
        let markup = Markup::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(markup, Markup::test_instance());
    }

    impl MarkupRange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:displacedByCustomXml="next"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Markup::test_instance(),
                displaced_by_custom_xml: Some(DisplacedByCustomXml::Next),
            }
        }
    }

    #[test]
    pub fn test_markup_range_from_xml() {
        let xml = MarkupRange::test_xml("markupRange");
        let markup_range = MarkupRange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(markup_range, MarkupRange::test_instance());
    }

    impl BookmarkRange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:displacedByCustomXml="next" w:colFirst="0" w:colLast="1">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: MarkupRange::test_instance(),
                first_column: Some(0),
                last_column: Some(1),
            }
        }
    }

    #[test]
    pub fn test_bookmark_range_from_xml() {
        let xml = BookmarkRange::test_xml("bookmarkRange");
        let bookmark_range = BookmarkRange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(bookmark_range, BookmarkRange::test_instance());
    }

    impl Bookmark {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:displacedByCustomXml="next" w:colFirst="0" w:colLast="1" w:name="Some name">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: BookmarkRange::test_instance(),
                name: String::from("Some name"),
            }
        }
    }

    #[test]
    fn test_bookmark_from_xml() {
        let xml = Bookmark::test_xml("bookmark");
        let bookmark = Bookmark::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(bookmark, Bookmark::test_instance());
    }

    impl MoveBookmark {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:id="0" w:displacedByCustomXml="next" w:colFirst="0" w:colLast="1" w:name="Some name" w:author="John Smith" w:date="2001-10-26T21:32:52">
        </{node_name}>"#, node_name=node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                base: Bookmark::test_instance(),
                author: String::from("John Smith"),
                date: DateTime::from("2001-10-26T21:32:52"),
            }
        }
    }

    #[test]
    fn test_move_bookmark_from_xml() {
        let xml = MoveBookmark::test_xml("moveBookmark");
        let move_bookmark = MoveBookmark::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(move_bookmark, MoveBookmark::test_instance());
    }

    impl TrackChange {
        pub const TEST_ATTRIBUTES: &'static str = r#"w:id="0" w:author="John Smith" w:date="2001-10-26T21:32:52""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}></{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Markup::test_instance(),
                author: String::from("John Smith"),
                date: Some(DateTime::from("2001-10-26T21:32:52")),
            }
        }
    }

    #[test]
    fn test_track_change_from_xml() {
        let xml = TrackChange::test_xml("trackChange");
        let track_change = TrackChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(track_change, TrackChange::test_instance());
    }

    impl Attr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="http://some/uri" w:name="Some name" w:val="Some value"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                uri: String::from("http://some/uri"),
                name: String::from("Some name"),
                value: String::from("Some value"),
            }
        }
    }

    #[test]
    pub fn test_attr_from_xml() {
        let xml = Attr::test_xml("attr");
        let attr = Attr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(attr, Attr::test_instance());
    }

    impl CustomXmlPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            <placeholder w:val="Placeholder" />
            {}
        </{node_name}>"#,
                Attr::test_xml("attr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                placeholder: Some(String::from("Placeholder")),
                attributes: vec![Attr::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_custom_xml_pr_from_xml() {
        let xml = CustomXmlPr::test_xml("customXmlPr");
        let custom_xml_pr = CustomXmlPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(custom_xml_pr, CustomXmlPr::test_instance());
    }

    impl SimpleField {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:instr="AUTHOR" w:fldLock="false" w:dirty="false"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_xml_recursive(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:instr="AUTHOR" w:fldLock="false" w:dirty="false">
            {}
        </{node_name}>"#,
                Self::test_xml("fldSimple"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                paragraph_contents: Vec::new(),
                field_codes: String::from("AUTHOR"),
                field_lock: Some(false),
                dirty: Some(false),
            }
        }

        pub fn test_instance_recursive() -> Self {
            Self {
                paragraph_contents: vec![PContent::SimpleField(Self::test_instance())],
                ..Self::test_instance()
            }
        }
    }

    #[test]
    pub fn test_simple_field_from_xml() {
        let xml = SimpleField::test_xml_recursive("simpleField");
        let simple_field = SimpleField::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(simple_field, SimpleField::test_instance_recursive());
    }

    impl Hyperlink {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:tgtFrame="_blank" w:tooltip="Some tooltip" w:docLocation="table" w:history="true" w:anchor="chapter1" r:id="rId1"></{node_name}>"#, node_name=node_name)
        }

        pub fn test_xml_recursive(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:tgtFrame="_blank" w:tooltip="Some tooltip" w:docLocation="table" w:history="true" w:anchor="chapter1" r:id="rId1">
            {}
        </{node_name}>"#, SimpleField::test_xml("fldSimple"), node_name=node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                paragraph_contents: Vec::new(),
                target_frame: Some(String::from("_blank")),
                tooltip: Some(String::from("Some tooltip")),
                document_location: Some(String::from("table")),
                history: Some(true),
                anchor: Some(String::from("chapter1")),
                rel_id: Some(RelationshipId::from("rId1")),
            }
        }

        pub fn test_instance_recursive() -> Self {
            Self {
                paragraph_contents: vec![PContent::test_simple_field_instance()],
                ..Self::test_instance()
            }
        }
    }

    #[test]
    pub fn test_hyperlink_from_xml() {
        let xml = Hyperlink::test_xml_recursive("hyperlink");
        let hyperlink = Hyperlink::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(hyperlink, Hyperlink::test_instance_recursive());
    }

    impl Rel {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} r:id="rId1"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                rel_id: RelationshipId::from("rId1"),
            }
        }
    }

    #[test]
    pub fn test_rel_from_xml() {
        let xml = Rel::test_xml("rel");
        let rel = Rel::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(rel, Rel::test_instance());
    }

    impl PContent {
        pub fn test_simple_field_xml() -> String {
            SimpleField::test_xml("fldSimple")
        }

        pub fn test_simple_field_instance() -> Self {
            PContent::SimpleField(SimpleField::test_instance())
        }
    }

    #[test]
    pub fn test_pcontent_content_run_content_from_xml() {
        // TODO
    }

    #[test]
    pub fn test_pcontent_simple_field_from_xml() {
        let xml = SimpleField::test_xml("fldSimple");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::SimpleField(SimpleField::test_instance()));
    }

    #[test]
    pub fn test_pcontent_hyperlink_from_xml() {
        let xml = Hyperlink::test_xml("hyperlink");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::Hyperlink(Hyperlink::test_instance()));
    }

    #[test]
    pub fn test_pcontent_subdocument_from_xml() {
        let xml = Rel::test_xml("subDoc");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::SubDocument(Rel::test_instance()));
    }

    impl CustomXmlRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="http://some/uri" w:element="Some element">
            {}
            {}
        </{node_name}>"#,
                CustomXmlPr::test_xml("customXmlPr"),
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_xml_properties: Some(CustomXmlPr::test_instance()),
                paragraph_contents: vec![PContent::test_simple_field_instance()],
                uri: String::from("http://some/uri"),
                element: String::from("Some element"),
            }
        }
    }

    #[test]
    pub fn test_custom_xml_run_from_xml() {
        let xml = CustomXmlRun::test_xml("customXmlRun");
        let custom_xml_run = CustomXmlRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(custom_xml_run, CustomXmlRun::test_instance());
    }

    impl SmartTagPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
        </{node_name}>"#,
                Attr::test_xml("attr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                attributes: vec![Attr::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_smart_tag_pr_from_xml() {
        let xml = SmartTagPr::test_xml("smartTagPr");
        let smart_tag_pr = SmartTagPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(smart_tag_pr, SmartTagPr::test_instance());
    }

    impl SmartTagRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="http://some/uri" w:element="Some element">
            {}
            {}
        </{node_name}>"#,
                SmartTagPr::test_xml("smartTagPr"),
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                smart_tag_properties: Some(SmartTagPr::test_instance()),
                paragraph_contents: vec![PContent::test_simple_field_instance()],
                uri: String::from("http://some/uri"),
                element: String::from("Some element"),
            }
        }
    }

    #[test]
    pub fn test_smart_tag_run_from_xml() {
        let xml = SmartTagRun::test_xml("smartTagRun");
        let smart_tag_run = SmartTagRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(smart_tag_run, SmartTagRun::test_instance());
    }

    impl Fonts {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:hint="default" w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"
            w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                hint: Some(Hint::Default),
                ascii: Some(String::from("Arial")),
                high_ansi: Some(String::from("Arial")),
                east_asia: Some(String::from("Arial")),
                complex_script: Some(String::from("Arial")),
                ascii_theme: Some(Theme::MajorAscii),
                high_ansi_theme: Some(Theme::MajorHighAnsi),
                east_asia_theme: Some(Theme::MajorEastAsia),
                complex_script_theme: Some(Theme::MajorBidirectional),
            }
        }
    }

    #[test]
    pub fn test_fonts_from_xml() {
        let xml = Fonts::test_xml("fonts");
        let fonts = Fonts::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(fonts, Fonts::test_instance());
    }

    impl Underline {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="single" w:color="ffffff" w:themeColor="accent1" w:themeTint="ff" w:themeShade="ff">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(UnderlineType::Single),
                color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
                theme_color: Some(ThemeColor::Accent1),
                theme_tint: Some(0xff),
                theme_shade: Some(0xff),
            }
        }
    }

    #[test]
    pub fn test_underline_from_xml() {
        let xml = Underline::test_xml("underline");
        let underline = Underline::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(underline, Underline::test_instance());
    }

    impl Border {
        const TEST_ATTRIBUTES: &'static str =
            r#"w:val="single" w:color="ffffff" w:themeColor="accent1" w:themeTint="ff"
            w:themeShade="ff" w:sz="100" w:space="100" w:shadow="true" w:frame="true""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                </{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: BorderType::Single,
                color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
                theme_color: Some(ThemeColor::Accent1),
                theme_tint: Some(0xff),
                theme_shade: Some(0xff),
                size: Some(100),
                spacing: Some(100),
                shadow: Some(true),
                frame: Some(true),
            }
        }
    }

    #[test]
    pub fn test_border_from_xml() {
        let xml = Border::test_xml("border");
        let border = Border::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(border, Border::test_instance());
    }

    impl Shd {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="solid" w:color="ffffff" w:themeColor="accent1" w:themeTint="ff"
            w:themeShade="ff" w:fill="ffffff" w:themeFill="accent1" w:themeFillTint="ff" w:themeFillShade="ff">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: ShdType::Solid,
                color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
                theme_color: Some(ThemeColor::Accent1),
                theme_tint: Some(0xff),
                theme_shade: Some(0xff),
                fill: Some(HexColor::RGB([0xff, 0xff, 0xff])),
                theme_fill: Some(ThemeColor::Accent1),
                theme_fill_tint: Some(0xff),
                theme_fill_shade: Some(0xff),
            }
        }
    }

    #[test]
    pub fn test_shd_from_xml() {
        let xml = Shd::test_xml("shd");
        let shd = Shd::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(shd, Shd::test_instance());
    }

    impl FitText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="123.456mm" w:id="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: TwipsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(
                    123.456,
                    UniversalMeasureUnit::Millimeter,
                )),
                id: Some(1),
            }
        }
    }

    #[test]
    pub fn test_fit_text_from_xml() {
        let xml = FitText::test_xml("fitText");
        let fit_text = FitText::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(fit_text, FitText::test_instance());
    }

    impl Language {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="en" w:eastAsia="jp" w:bidi="fa"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(Lang::from("en")),
                east_asia: Some(Lang::from("jp")),
                bidirectional: Some(Lang::from("fa")),
            }
        }
    }

    #[test]
    pub fn test_language_from_xml() {
        let xml = Language::test_xml("language");
        let language = Language::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap());
        assert_eq!(language, Language::test_instance());
    }

    impl EastAsianLayout {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="1" w:combine="true" w:combineBrackets="square" w:vert="true" w:vertCompress="true">
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                id: Some(1),
                combine: Some(true),
                combine_brackets: Some(CombineBrackets::Square),
                vertical: Some(true),
                vertical_compress: Some(true),
            }
        }
    }

    #[test]
    pub fn test_east_asian_layout_from_xml() {
        let xml = EastAsianLayout::test_xml("eastAsianLayout");
        let east_asian_layout = EastAsianLayout::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(east_asian_layout, EastAsianLayout::test_instance());
    }

    impl RPrBase {
        pub fn test_run_style_xml() -> &'static str {
            r#"<rStyle w:val="Arial"></rStyle>"#
        }

        pub fn test_run_style_instance() -> Self {
            RPrBase::RunStyle(String::from("Arial"))
        }
    }

    // TODO Write some more unit tests

    #[test]
    pub fn test_r_pr_base_run_style_from_xml() {
        let xml = RPrBase::test_run_style_xml();
        let r_pr_base = RPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(r_pr_base, RPrBase::test_run_style_instance());
    }

    impl RPrOriginal {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                RPrBase::test_run_style_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                r_pr_bases: vec![RPrBase::test_run_style_instance()],
            }
        }
    }

    #[test]
    pub fn test_r_pr_original_from_xml() {
        let xml = RPrOriginal::test_xml("rPrOriginal");
        let r_pr_original = RPrOriginal::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(r_pr_original, RPrOriginal::test_instance());
    }

    impl RPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:author="John Smith" w:date="2001-10-26T21:32:52">
            {}
        </{node_name}>"#,
                RPrOriginal::test_xml("rPr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                run_properties: RPrOriginal::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_r_pr_change_from_xml() {
        let xml = RPrChange::test_xml("rRpChange");
        let r_pr_change = RPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(r_pr_change, RPrChange::test_instance());
    }

    impl RPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
            {}
        </{node_name}>"#,
                RPrBase::test_run_style_xml(),
                RPrChange::test_xml("rPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                r_pr_bases: vec![RPrBase::test_run_style_instance()],
                run_properties_change: Some(RPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_r_pr_from_xml() {
        let xml = RPr::test_xml("rPr");
        let r_pr_content = RPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(r_pr_content, RPr::test_instance());
    }

    impl SdtListItem {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:displayText="Displayed" w:value="Some value"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                display_text: String::from("Displayed"),
                value: String::from("Some value"),
            }
        }
    }

    #[test]
    pub fn test_sdt_list_item_from_xml() {
        let xml = SdtListItem::test_xml("sdtListItem");
        let sdt_list_item = SdtListItem::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_list_item, SdtListItem::test_instance());
    }

    impl SdtComboBox {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:lastValue="Some value">
            {}
            {}
        </{node_name}>"#,
                SdtListItem::test_xml("listItem"),
                SdtListItem::test_xml("listItem"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                list_items: vec![SdtListItem::test_instance(), SdtListItem::test_instance()],
                last_value: Some(String::from("Some value")),
            }
        }
    }

    #[test]
    pub fn test_sdt_combo_box_from_xml() {
        let xml = SdtComboBox::test_xml("sdtComboBox");
        let sdt_combo_box = SdtComboBox::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_combo_box, SdtComboBox::test_instance());
    }

    impl SdtDate {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:fullDate="2001-10-26T21:32:52">
            <dateFormat w:val="MM-YYYY" />
            <lid w:val="ja-JP" />
            <storeMappedDataAs w:val="dateTime" />
            <calendar w:val="gregorian" />
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                date_format: Some(String::from("MM-YYYY")),
                language_id: Some(Lang::from("ja-JP")),
                store_mapped_data_as: Some(SdtDateMappingType::DateTime),
                calendar: Some(CalendarType::Gregorian),
                full_date: Some(DateTime::from("2001-10-26T21:32:52")),
            }
        }
    }

    #[test]
    pub fn test_sdt_date_from_xml() {
        let xml = SdtDate::test_xml("sdtDate");
        let sdt_date = SdtDate::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_date, SdtDate::test_instance());
    }

    impl SdtDocPart {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            <docPartGallery w:val="Some string" />
            <docPartCategory w:val="Some string" />
            <docPartUnique w:val="true" />
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                doc_part_gallery: Some(String::from("Some string")),
                doc_part_category: Some(String::from("Some string")),
                doc_part_unique: Some(true),
            }
        }
    }

    #[test]
    pub fn test_sdt_doc_part_from_xml() {
        let xml = SdtDocPart::test_xml("sdtDocPart");
        let sdt_doc_part = SdtDocPart::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_doc_part, SdtDocPart::test_instance());
    }

    impl SdtDropDownList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:lastValue="Some value">
            {}
            {}
        </{node_name}>"#,
                SdtListItem::test_xml("listItem"),
                SdtListItem::test_xml("listItem"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                list_items: vec![SdtListItem::test_instance(), SdtListItem::test_instance()],
                last_value: Some(String::from("Some value")),
            }
        }
    }

    #[test]
    pub fn test_sdt_drop_down_list_from_xml() {
        let xml = SdtDropDownList::test_xml("sdtDropDownList");
        let sdt_combo_box = SdtDropDownList::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_combo_box, SdtDropDownList::test_instance());
    }

    impl SdtText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:multiLine="true"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self { is_multi_line: true }
        }
    }

    #[test]
    pub fn test_sdt_text_from_xml() {
        let xml = SdtText::test_xml("sdtText");
        let sdt_text = SdtText::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(sdt_text, SdtText::test_instance());
    }

    #[test]
    pub fn test_sdt_pr_control_choice_from_xml() {
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<equation></equation>").unwrap()).unwrap(),
            SdtPrChoice::Equation,
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtComboBox::test_xml("comboBox").as_str()).unwrap())
                .unwrap(),
            SdtPrChoice::ComboBox(SdtComboBox::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDate::test_xml("date").as_str()).unwrap()).unwrap(),
            SdtPrChoice::Date(SdtDate::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDocPart::test_xml("docPartObj").as_str()).unwrap())
                .unwrap(),
            SdtPrChoice::DocumentPartObject(SdtDocPart::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDocPart::test_xml("docPartList").as_str()).unwrap())
                .unwrap(),
            SdtPrChoice::DocumentPartList(SdtDocPart::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(
                &XmlNode::from_str(SdtDropDownList::test_xml("dropDownList").as_str()).unwrap()
            )
            .unwrap(),
            SdtPrChoice::DropDownList(SdtDropDownList::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<picture></picture>").unwrap()).unwrap(),
            SdtPrChoice::Picture,
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<richText></richText>").unwrap()).unwrap(),
            SdtPrChoice::RichText,
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtText::test_xml("text").as_str()).unwrap()).unwrap(),
            SdtPrChoice::Text(SdtText::test_instance()),
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<citation></citation>").unwrap()).unwrap(),
            SdtPrChoice::Citation,
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<group></group>").unwrap()).unwrap(),
            SdtPrChoice::Group,
        );
        assert_eq!(
            SdtPrChoice::from_xml_element(&XmlNode::from_str("<bibliography></bibliography>").unwrap()).unwrap(),
            SdtPrChoice::Bibliography,
        );
    }

    impl Placeholder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            <docPart w:val="title" />
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                document_part: String::from("title"),
            }
        }
    }

    #[test]
    pub fn test_placeholder_from_xml() {
        let xml = Placeholder::test_xml("placeholder");
        assert_eq!(
            Placeholder::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Placeholder::test_instance()
        );
    }

    impl DataBinding {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:prefixMappings="xmlns:ns0='http://example.com/example'" w:xpath="//ns0:book" w:storeItemID="testXmlPart">
        </{node_name}>"#
            , node_name=node_name
        )
        }

        pub fn test_instance() -> Self {
            Self {
                prefix_mappings: Some(String::from("xmlns:ns0='http://example.com/example'")),
                xpath: String::from("//ns0:book"),
                store_item_id: String::from("testXmlPart"),
            }
        }
    }

    #[test]
    pub fn test_data_binding_from_xml() {
        let xml = DataBinding::test_xml("dataBinding");
        assert_eq!(
            DataBinding::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DataBinding::test_instance()
        );
    }

    impl SdtPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
            <alias w:val="Alias" />
            <tag w:val="Tag"/>
            <id w:val="1" />
            <lock w:val="unlocked" />
            {}
            <temporary w:val="false" />
            <showingPlcHdr w:val="false" />
            {}
            <label w:val="1" />
            <tabIndex w:val="1" />
            <equation />
        </{node_name}>"#,
                RPr::test_xml("rPr"),
                Placeholder::test_xml("placeholder"),
                DataBinding::test_xml("dataBinding"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties: Some(RPr::test_instance()),
                alias: Some(String::from("Alias")),
                tag: Some(String::from("Tag")),
                id: Some(1),
                lock: Some(Lock::Unlocked),
                placeholder: Some(Placeholder::test_instance()),
                temporary: Some(false),
                showing_placeholder_header: Some(false),
                data_binding: Some(DataBinding::test_instance()),
                label: Some(1),
                tab_index: Some(1),
                control_choice: Some(SdtPrChoice::Equation),
            }
        }
    }

    #[test]
    pub fn test_sdt_pr_from_xml() {
        let xml = SdtPr::test_xml("sdtPr");
        assert_eq!(
            SdtPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtPr::test_instance()
        );
    }

    impl SdtEndPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
            {rpr}
            {rpr}
        </{node_name}>",
                rpr = RPr::test_xml("rPr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties_vec: vec![RPr::test_instance(), RPr::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_std_end_pr_from_xml() {
        let xml = SdtEndPr::test_xml("sdtEndPr");
        assert_eq!(
            SdtEndPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtEndPr::test_instance(),
        );
    }

    impl SdtContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {pcontent}
            {pcontent}
        </{node_name}>"#,
                pcontent = PContent::test_simple_field_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![
                    PContent::test_simple_field_instance(),
                    PContent::test_simple_field_instance(),
                ],
            }
        }
    }

    #[test]
    pub fn test_sdt_content_run_from_xml() {
        let xml = SdtContentRun::test_xml("sdtContentRun");
        assert_eq!(
            SdtContentRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtContentRun::test_instance()
        );
    }

    impl SdtRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
            {}
            {}
        </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentRun::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                sdt_properties: Some(SdtPr::test_instance()),
                sdt_end_properties: Some(SdtEndPr::test_instance()),
                sdt_content: Some(SdtContentRun::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sdt_run_from_xml() {
        let xml = SdtRun::test_xml("sdtRun");
        assert_eq!(
            SdtRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtRun::test_instance()
        );
    }

    impl DirContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="ltr">
                {}
            </{node_name}>"#,
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![PContent::test_simple_field_instance()],
                value: Some(Direction::LeftToRight),
            }
        }
    }

    #[test]
    pub fn test_dir_content_run_from_xml() {
        let xml = DirContentRun::test_xml("dirContentRun");
        assert_eq!(
            DirContentRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DirContentRun::test_instance()
        );
    }

    impl BdoContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="ltr">
                {}
            </{node_name}>"#,
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![PContent::test_simple_field_instance()],
                value: Some(Direction::LeftToRight),
            }
        }
    }

    #[test]
    pub fn test_bdo_content_run_from_xml() {
        let xml = DirContentRun::test_xml("bdoContentRun");
        assert_eq!(
            DirContentRun::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DirContentRun::test_instance()
        );
    }

    impl Br {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="page" w:clear="none"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                break_type: Some(BrType::Page),
                clear: Some(BrClear::None),
            }
        }
    }

    #[test]
    pub fn test_br_from_xml() {
        let xml = Br::test_xml("br");
        assert_eq!(
            Br::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Br::test_instance()
        );
    }

    impl Text {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} xml:space="default">Some text</{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                text: String::from("Some text"),
                xml_space: Some(String::from("default")),
            }
        }
    }

    #[test]
    pub fn test_text_from_xml() {
        let xml = Text::test_xml("text");
        assert_eq!(
            Text::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Text::test_instance()
        );
    }

    impl Sym {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:font="Arial" w:char="ffff"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                font: Some(String::from("Arial")),
                character: Some(0xffff),
            }
        }
    }

    #[test]
    pub fn test_sym_from_xml() {
        let xml = Sym::test_xml("sym");
        assert_eq!(
            Sym::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Sym::test_instance()
        );
    }

    impl Control {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:name="Name" w:shapeid="Id" r:id="rId1" >
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: Some(String::from("Name")),
                shapeid: Some(String::from("Id")),
                rel_id: Some(String::from("rId1")),
            }
        }
    }

    #[test]
    pub fn test_control_from_xml() {
        let xml = Control::test_xml("control");
        assert_eq!(
            Control::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()),
            Control::test_instance()
        );
    }

    impl ObjectEmbed {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:drawAspect="content" r:id="rId1" w:progId="AVIFile" w:shapeId="1" w:fieldCodes="\f 0">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                draw_aspect: Some(ObjectDrawAspect::Content),
                rel_id: String::from("rId1"),
                application_id: Some(String::from("AVIFile")),
                shape_id: Some(String::from("1")),
                field_codes: Some(String::from(r#"\f 0"#)),
            }
        }
    }

    #[test]
    pub fn test_object_embed_from_xml() {
        let xml = ObjectEmbed::test_xml("objectEmbed");
        assert_eq!(
            ObjectEmbed::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ObjectEmbed::test_instance()
        );
    }

    impl ObjectLink {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:drawAspect="content" r:id="rId1" w:progId="AVIFile" w:shapeId="1" w:fieldCodes="\f 0" w:updateMode="always" w:lockedField="true">
            </{node_name}>"#,
                node_name=node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: ObjectEmbed::test_instance(),
                update_mode: ObjectUpdateMode::Always,
                locked_field: Some(true),
            }
        }
    }

    #[test]
    pub fn test_object_link_from_xml() {
        let xml = ObjectLink::test_xml("objectLink");
        assert_eq!(
            ObjectLink::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ObjectLink::test_instance()
        );
    }

    impl Drawing {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                Anchor::test_xml("wp:anchor"),
                Inline::test_xml("wp:inline"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![
                DrawingChoice::Anchor(Anchor::test_instance()),
                DrawingChoice::Inline(Inline::test_instance()),
            ])
        }
    }

    #[test]
    pub fn test_drawing_from_xml() {
        let xml = Drawing::test_xml("drawing");
        assert_eq!(
            Drawing::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Drawing::test_instance()
        );
    }

    impl Object {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:dxaOrig="123.456mm" w:dyaOrig="123">
                {}
                {}
            </{node_name}>"#,
                Drawing::test_xml("drawing"),
                Control::test_xml("control"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                drawing: Some(Drawing::test_instance()),
                choice: Some(ObjectChoice::Control(Control::test_instance())),
                original_image_width: Some(TwipsMeasure::UniversalMeasure(UniversalMeasure::new(
                    123.456,
                    UniversalMeasureUnit::Millimeter,
                ))),
                original_image_height: Some(TwipsMeasure::Decimal(123)),
            }
        }
    }

    #[test]
    pub fn test_object_from_xml() {
        let xml = Object::test_xml("object");
        assert_eq!(
            Object::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Object::test_instance()
        );
    }

    impl FFHelpText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="text" w:val="Help text"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                info_text_type: Some(InfoTextType::Text),
                value: Some(FFHelpTextVal::from("Help text")),
            }
        }
    }

    #[test]
    pub fn test_ff_help_text_from_xml() {
        let xml = FFHelpText::test_xml("ffHelpText");
        assert_eq!(
            FFHelpText::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FFHelpText::test_instance()
        );
    }

    impl FFStatusText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="text" w:val="Status text"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                info_text_type: Some(InfoTextType::Text),
                value: Some(FFStatusTextVal::from("Status text")),
            }
        }
    }

    #[test]
    pub fn test_ff_status_text_from_xml() {
        let xml = FFStatusText::test_xml("ffStatusText");
        assert_eq!(
            FFStatusText::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FFStatusText::test_instance()
        );
    }

    #[test]
    pub fn test_ff_check_box_size_choice_from_xml() {
        let xml = r#"<size w:val="123"></size>"#;
        assert_eq!(
            FFCheckBoxSizeChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFCheckBoxSizeChoice::Explicit(HpsMeasure::Decimal(123)),
        );
        let xml = r#"<sizeAuto w:val="true"></sizeAuto>"#;
        assert_eq!(
            FFCheckBoxSizeChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFCheckBoxSizeChoice::Auto(true),
        );
    }

    impl FFCheckBox {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <sizeAuto w:val="true" />
                <default w:val="true" />
                <checked w:val="true" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                size: FFCheckBoxSizeChoice::Auto(true),
                is_default: Some(true),
                is_checked: Some(true),
            }
        }
    }

    #[test]
    pub fn test_ff_check_box_from_xml() {
        let xml = FFCheckBox::test_xml("ffCheckBox");
        assert_eq!(
            FFCheckBox::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FFCheckBox::test_instance()
        );
    }

    impl FFDDList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <result w:val="1" />
                <default w:val="1" />
                <listEntry w:val="Entry1" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                result: Some(1),
                default: Some(1),
                list_entries: vec![String::from("Entry1")],
            }
        }
    }

    #[test]
    pub fn test_ff_ddlist_from_xml() {
        let xml = FFDDList::test_xml("ffDDList");
        assert_eq!(
            FFDDList::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FFDDList::test_instance()
        );
    }

    impl FFTextInput {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <type w:val="regular" />
                <default w:val="Default" />
                <maxLength w:val="100" />
                <format w:val=".*" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                text_type: Some(FFTextType::Regular),
                default: Some(String::from("Default")),
                max_length: Some(100),
                format: Some(String::from(".*")),
            }
        }
    }

    #[test]
    pub fn test_ff_text_input_from_xml() {
        let xml = FFTextInput::test_xml("ffTextInput");
        assert_eq!(
            FFTextInput::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FFTextInput::test_instance(),
        );
    }

    impl FldChar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:fldCharType="begin" w:fldLock="false" w:dirty="false">
                <name w:val="Some name" />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                form_field_properties: Some(FFData::Name(FFName::from("Some name"))),
                field_char_type: FldCharType::Begin,
                field_lock: Some(false),
                dirty: Some(false),
            }
        }
    }

    #[test]
    pub fn test_fld_char_from_xml() {
        let xml = FldChar::test_xml("fldChar");
        assert_eq!(
            FldChar::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FldChar::test_instance()
        );
    }

    impl RubyPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <rubyAlign w:val="left" />
                <hps w:val="123" />
                <hpsRaise w:val="123" />
                <hpsBaseText w:val="123" />
                <lid w:val="en-US" />
                <dirty w:val="true" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_align: RubyAlign::Left,
                hps: HpsMeasure::Decimal(123),
                hps_raise: HpsMeasure::Decimal(123),
                hps_base_text: HpsMeasure::Decimal(123),
                language_id: Lang::from("en-US"),
                dirty: Some(true),
            }
        }
    }

    #[test]
    pub fn test_ruby_pr_from_xml() {
        let xml = RubyPr::test_xml("rubyPr");
        assert_eq!(
            RubyPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            RubyPr::test_instance()
        );
    }

    #[test]
    pub fn test_ruby_content_choice_from_xml() {
        let xml = R::test_xml("r");
        assert_eq!(
            RubyContentChoice::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            RubyContentChoice::Run(R::test_instance())
        );
        let xml = ProofErr::test_xml("proofErr");
        assert_eq!(
            RubyContentChoice::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            RubyContentChoice::RunLevelElement(RunLevelElts::ProofError(ProofErr::test_instance()))
        );
    }

    impl RubyContent {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                R::test_xml("r"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_contents: vec![RubyContentChoice::Run(R::test_instance())],
            }
        }
    }

    #[test]
    pub fn test_ruby_content_from_xml() {
        let xml = RubyContent::test_xml("rubyContent");
        assert_eq!(
            RubyContent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            RubyContent::test_instance()
        );
    }

    impl R {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:rsidRPr="ffffffff" w:rsidDel="ffffffff" w:rsidR="ffffffff">
                {}
                {}
            </{node_name}>"#,
                RPr::test_xml("rPr"),
                Br::test_xml("br"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties: Some(RPr::test_instance()),
                run_inner_contents: vec![RunInnerContent::Break(Br::test_instance())],
                run_properties_revision_id: Some(0xffffffff),
                deletion_revision_id: Some(0xffffffff),
                run_revision_id: Some(0xffffffff),
            }
        }
    }

    #[test]
    pub fn test_r_from_xml() {
        let xml = R::test_xml("r");
        assert_eq!(
            R::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            R::test_instance()
        );
    }

    impl Ruby {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                RubyPr::test_xml("rubyPr"),
                RubyContent::test_xml("rt"),
                RubyContent::test_xml("rubyBase"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_properties: RubyPr::test_instance(),
                ruby_content: RubyContent::test_instance(),
                ruby_base: RubyContent::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_ruby_from_xml() {
        let xml = Ruby::test_xml("ruby");
        assert_eq!(
            Ruby::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Ruby::test_instance()
        );
    }

    impl FtnEdnRef {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:customMarkFollows="true" w:id="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_mark_follows: Some(true),
                id: 1,
            }
        }
    }

    #[test]
    pub fn test_ftn_edn_ref_from_xml() {
        let xml = FtnEdnRef::test_xml("ftnEdnRef");
        assert_eq!(
            FtnEdnRef::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FtnEdnRef::test_instance()
        );
    }

    impl PTab {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:alignment="left" w:relativeTo="margin" w:leader="none">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                alignment: PTabAlignment::Left,
                relative_to: PTabRelativeTo::Margin,
                leader: PTabLeader::None,
            }
        }
    }

    #[test]
    pub fn test_p_tab_from_xml() {
        let xml = PTab::test_xml("pTab");
        assert_eq!(
            PTab::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PTab::test_instance()
        );
    }

    impl RunTrackChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:author="John Smith" w:date="2001-10-26T21:32:52">
                {}
            </{node_name}>"#,
                R::test_xml("r"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                choices: vec![RunTrackChangeChoice::ContentRunContent(ContentRunContent::Run(
                    R::test_instance(),
                ))],
            }
        }
    }

    #[test]
    pub fn test_run_track_change_from_xml() {
        let xml = RunTrackChange::test_xml("runTrackChange");
        assert_eq!(
            RunTrackChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            RunTrackChange::test_instance()
        );
    }

    impl CustomXmlBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="https://some/uri" w:element="Some element">
                {}
            </{node_name}>"#,
                CustomXmlPr::test_xml("customXmlPr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_xml_properties: Some(CustomXmlPr::test_instance()),
                block_contents: Vec::new(),
                uri: Some(String::from("https://some/uri")),
                element: XmlName::from("Some element"),
            }
        }
    }

    #[test]
    pub fn test_custom_xml_block_from_xml() {
        let xml = CustomXmlBlock::test_xml("customXmlBlock");
        println!("{}", xml);
        assert_eq!(
            CustomXmlBlock::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            CustomXmlBlock::test_instance()
        );
    }

    impl SdtContentBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
                {}
            </{node_name}>",
                CustomXmlBlock::test_xml("customXml"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                block_contents: vec![ContentBlockContent::CustomXml(CustomXmlBlock::test_instance())],
            }
        }
    }

    #[test]
    pub fn test_sdt_content_block_from_xml() {
        let xml = SdtContentBlock::test_xml("sdtContentBlock");
        assert_eq!(
            SdtContentBlock::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtContentBlock::test_instance()
        );
    }

    impl SdtBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentBlock::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                sdt_properties: Some(SdtPr::test_instance()),
                sdt_end_properties: Some(SdtEndPr::test_instance()),
                sdt_content: Some(SdtContentBlock::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sdt_block_from_xml() {
        let xml = SdtBlock::test_xml("sdtBlock");
        assert_eq!(
            SdtBlock::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtBlock::test_instance()
        );
    }

    impl FramePr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:dropCap="drop" w:lines="1" w:w="100" w:h="100" w:vSpace="50" w:hSpace="50" w:wrap="auto"
                w:hAnchor="text" w:vAnchor="text" w:x="0" w:xAlign="left" w:y="0" w:yAlign="top" w:hRule="auto" w:anchorLock="true">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                drop_cap: Some(DropCap::Drop),
                lines: Some(1),
                width: Some(TwipsMeasure::Decimal(100)),
                height: Some(TwipsMeasure::Decimal(100)),
                vertical_space: Some(TwipsMeasure::Decimal(50)),
                horizontal_space: Some(TwipsMeasure::Decimal(50)),
                wrap: Some(Wrap::Auto),
                horizontal_anchor: Some(HAnchor::Text),
                vertical_anchor: Some(VAnchor::Text),
                x: Some(SignedTwipsMeasure::Decimal(0)),
                x_align: Some(XAlign::Left),
                y: Some(SignedTwipsMeasure::Decimal(0)),
                y_align: Some(YAlign::Top),
                height_rule: Some(HeightRule::Auto),
                anchor_lock: Some(true),
            }
        }
    }

    #[test]
    pub fn test_frame_pr_from_xml() {
        let xml = FramePr::test_xml("framePr");
        assert_eq!(
            FramePr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FramePr::test_instance()
        );
    }

    impl NumPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <ilvl w:val="1" />
                <numId w:val="1" />
                {}
            </{node_name}>"#,
                TrackChange::test_xml("ins"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                indent_level: Some(1),
                numbering_id: Some(1),
                inserted: Some(TrackChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_num_pr_from_xml() {
        let xml = NumPr::test_xml("numPr");
        assert_eq!(
            NumPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            NumPr::test_instance(),
        );
    }

    impl PBdr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Border::test_xml("top"),
                Border::test_xml("left"),
                Border::test_xml("bottom"),
                Border::test_xml("right"),
                Border::test_xml("between"),
                Border::test_xml("bar"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                left: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                right: Some(Border::test_instance()),
                between: Some(Border::test_instance()),
                bar: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_bdr_from_xml() {
        let xml = PBdr::test_xml("pBdr");
        assert_eq!(
            PBdr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PBdr::test_instance(),
        );
    }

    impl TabStop {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="start" w:leader="dot" w:pos="0"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: TabJc::Start,
                leader: Some(TabTlc::Dot),
                position: SignedTwipsMeasure::Decimal(0),
            }
        }
    }

    #[test]
    pub fn test_tab_stop_from_xml() {
        let xml = TabStop::test_xml("tabStop");
        assert_eq!(
            TabStop::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TabStop::test_instance(),
        );
    }

    impl Tabs {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
                {tab_stop}
                {tab_stop}
            </{node_name}>",
                tab_stop = TabStop::test_xml("tab"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![TabStop::test_instance(), TabStop::test_instance()])
        }
    }

    #[test]
    pub fn test_tabs_from_xml() {
        let xml = Tabs::test_xml("tabs");
        assert_eq!(
            Tabs::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Tabs::test_instance(),
        );
    }

    impl Spacing {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:before="10" w:beforeLines="1" w:beforeAutospacing="true"
                w:after="10" w:afterLines="1" w:afterAutospacing="true" w:line="50" w:lineRule="auto">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                before: Some(TwipsMeasure::Decimal(10)),
                before_lines: Some(1),
                before_autospacing: Some(true),
                after: Some(TwipsMeasure::Decimal(10)),
                after_lines: Some(1),
                after_autospacing: Some(true),
                line: Some(SignedTwipsMeasure::Decimal(50)),
                line_rule: Some(LineSpacingRule::Auto),
            }
        }
    }

    #[test]
    pub fn test_spacing_from_xml() {
        let xml = Spacing::test_xml("spacing");
        assert_eq!(
            Spacing::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Spacing::test_instance(),
        );
    }

    impl Ind {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:start="50" w:startChars="0" w:end="50" w:endChars="10" w:hanging="50" w:hangingChars="5"
                w:firstLine="50" w:firstLineChars="5">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start: Some(SignedTwipsMeasure::Decimal(50)),
                start_chars: Some(0),
                end: Some(SignedTwipsMeasure::Decimal(50)),
                end_chars: Some(10),
                left: None,
                left_chars: None,
                right: None,
                right_chars: None,
                hanging: Some(TwipsMeasure::Decimal(50)),
                hanging_chars: Some(5),
                first_line: Some(TwipsMeasure::Decimal(50)),
                first_line_chars: Some(5),
            }
        }
    }

    #[test]
    pub fn test_ind_from_xml() {
        let xml = Ind::test_xml("ind");
        assert_eq!(
            Ind::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Ind::test_instance(),
        );
    }

    impl Cnf {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:firstRow="true" w:lastRow="true" w:firstColumn="true" w:lastColumn="true" w:oddVBand="true"
                w:evenVBand="true" w:oddHBand="true" w:evenHBand="true" w:firstRowFirstColumn="true" w:firstRowLastColumn="true"
                w:lastRowFirstColumn="true" w:lastRowLastColumn="true">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first_row: Some(true),
                last_row: Some(true),
                first_column: Some(true),
                last_column: Some(true),
                odd_vertical_band: Some(true),
                even_vertical_band: Some(true),
                odd_horizontal_band: Some(true),
                even_horizontal_band: Some(true),
                first_row_first_column: Some(true),
                first_row_last_column: Some(true),
                last_row_first_column: Some(true),
                last_row_last_column: Some(true),
            }
        }
    }

    #[test]
    pub fn test_cnf_from_xml() {
        let xml = Cnf::test_xml("cnf");
        assert_eq!(
            Cnf::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Cnf::test_instance(),
        );
    }

    impl PPrBase {
        pub fn test_extension_xml() -> String {
            format!(
                r#"<pStyle w:val="Normal" />
                <keepNext w:val="true" />
                <keepLines w:val="true" />
                <pageBreakBefore w:val="true" />
                {}
                <widowControl w:val="true" />
                {}
                <suppressLineNumbers w:val="true" />
                {}
                {}
                {}
                <suppressAutoHyphens w:val="true" />
                <kinsoku w:val="true" />
                <wordWrap w:val="true" />
                <overflowPunct w:val="true" />
                <topLinePunct w:val="true" />
                <autoSpaceDE w:val="true" />
                <autoSpaceDN w:val="true" />
                <bidi w:val="true" />
                <adjustRightInd w:val="true" />
                <snapToGrid w:val="true" />
                {}
                {}
                <contextualSpacing w:val="true" />
                <mirrorIndents w:val="true" />
                <suppressOverlap w:val="true" />
                <jc w:val="start" />
                <textDirection w:val="lr" />
                <textAlignment w:val="auto" />
                <textboxTightWrap w:val="none" />
                <outlineLvl w:val="1" />
                <divId w:val="1" />
                {}"#,
                FramePr::test_xml("framePr"),
                NumPr::test_xml("numPr"),
                PBdr::test_xml("pBdr"),
                Shd::test_xml("shd"),
                Tabs::test_xml("tabs"),
                Spacing::test_xml("spacing"),
                Ind::test_xml("ind"),
                Cnf::test_xml("cnfStyle"),
            )
        }

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                style: Some(String::from("Normal")),
                keep_with_next: Some(true),
                keep_lines_on_one_page: Some(true),
                start_on_next_page: Some(true),
                frame_properties: Some(FramePr::test_instance()),
                widow_control: Some(true),
                numbering_properties: Some(NumPr::test_instance()),
                suppress_line_numbers: Some(true),
                borders: Some(PBdr::test_instance()),
                shading: Some(Shd::test_instance()),
                tabs: Some(Tabs::test_instance()),
                suppress_auto_hyphens: Some(true),
                kinsoku: Some(true),
                word_wrapping: Some(true),
                overflow_punctuations: Some(true),
                top_line_punctuations: Some(true),
                auto_space_latin_and_east_asian: Some(true),
                auto_space_east_asian_and_numbers: Some(true),
                bidirectional: Some(true),
                adjust_right_indent: Some(true),
                snap_to_grid: Some(true),
                spacing: Some(Spacing::test_instance()),
                indent: Some(Ind::test_instance()),
                contextual_spacing: Some(true),
                mirror_indents: Some(true),
                suppress_overlapping: Some(true),
                alignment: Some(Jc::Start),
                text_direction: Some(TextDirection::LeftToRight),
                text_alignment: Some(TextAlignment::Auto),
                textbox_tight_wrap: Some(TextboxTightWrap::None),
                outline_level: Some(1),
                div_id: Some(1),
                conditional_formatting: Some(Cnf::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_pr_base_from_xml() {
        let xml = PPrBase::test_xml("pPrBase");
        assert_eq!(
            PPrBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PPrBase::test_instance(),
        );
    }

    impl PPrGeneral {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                PPrBase::test_extension_xml(),
                PPrChange::test_xml("pPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PPrBase::test_instance(),
                change: Some(PPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_pr_general_from_xml() {
        let xml = PPrGeneral::test_xml("pPrGeneral");
        assert_eq!(
            PPrGeneral::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PPrGeneral::test_instance(),
        );
    }

    impl ParaRPrTrackChanges {
        pub fn test_xml() -> String {
            format!(
                r#"{}
                {}
                {}
                {}
            "#,
                TrackChange::test_xml("ins"),
                TrackChange::test_xml("del"),
                TrackChange::test_xml("moveFrom"),
                TrackChange::test_xml("moveTo"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                inserted: Some(TrackChange::test_instance()),
                deleted: Some(TrackChange::test_instance()),
                move_from: Some(TrackChange::test_instance()),
                move_to: Some(TrackChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_track_changes_from_xml() {
        let xml = format!("<node>{}</node>", ParaRPrTrackChanges::test_xml());
        assert_eq!(
            ParaRPrTrackChanges::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Some(ParaRPrTrackChanges::test_instance()),
        );
    }

    impl ParaRPrOriginal {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                ParaRPrTrackChanges::test_xml(),
                RPrBase::test_run_style_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                track_changes: Some(ParaRPrTrackChanges::test_instance()),
                bases: vec![RPrBase::test_run_style_instance()],
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_original_from_xml() {
        let xml = ParaRPrOriginal::test_xml("paraRPrOriginal");
        assert_eq!(
            ParaRPrOriginal::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ParaRPrOriginal::test_instance(),
        );
    }

    impl ParaRPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0" w:author="John Smith" w:date="2001-10-26T21:32:52">
                {}
            </{node_name}>"#,
                ParaRPrOriginal::test_xml("rPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                run_properties: ParaRPrOriginal::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_para_r_change_from_xml() {
        let xml = ParaRPrChange::test_xml("paraRPrChange");
        assert_eq!(
            ParaRPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ParaRPrChange::test_instance(),
        );
    }

    impl ParaRPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                ParaRPrTrackChanges::test_xml(),
                RPrBase::test_run_style_xml(),
                ParaRPrChange::test_xml("rPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                track_changes: Some(ParaRPrTrackChanges::test_instance()),
                bases: vec![RPrBase::test_run_style_instance()],
                change: Some(ParaRPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_from_xml() {
        let xml = ParaRPr::test_xml("paraRPr");
        assert_eq!(
            ParaRPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ParaRPr::test_instance(),
        );
    }

    impl HdrFtrRef {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} r:id="rId1" w:type="default"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Rel::test_instance(),
                header_footer_type: HdrFtr::Default,
            }
        }
    }

    #[test]
    pub fn test_hdr_ftr_ref_from_xml() {
        let xml = HdrFtrRef::test_xml("hdrFtrRef");
        assert_eq!(
            HdrFtrRef::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            HdrFtrRef::test_instance(),
        );
    }

    impl NumFmt {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="decimal" w:format="&#x30A2;"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: NumberFormat::Decimal,
                format: Some(String::from("&#x30A2;")),
            }
        }
    }

    #[test]
    pub fn test_num_fmt_from_xml() {
        let xml = NumFmt::test_xml("numFmt");
        assert_eq!(
            NumFmt::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            NumFmt::test_instance(),
        );
    }

    impl FtnEdnNumProps {
        pub fn test_xml() -> String {
            format!(
                r#"<numStart w:val="1" />
            <numRestart w:val="continuous" />
            "#
            )
        }

        pub fn test_instance() -> Self {
            Self {
                numbering_start: Some(1),
                numbering_restart: Some(RestartNumber::Continuous),
            }
        }
    }

    #[test]
    pub fn test_ftn_edn_num_props_from_xml() {
        let xml = format!("<node>{}</node>", FtnEdnNumProps::test_xml());
        assert_eq!(
            FtnEdnNumProps::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Some(FtnEdnNumProps::test_instance()),
        );
    }

    impl FtnProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"<pos w:val="pageBottom" />{}{}"#,
                NumFmt::test_xml("numFmt"),
                FtnEdnNumProps::test_xml()
            )
        }

        pub fn test_instance() -> Self {
            Self {
                position: Some(FtnPos::PageBottom),
                numbering_format: Some(NumFmt::test_instance()),
                numbering_properties: Some(FtnEdnNumProps::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_ftn_props_from_xml() {
        let xml = FtnProps::test_xml("ftnProps");
        assert_eq!(
            FtnProps::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FtnProps::test_instance(),
        );
    }

    impl EdnProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"<pos w:val="docEnd" />
                {}
                {}"#,
                NumFmt::test_xml("numFmt"),
                FtnEdnNumProps::test_xml(),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                position: Some(EdnPos::DocumentEnd),
                numbering_format: Some(NumFmt::test_instance()),
                numbering_properties: Some(FtnEdnNumProps::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_edn_props_from_xml() {
        let xml = EdnProps::test_xml("endProps");
        assert_eq!(
            EdnProps::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            EdnProps::test_instance(),
        );
    }

    impl PageSz {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:w="100" w:h="100" w:orient="portrait" w:code="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
                height: Some(TwipsMeasure::Decimal(100)),
                orientation: Some(PageOrientation::Portrait),
                code: Some(1),
            }
        }
    }

    #[test]
    pub fn test_page_sz_from_xml() {
        let xml = PageSz::test_xml("pageSz");
        assert_eq!(
            PageSz::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PageSz::test_instance(),
        );
    }

    impl PageMar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:top="10" w:right="10" w:bottom="10" w:left="10" w:header="10" w:footer="10" w:gutter="10">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: SignedTwipsMeasure::Decimal(10),
                right: TwipsMeasure::Decimal(10),
                bottom: SignedTwipsMeasure::Decimal(10),
                left: TwipsMeasure::Decimal(10),
                header: TwipsMeasure::Decimal(10),
                footer: TwipsMeasure::Decimal(10),
                gutter: TwipsMeasure::Decimal(10),
            }
        }
    }

    #[test]
    pub fn test_page_mar_from_xml() {
        let xml = PageMar::test_xml("pageMar");
        assert_eq!(
            PageMar::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PageMar::test_instance(),
        );
    }

    impl PaperSource {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:first="1" w:other="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first: Some(1),
                other: Some(1),
            }
        }
    }

    #[test]
    pub fn test_paper_source_from_xml() {
        let xml = PaperSource::test_xml("paperSource");
        assert_eq!(
            PaperSource::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PaperSource::test_instance(),
        );
    }

    impl PageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:id="rId1"></{node_name}>"#,
                Border::TEST_ATTRIBUTES,
                node_name = node_name
            )
        }

        pub fn test_attributes() -> String {
            format!(r#"{} r:id="rId1""#, Border::TEST_ATTRIBUTES)
        }

        pub fn test_instance() -> Self {
            Self {
                base: Border::test_instance(),
                rel_id: Some(RelationshipId::from("rId1")),
            }
        }
    }

    #[test]
    pub fn test_page_border_from_xml() {
        let xml = PageBorder::test_xml("pageBorder");
        assert_eq!(
            PageBorder::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PageBorder::test_instance(),
        );
    }

    impl TopPageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:topLeft="rId2" r:topRight="rId3"></{node_name}>"#,
                PageBorder::test_attributes(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PageBorder::test_instance(),
                top_left: Some(RelationshipId::from("rId2")),
                top_right: Some(RelationshipId::from("rId3")),
            }
        }
    }

    #[test]
    pub fn test_top_page_border_from_xml() {
        let xml = TopPageBorder::test_xml("topPageBorder");
        assert_eq!(
            TopPageBorder::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TopPageBorder::test_instance(),
        );
    }

    impl BottomPageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:bottomLeft="rId2" r:bottomRight="rId3"></{node_name}>"#,
                PageBorder::test_attributes(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PageBorder::test_instance(),
                bottom_left: Some(RelationshipId::from("rId2")),
                bottom_right: Some(RelationshipId::from("rId3")),
            }
        }
    }

    #[test]
    pub fn test_bottom_page_border_from_xml() {
        let xml = BottomPageBorder::test_xml("bottomPageBorder");
        assert_eq!(
            BottomPageBorder::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            BottomPageBorder::test_instance(),
        );
    }

    impl PageBorders {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:zOrder="front" w:display="allPages" w:offsetFrom="page">
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TopPageBorder::test_xml("top"),
                PageBorder::test_xml("left"),
                BottomPageBorder::test_xml("bottom"),
                PageBorder::test_xml("right"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(TopPageBorder::test_instance()),
                left: Some(PageBorder::test_instance()),
                bottom: Some(BottomPageBorder::test_instance()),
                right: Some(PageBorder::test_instance()),
                z_order: Some(PageBorderZOrder::Front),
                display: Some(PageBorderDisplay::AllPages),
                offset_from: Some(PageBorderOffset::Page),
            }
        }
    }

    #[test]
    pub fn test_page_borders_from_xml() {
        let xml = PageBorders::test_xml("pageBorders");
        assert_eq!(
            PageBorders::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PageBorders::test_instance(),
        );
    }

    impl LineNumber {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:countBy="1" w:start="1" w:distance="100" w:restart="newPage"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                count_by: Some(1),
                start: Some(1),
                distance: Some(TwipsMeasure::Decimal(100)),
                restart: Some(LineNumberRestart::NewPage),
            }
        }
    }

    #[test]
    pub fn test_line_number_from_xml() {
        let xml = LineNumber::test_xml("lineNumber");
        assert_eq!(
            LineNumber::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            LineNumber::test_instance(),
        );
    }

    impl PageNumber {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:fmt="decimal" w:start="1" w:chapStyle="1" w:chapSep="hyphen"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                format: Some(NumberFormat::Decimal),
                start: Some(1),
                chapter_style: Some(1),
                chapter_separator: Some(ChapterSep::Hyphen),
            }
        }
    }

    #[test]
    pub fn test_page_number_from_xml() {
        let xml = PageNumber::test_xml("pageNumber");
        assert_eq!(
            PageNumber::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PageNumber::test_instance(),
        );
    }

    impl Column {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:w="100" w:space="10"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
                spacing: Some(TwipsMeasure::Decimal(10)),
            }
        }
    }

    #[test]
    pub fn test_column_from_xml() {
        let xml = Column::test_xml("column");
        assert_eq!(
            Column::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Column::test_instance(),
        );
    }

    impl Columns {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:equalWidth="true" w:space="10" w:num="2" w:sep="true">
                {col}
                {col}
            </{node_name}>"#,
                col = Column::test_xml("col"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                columns: vec![Column::test_instance(), Column::test_instance()],
                equal_width: Some(true),
                spacing: Some(TwipsMeasure::Decimal(10)),
                number: Some(2),
                separator: Some(true),
            }
        }
    }

    #[test]
    pub fn test_columns_from_xml() {
        let xml = Columns::test_xml("columns");
        assert_eq!(
            Columns::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Columns::test_instance(),
        );
    }

    impl DocGrid {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="default" w:linePitch="1" w:charSpace="10"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                doc_grid_type: Some(DocGridType::Default),
                line_pitch: Some(1),
                char_spacing: Some(10),
            }
        }
    }

    #[test]
    pub fn test_doc_grid_from_xml() {
        let xml = DocGrid::test_xml("docGrid");
        assert_eq!(
            DocGrid::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocGrid::test_instance(),
        );
    }

    impl SectPrContents {
        pub fn test_xml() -> String {
            format!(
                r#"{}
                {}
                <type w:val="nextPage" />
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                <formProt w:val="false" />
                <vAlign w:val="top" />
                <noEndnote w:val="false" />
                <titlePg w:val="true" />
                <textDirection w:val="lr" />
                <bidi w:val="false" />
                <rtlGutter w:val="false" />
                {}
                <printerSettings r:id="rId1" />"#,
                FtnProps::test_xml("footnotePr"),
                EdnProps::test_xml("endnotePr"),
                PageSz::test_xml("pgSz"),
                PageMar::test_xml("pgMar"),
                PaperSource::test_xml("paperSrc"),
                PageBorders::test_xml("pgBorders"),
                LineNumber::test_xml("lnNumType"),
                PageNumber::test_xml("pgNumType"),
                Columns::test_xml("cols"),
                DocGrid::test_xml("docGrid"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                footnote_properties: Some(FtnProps::test_instance()),
                endnote_properties: Some(EdnProps::test_instance()),
                section_type: Some(SectionMark::NextPage),
                page_size: Some(PageSz::test_instance()),
                page_margin: Some(PageMar::test_instance()),
                paper_source: Some(PaperSource::test_instance()),
                page_borders: Some(PageBorders::test_instance()),
                line_number_type: Some(LineNumber::test_instance()),
                page_number_type: Some(PageNumber::test_instance()),
                columns: Some(Columns::test_instance()),
                protect_form_fields: Some(false),
                vertical_align: Some(VerticalJc::Top),
                no_endnote: Some(false),
                title_page: Some(true),
                text_direction: Some(TextDirection::LeftToRight),
                bidirectional: Some(false),
                rtl_gutter: Some(false),
                document_grid: Some(DocGrid::test_instance()),
                printer_settings: Some(Rel::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_contents_from_xml() {
        let xml = format!(r#"<node>{}</node>"#, SectPrContents::test_xml());
        assert_eq!(
            SectPrContents::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Some(SectPrContents::test_instance()),
        );
    }

    impl SectPrAttributes {
        const TEST_ATTRIBUTES: &'static str =
            r#"w:rsidRPr="ffffffff" w:rsidDel="fefefefe" w:rsidR="fdfdfdfd" w:rsidSect="fcfcfcfc""#;

        pub fn test_instance() -> Self {
            Self {
                run_properties_revision_id: Some(0xffffffff),
                deletion_revision_id: Some(0xfefefefe),
                run_revision_id: Some(0xfdfdfdfd),
                section_revision_id: Some(0xfcfcfcfc),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_attributes_from_xml() {
        let xml = format!(r#"<node {}></node>"#, SectPrAttributes::TEST_ATTRIBUTES);
        assert_eq!(
            SectPrAttributes::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SectPrAttributes::test_instance(),
        );
    }

    impl SectPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                SectPrAttributes::TEST_ATTRIBUTES,
                SectPrContents::test_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                contents: Some(SectPrContents::test_instance()),
                attributes: SectPrAttributes::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_base_from_xml() {
        let xml = SectPrBase::test_xml("sectPrBase");
        assert_eq!(
            SectPrBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SectPrBase::test_instance(),
        );
    }

    impl SectPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                SectPrBase::test_xml("sectPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                section_properties: Some(SectPrBase::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_change_from_xml() {
        let xml = SectPrChange::test_xml("sectPrChange");
        assert_eq!(
            SectPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SectPrChange::test_instance(),
        );
    }

    impl SectPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {header_ref}
                {footer_ref}
                {}
                {}
            </{node_name}>"#,
                SectPrAttributes::TEST_ATTRIBUTES,
                SectPrContents::test_xml(),
                SectPrChange::test_xml("sectPrChange"),
                node_name = node_name,
                header_ref = HdrFtrRef::test_xml("headerReference"),
                footer_ref = HdrFtrRef::test_xml("footerReference"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                header_footer_references: vec![
                    HdrFtrReferences::Header(HdrFtrRef::test_instance()),
                    HdrFtrReferences::Footer(HdrFtrRef::test_instance()),
                ],
                contents: Some(SectPrContents::test_instance()),
                change: Some(SectPrChange::test_instance()),
                attributes: SectPrAttributes::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_from_xml() {
        let xml = SectPr::test_xml("sectPr");
        assert_eq!(
            SectPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SectPr::test_instance(),
        );
    }

    impl PPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                PPrBase::test_xml("pPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: PPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_p_pr_change_from_xml() {
        let xml = PPrChange::test_xml("pPrChange");
        assert_eq!(
            PPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PPrChange::test_instance(),
        );
    }

    impl PPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                PPrBase::test_extension_xml(),
                ParaRPr::test_xml("rPr"),
                SectPr::test_xml("sectPr"),
                PPrChange::test_xml("pPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PPrBase::test_instance(),
                run_properties: Some(ParaRPr::test_instance()),
                section_properties: Some(SectPr::test_instance()),
                properties_change: Some(PPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_pr_from_xml() {
        let xml = PPr::test_xml("pPr");
        assert_eq!(
            PPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            PPr::test_instance(),
        );
    }

    impl P {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:rsidRPr="ffffffff" w:rsidR="fefefefe" w:rsidDel="fdfdfdfd" w:rsidP="fcfcfcfc" w:rsidRDefault="fbfbfbfb">
                {}
                {}
            </{node_name}>"#,
                PPr::test_xml("pPr"),
                PContent::test_simple_field_xml(),
                node_name=node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(PPr::test_instance()),
                contents: vec![PContent::test_simple_field_instance()],
                run_properties_revision_id: Some(0xffffffff),
                run_revision_id: Some(0xfefefefe),
                deletion_revision_id: Some(0xfdfdfdfd),
                paragraph_revision_id: Some(0xfcfcfcfc),
                run_default_revision_id: Some(0xfbfbfbfb),
            }
        }
    }

    #[test]
    pub fn test_p_from_xml() {
        let xml = P::test_xml("p");
        assert_eq!(
            P::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            P::test_instance(),
        );
    }

    #[test]
    pub fn test_decimal_number_or_percent_from_str() {
        assert_eq!(
            "123".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Decimal(123),
        );
        assert_eq!(
            "-123".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Decimal(-123),
        );
        assert_eq!(
            "123%".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Percentage(Percentage(123.0)),
        );
        assert_eq!(
            "-123%".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Percentage(Percentage(-123.0)),
        );

        match DecimalNumberOrPercent::Percentage(Percentage(100.0)) {
            DecimalNumberOrPercent::Percentage(Percentage(value)) => println!("{}", value),
            _ => (),
        }
    }

    #[test]
    pub fn test_measurement_or_percent_from_str() {
        assert_eq!(
            "123".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(123)),
        );
        assert_eq!(
            "-123".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(-123)),
        );
        assert_eq!(
            "123%".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Percentage(Percentage(123.0))),
        );
        assert_eq!(
            "-123%".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Percentage(Percentage(-123.0))),
        );
        assert_eq!(
            "123mm".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
        );
        assert_eq!(
            "-123mm".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::UniversalMeasure(UniversalMeasure::new(-123.0, UniversalMeasureUnit::Millimeter)),
        );
    }

    impl AltChunkPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <matchSrc w:val="true" />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                match_source: Some(true),
            }
        }
    }

    #[test]
    pub fn test_alt_chunk_pr_from_xml() {
        let xml = AltChunkPr::test_xml("altChunkPr");
        assert_eq!(
            AltChunkPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            AltChunkPr::test_instance(),
        );
    }

    impl AltChunk {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} r:id="rId1">{}</{node_name}>"#,
                AltChunkPr::test_xml("altChunkPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(AltChunkPr::test_instance()),
                rel_id: Some(RelationshipId::from("rId1")),
            }
        }
    }

    #[test]
    pub fn test_alt_chunk_from_xml() {
        let xml = AltChunk::test_xml("altChunk");
        assert_eq!(
            AltChunk::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            AltChunk::test_instance(),
        );
    }

    impl Background {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:color="ffffff" w:themeColor="light1" w:themeTint="ff" w:themeShade="ff">
                {}
            </{node_name}>"#,
                Drawing::test_xml("drawing"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                drawing: Some(Drawing::test_instance()),
                color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
                theme_color: Some(ThemeColor::Light1),
                theme_tint: Some(0xff),
                theme_shade: Some(0xff),
            }
        }
    }

    #[test]
    pub fn test_background_from_xml() {
        let xml = Background::test_xml("background");
        assert_eq!(
            Background::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Background::test_instance(),
        );
    }

    impl DocumentBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            Background::test_xml("background")
        }

        pub fn test_instance() -> Self {
            Self {
                background: Some(Background::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_document_base_from_xml() {
        let xml = DocumentBase::test_xml("documentBase");
        assert_eq!(
            DocumentBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocumentBase::test_instance(),
        );
    }

    impl Body {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                P::test_xml("p"),
                SectPr::test_xml("sectPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                block_level_elements: vec![BlockLevelElts::Chunk(ContentBlockContent::Paragraph(Box::new(
                    P::test_instance(),
                )))],
                section_properties: Some(SectPr::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_body_from_xml() {
        let xml = Body::test_xml("body");
        assert_eq!(
            Body::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Body::test_instance(),
        );
    }

    impl Document {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:conformance="transitional">
                {}
                {}
            </{node_name}>"#,
                DocumentBase::test_extension_xml(),
                Body::test_xml("body"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: DocumentBase::test_instance(),
                body: Some(Body::test_instance()),
                conformance: Some(ConformanceClass::Transitional),
            }
        }
    }

    #[test]
    pub fn test_document_from_xml() {
        let xml = Document::test_xml("document");
        assert_eq!(
            Document::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Document::test_instance(),
        );
    }
}
