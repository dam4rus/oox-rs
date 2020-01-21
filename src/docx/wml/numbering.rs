use super::{
    document::{Control, Drawing, Jc, NumFmt, PPrGeneral, RPr, Rel},
    simpletypes::{parse_on_off_xml_element, DecimalNumber, LongHexNumber},
    util::XmlNodeExt,
};
use log::info;
use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    shared::sharedtypes::OnOff,
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::any::Any;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Default)]
pub struct PictureBase {
    pub vml_element: Option<Box<dyn Any>>,
    pub office_element: Option<Box<dyn Any>>,
}

#[derive(Debug, Default)]
pub struct Picture {
    pub base: PictureBase,
    pub movie: Option<Rel>,
    pub control: Option<Control>,
}

impl Picture {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Picture");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "movie" => instance.movie = Some(Rel::from_xml_element(child_node)?),
                    "control" => instance.control = Some(Control::from_xml_element(child_node)),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug)]
pub enum NumPicBulletChoice {
    Drawing(Drawing),
    Picture(Picture),
}

impl XsdType for NumPicBulletChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing NumPicBulletChoice");

        match xml_node.local_name() {
            "drawing" => Ok(NumPicBulletChoice::Drawing(Drawing::from_xml_element(xml_node)?)),
            "pict" => Ok(NumPicBulletChoice::Picture(Picture::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "NumPicBulletChoice",
            ))),
        }
    }
}

impl XsdChoice for NumPicBulletChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "drawing" | "pict" => true,
            _ => false,
        }
    }
}

#[derive(Debug)]
pub struct NumPicBullet {
    pub choice: NumPicBulletChoice,
    pub symbol_id: DecimalNumber,
}

impl NumPicBullet {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing NumPicBullet");

        let choice = xml_node
            .child_nodes
            .iter()
            .find_map(NumPicBulletChoice::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "w:drawing|w:pict"))?;

        let symbol_id = xml_node
            .attributes
            .get("w:numPicBulletId")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:numPicBulletId"))?
            .parse()?;

        Ok(Self { choice, symbol_id })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum MultiLevelType {
    #[strum(serialize = "singleLevel")]
    SingleLevel,
    #[strum(serialize = "multilevel")]
    MultiLevel,
    #[strum(serialize = "hybridMultilevel")]
    HybridMultiLevel,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum LevelSuffix {
    #[strum(serialize = "tab")]
    Tab,
    #[strum(serialize = "space")]
    Space,
    #[strum(serialize = "nothing")]
    Nothing,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct LevelText {
    pub value: Option<String>,
    pub is_null: Option<OnOff>,
}

impl LevelText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing LevelText");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, attr_value)| {
                match attr.as_ref() {
                    "w:val" => instance.value = Some(attr_value.clone()),
                    "w:null" => instance.is_null = Some(parse_xml_bool(attr_value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Lvl {
    pub start: Option<DecimalNumber>,
    pub numbering_format: Option<NumFmt>,
    pub level_restart: Option<DecimalNumber>,
    pub paragraph_style: Option<String>,
    pub display_as_arabic_numerals: Option<OnOff>,
    pub suffix: Option<LevelSuffix>,
    pub level_text: Option<LevelText>,
    pub level_picture_bullet_id: Option<DecimalNumber>,
    pub level_alignment: Option<Jc>,
    pub paragraph_properties: Option<PPrGeneral>,
    pub run_properties: Option<RPr>,
    pub level: DecimalNumber,
    pub template_code: Option<LongHexNumber>,
    pub tentative: Option<OnOff>,
}

impl Lvl {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Lvl");

        let mut level = None;
        let mut template_code = None;
        let mut tentative = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:ilvl" => level = Some(value.parse()?),
                "w:tplc" => template_code = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:tentative" => tentative = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let level = level.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:ilvl"))?;

        let mut start = None;
        let mut numbering_format = None;
        let mut level_restart = None;
        let mut paragraph_style = None;
        let mut display_as_arabic_numerals = None;
        let mut suffix = None;
        let mut level_text = None;
        let mut level_picture_bullet_id = None;
        let mut level_alignment = None;
        let mut paragraph_properties = None;
        let mut run_properties = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "start" => start = Some(child_node.get_val_attribute()?.parse()?),
                "numFmt" => numbering_format = Some(NumFmt::from_xml_element(child_node)?),
                "lvlRestart" => level_restart = Some(child_node.get_val_attribute()?.parse()?),
                "pStyle" => paragraph_style = Some(child_node.get_val_attribute()?.clone()),
                "isLgl" => display_as_arabic_numerals = Some(parse_on_off_xml_element(child_node)?),
                "suff" => suffix = Some(child_node.get_val_attribute()?.parse()?),
                "lvlText" => level_text = Some(LevelText::from_xml_element(child_node)?),
                "lvlPicBulletId" => level_picture_bullet_id = Some(child_node.get_val_attribute()?.parse()?),
                "lvlJc" => level_alignment = Some(child_node.get_val_attribute()?.parse()?),
                "pPr" => paragraph_properties = Some(PPrGeneral::from_xml_element(child_node)?),
                "rPr" => run_properties = Some(RPr::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            start,
            numbering_format,
            level_restart,
            paragraph_style,
            display_as_arabic_numerals,
            suffix,
            level_text,
            level_picture_bullet_id,
            level_alignment,
            paragraph_properties,
            run_properties,
            level,
            template_code,
            tentative,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AbstractNum {
    pub definition_id: Option<LongHexNumber>,
    pub multi_level_type: Option<MultiLevelType>,
    pub template: Option<LongHexNumber>,
    pub name: Option<String>,
    pub style_link: Option<String>,
    pub numbering_style_link: Option<String>,
    pub levels: Vec<Lvl>,
    pub abstract_num_id: DecimalNumber,
}

impl AbstractNum {
    pub fn new(abstract_num_id: DecimalNumber) -> Self {
        Self {
            definition_id: None,
            multi_level_type: None,
            template: None,
            name: None,
            style_link: None,
            numbering_style_link: None,
            levels: Vec::new(),
            abstract_num_id,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing AbstractNum");

        let abstract_num_id = xml_node
            .attributes
            .get("w:abstractNumId")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:abstractNumId"))?
            .parse()?;

        xml_node
            .child_nodes
            .iter()
            .try_fold(Self::new(abstract_num_id), |mut instance, child_node| {
                match child_node.local_name() {
                    "nsid" => {
                        let val_attr = child_node.get_val_attribute()?;
                        instance.definition_id = Some(LongHexNumber::from_str_radix(val_attr, 16)?)
                    }
                    "multiLevelType" => instance.multi_level_type = Some(child_node.get_val_attribute()?.parse()?),
                    "tmpl" => {
                        let val_attr = child_node.get_val_attribute()?;
                        instance.template = Some(LongHexNumber::from_str_radix(val_attr, 16)?)
                    }
                    "name" => instance.name = Some(child_node.get_val_attribute()?.clone()),
                    "styleLink" => instance.style_link = Some(child_node.get_val_attribute()?.clone()),
                    "numStyleLink" => instance.numbering_style_link = Some(child_node.get_val_attribute()?.clone()),
                    "lvl" => instance.levels.push(Lvl::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|instance| match instance.levels.len() {
                0..=9 => Ok(instance),
                len => Err(Box::new(LimitViolationError::new(
                    xml_node.name.clone(),
                    "w:lvl",
                    0,
                    MaxOccurs::Value(9),
                    len as u32,
                ))
                .into()),
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct NumLvl {
    pub start_override: Option<DecimalNumber>,
    pub level: Option<Lvl>,
    pub numbering_level: DecimalNumber,
}

impl NumLvl {
    pub fn new(numbering_level: DecimalNumber) -> Self {
        Self {
            start_override: None,
            level: None,
            numbering_level,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing NumLvl");

        let numbering_level = xml_node
            .attributes
            .get("w:ilvl")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:ilvl"))?
            .parse()?;

        xml_node
            .child_nodes
            .iter()
            .try_fold(Self::new(numbering_level), |mut instance, child_node| {
                match child_node.local_name() {
                    "startOverride" => instance.start_override = Some(child_node.get_val_attribute()?.parse()?),
                    "lvl" => instance.level = Some(Lvl::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Num {
    pub abstract_num_id: DecimalNumber,
    pub level_overrides: Vec<NumLvl>,
    pub numbering_id: DecimalNumber,
}

impl Num {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Num");

        let numbering_id = xml_node
            .attributes
            .get("w:numId")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:numId"))?
            .parse()?;

        let mut abstract_num_id = None;
        let mut level_overrides = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "abstractNumId" => abstract_num_id = Some(child_node.get_val_attribute()?.parse()?),
                "lvlOverride" => level_overrides.push(NumLvl::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let abstract_num_id =
            abstract_num_id.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "w:abstractNumId"))?;

        match level_overrides.len() {
            0..=9 => Ok(Self {
                abstract_num_id,
                level_overrides,
                numbering_id,
            }),
            len => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "w:lvlOverride",
                0,
                MaxOccurs::Value(9),
                len as u32,
            ))),
        }
    }
}

#[derive(Debug, Default)]
pub struct Numbering {
    pub picture_numbering_symbols: Vec<NumPicBullet>,
    pub abstract_numberings: Vec<AbstractNum>,
    pub numberings: Vec<Num>,
    pub numbering_id_mac_at_cleanup: Option<DecimalNumber>,
}

impl Numbering {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Numbering");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "numPicBullet" => instance
                        .picture_numbering_symbols
                        .push(NumPicBullet::from_xml_element(child_node)?),
                    "abstractNum" => instance
                        .abstract_numberings
                        .push(AbstractNum::from_xml_element(child_node)?),
                    "num" => instance.numberings.push(Num::from_xml_element(child_node)?),
                    "numIdMacAtCleanup" => {
                        instance.numbering_id_mac_at_cleanup = Some(child_node.get_val_attribute()?.parse()?)
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

    impl Picture {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                Rel::test_xml("w:movie"),
                Control::test_xml("w:control"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Default::default(),
                movie: Some(Rel::test_instance()),
                control: Some(Control::test_instance()),
            }
        }
    }

    #[test]
    fn test_picture_from_xml() {
        let xml = Picture::test_xml("w:numPicBullet");
        let picture = Picture::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        let test_instance = Picture::test_instance();
        assert_eq!(picture.movie, test_instance.movie);
        assert_eq!(picture.control, test_instance.control);
    }

    impl LevelText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="Example" w:null="false"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(String::from("Example")),
                is_null: Some(false),
            }
        }
    }

    #[test]
    fn test_level_text_from_xml() {
        let xml = LevelText::test_xml("w:levelText");
        assert_eq!(
            LevelText::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            LevelText::test_instance()
        );
    }

    impl Lvl {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:ilvl="1" w:tplc="ffffffff" w:tentative="false">
                <w:start w:val="1" />
                {}
                <w:lvlRestart w:val="1" />
                <w:pStyle w:val="Example" />
                <w:isLgl w:val="false" />
                <w:suff w:val="nothing" />
                {}
                <w:lvlPicBulletId w:val="1" />
                <w:lvlJc w:val="start" />
                {}
                {}
            </{node_name}>"#,
                NumFmt::test_xml("w:numFmt"),
                LevelText::test_xml("w:lvlText"),
                PPrGeneral::test_xml("w:pPr"),
                RPr::test_xml("w:rPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start: Some(1),
                numbering_format: Some(NumFmt::test_instance()),
                level_restart: Some(1),
                paragraph_style: Some(String::from("Example")),
                display_as_arabic_numerals: Some(false),
                suffix: Some(LevelSuffix::Nothing),
                level_text: Some(LevelText::test_instance()),
                level_picture_bullet_id: Some(1),
                level_alignment: Some(Jc::Start),
                paragraph_properties: Some(PPrGeneral::test_instance()),
                run_properties: Some(RPr::test_instance()),
                level: 1,
                template_code: Some(0xffffffff),
                tentative: Some(false),
            }
        }
    }

    #[test]
    fn test_lvl_from_xml() {
        let xml = Lvl::test_xml("w:lvl");
        assert_eq!(
            Lvl::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Lvl::test_instance()
        );
    }

    impl AbstractNum {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:abstractNumId="1">
                <w:nsid w:val="ffffffff" />
                <w:multiLevelType w:val="singleLevel" />
                <w:tmpl w:val="fefefefe" />
                <w:name w:val="Example" />
                <w:styleLink w:val="Example" />
                <w:numStyleLink w:val="Example" />
                {lvl}
                {lvl}
            </{node_name}>"#,
                lvl = Lvl::test_xml("w:lvl"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                definition_id: Some(0xffffffff),
                multi_level_type: Some(MultiLevelType::SingleLevel),
                template: Some(0xfefefefe),
                name: Some(String::from("Example")),
                style_link: Some(String::from("Example")),
                numbering_style_link: Some(String::from("Example")),
                levels: vec![Lvl::test_instance(), Lvl::test_instance()],
                abstract_num_id: 1,
            }
        }
    }

    #[test]
    fn test_abstract_num_from_xml() {
        let xml = AbstractNum::test_xml("w:abstractNum");
        assert_eq!(
            AbstractNum::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            AbstractNum::test_instance()
        );
    }

    impl NumLvl {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:ilvl="1">
                <startOverride w:val="1" />
                {}
            </{node_name}>"#,
                Lvl::test_xml("w:lvl"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start_override: Some(1),
                level: Some(Lvl::test_instance()),
                numbering_level: 1,
            }
        }
    }

    #[test]
    fn test_num_lvl_from_xml() {
        let xml = NumLvl::test_xml("w:numLvl");
        assert_eq!(
            NumLvl::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            NumLvl::test_instance()
        );
    }

    impl Num {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:numId="1">
                <w:abstractNumId w:val="1" />
                {lvl_override}
                {lvl_override}
            </{node_name}>"#,
                lvl_override = NumLvl::test_xml("w:lvlOverride"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                abstract_num_id: 1,
                level_overrides: vec![NumLvl::test_instance(), NumLvl::test_instance()],
                numbering_id: 1,
            }
        }
    }

    #[test]
    fn test_num_from_xml() {
        let xml = Num::test_xml("w:num");
        assert_eq!(
            Num::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Num::test_instance()
        );
    }

    impl Numbering {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                <w:numIdMacAtCleanup w:val="1" />
            </{node_name}>"#,
                AbstractNum::test_xml("w:abstractNum"),
                Num::test_xml("w:num"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                picture_numbering_symbols: Vec::new(),
                abstract_numberings: vec![AbstractNum::test_instance()],
                numberings: vec![Num::test_instance()],
                numbering_id_mac_at_cleanup: Some(1),
            }
        }
    }

    #[test]
    fn test_numbering_from_xml() {
        let xml = Numbering::test_xml("w:numbering");
        let numbering = Numbering::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        let test_instance = Numbering::test_instance();
        assert_eq!(numbering.abstract_numberings, test_instance.abstract_numberings);
        assert_eq!(numbering.numberings, test_instance.numberings);
        assert_eq!(
            numbering.numbering_id_mac_at_cleanup,
            test_instance.numbering_id_mac_at_cleanup
        );
    }
}
