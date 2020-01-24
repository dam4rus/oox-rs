use super::{
    document::{PPr, PPrGeneral, RPr},
    simpletypes::{parse_on_off_xml_element, DecimalNumber, LongHexNumber},
    table::{TblPrBase, TcPr, TrPr},
    util::XmlNodeExt,
};
use crate::{
    error::MissingAttributeError,
    shared::sharedtypes::OnOff,
    xml::{parse_xml_bool, XmlNode},
};
use log::info;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPrDefault(pub Option<RPr>);

impl RPrDefault {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing RPrDefault");

        let run_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .map(RPr::from_xml_element)
            .transpose()?;

        Ok(Self(run_properties))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPrDefault(pub Option<PPr>);

impl PPrDefault {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing PPrDefault");

        let paragraph_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pPr")
            .map(PPr::from_xml_element)
            .transpose()?;

        Ok(Self(paragraph_properties))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocDefaults {
    pub run_properties_default: Option<RPrDefault>,
    pub paragraph_properties_default: Option<PPrDefault>,
}

impl DocDefaults {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing DocDefaults");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "rPrDefault" => instance.run_properties_default = Some(RPrDefault::from_xml_element(child_node)?),
                    "pPrDefault" => {
                        instance.paragraph_properties_default = Some(PPrDefault::from_xml_element(child_node)?)
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct LsdException {
    pub name: String,
    pub locked: Option<OnOff>,
    pub ui_priority: Option<DecimalNumber>,
    pub semi_hidden: Option<OnOff>,
    pub unhide_when_used: Option<OnOff>,
    pub primary_style: Option<OnOff>,
}

impl LsdException {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing LsdException");

        let mut name = None;
        let mut locked = None;
        let mut ui_priority = None;
        let mut semi_hidden = None;
        let mut unhide_when_used = None;
        let mut primary_style = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:name" => name = Some(value.clone()),
                "w:locked" => locked = Some(parse_xml_bool(value)?),
                "w:uiPriority" => ui_priority = Some(value.parse()?),
                "w:semiHidden" => semi_hidden = Some(parse_xml_bool(value)?),
                "w:unhideWhenUsed" => unhide_when_used = Some(parse_xml_bool(value)?),
                "w:qFormat" => primary_style = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        Ok(Self {
            name,
            locked,
            ui_priority,
            semi_hidden,
            unhide_when_used,
            primary_style,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct LatentStyles {
    pub lsd_exceptions: Vec<LsdException>,
    pub default_locked_state: Option<OnOff>,
    pub default_ui_priority: Option<DecimalNumber>,
    pub default_semi_hidden: Option<OnOff>,
    pub default_unhide_when_used: Option<OnOff>,
    pub default_primary_style: Option<OnOff>,
    pub count: Option<DecimalNumber>,
}

impl LatentStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing LatentStyles");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:defLockedState" => instance.default_locked_state = Some(parse_xml_bool(value)?),
                "w:defUIPriority" => instance.default_ui_priority = Some(value.parse()?),
                "w:defSemiHidden" => instance.default_semi_hidden = Some(parse_xml_bool(value)?),
                "w:defUnhideWhenUsed" => instance.default_unhide_when_used = Some(parse_xml_bool(value)?),
                "w:defQFormat" => instance.default_primary_style = Some(parse_xml_bool(value)?),
                "w:count" => instance.count = Some(value.parse()?),
                _ => (),
            }
        }

        instance.lsd_exceptions = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "lsdException")
            .map(LsdException::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(instance)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TblStyleOverrideType {
    #[strum(serialize = "wholeTable")]
    WholeTable,
    #[strum(serialize = "firstRow")]
    FirstRow,
    #[strum(serialize = "lastRow")]
    LastRow,
    #[strum(serialize = "firstCol")]
    FirstColumn,
    #[strum(serialize = "lastCol")]
    LastColumn,
    #[strum(serialize = "band1Vert")]
    Band1Vertical,
    #[strum(serialize = "band2Vert")]
    Band2Vertical,
    #[strum(serialize = "band1Horz")]
    Band1Horizontal,
    #[strum(serialize = "band2Horz")]
    Band2Horizontal,
    #[strum(serialize = "neCell")]
    NorthEastCell,
    #[strum(serialize = "nwCell")]
    NorthWestCell,
    #[strum(serialize = "seCell")]
    SouthEastCell,
    #[strum(serialize = "swCell")]
    SouthWestCell,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblStylePr {
    pub paragraph_properties: Option<PPrGeneral>,
    pub run_properties: Option<RPr>,
    pub table_properties: Option<TblPrBase>,
    pub table_row_properties: Option<TrPr>,
    pub table_cell_properties: Option<TcPr>,
    pub override_type: TblStyleOverrideType,
}

impl TblStylePr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing TblStylePr");

        let override_type = xml_node
            .attributes
            .get("w:type")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?
            .parse()?;

        let initial_state = Self {
            paragraph_properties: None,
            run_properties: None,
            table_properties: None,
            table_row_properties: None,
            table_cell_properties: None,
            override_type,
        };

        xml_node
            .child_nodes
            .iter()
            .try_fold(initial_state, |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "pPr" => instance.paragraph_properties = Some(PPrGeneral::from_xml_element(child_node)?),
                    "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                    "tblPr" => instance.table_properties = Some(TblPrBase::from_xml_element(child_node)?),
                    "trPr" => instance.table_row_properties = Some(TrPr::from_xml_element(child_node)?),
                    "tcPr" => instance.table_cell_properties = Some(TcPr::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum StyleType {
    #[strum(serialize = "paragraph")]
    Paragraph,
    #[strum(serialize = "character")]
    Character,
    #[strum(serialize = "table")]
    Table,
    #[strum(serialize = "numbering")]
    Numbering,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Style {
    pub name: Option<String>,
    pub aliases: Option<String>,
    pub based_on: Option<String>,
    pub next: Option<String>,
    pub link: Option<String>,
    pub auto_redefine: Option<OnOff>,
    pub hidden: Option<OnOff>,
    pub ui_priority: Option<DecimalNumber>,
    pub semi_hidden: Option<OnOff>,
    pub unhide_when_used: Option<OnOff>,
    pub primary_style: Option<OnOff>,
    pub locked: Option<OnOff>,
    pub personal: Option<OnOff>,
    pub personal_compose: Option<OnOff>,
    pub personal_reply: Option<OnOff>,
    pub revision_id: Option<LongHexNumber>,
    pub paragraph_properties: Option<PPrGeneral>,
    pub run_properties: Option<RPr>,
    pub table_properties: Option<TblPrBase>,
    pub table_row_properties: Option<TrPr>,
    pub table_cell_properties: Option<TcPr>,
    pub table_style_properties_vec: Vec<TblStylePr>,
    pub style_type: Option<StyleType>,
    pub style_id: Option<String>,
    pub is_default: Option<OnOff>,
    pub custom_style: Option<OnOff>,
}

impl Style {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:type" => instance.style_type = Some(value.parse()?),
                "w:styleId" => instance.style_id = Some(value.clone()),
                "w:default" => instance.is_default = Some(parse_xml_bool(value)?),
                "w:customStyle" => instance.custom_style = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "name" => instance.name = Some(child_node.get_val_attribute()?.clone()),
                "aliases" => instance.aliases = Some(child_node.get_val_attribute()?.clone()),
                "basedOn" => instance.based_on = Some(child_node.get_val_attribute()?.clone()),
                "next" => instance.next = Some(child_node.get_val_attribute()?.clone()),
                "link" => instance.link = Some(child_node.get_val_attribute()?.clone()),
                "autoRedefine" => instance.auto_redefine = Some(parse_on_off_xml_element(child_node)?),
                "hidden" => instance.hidden = Some(parse_on_off_xml_element(child_node)?),
                "uiPriority" => instance.ui_priority = Some(child_node.get_val_attribute()?.parse()?),
                "semiHidden" => instance.semi_hidden = Some(parse_on_off_xml_element(child_node)?),
                "unhideWhenUsed" => instance.unhide_when_used = Some(parse_on_off_xml_element(child_node)?),
                "qFormat" => instance.primary_style = Some(parse_on_off_xml_element(child_node)?),
                "locked" => instance.locked = Some(parse_on_off_xml_element(child_node)?),
                "personal" => instance.personal = Some(parse_on_off_xml_element(child_node)?),
                "personalCompose" => instance.personal_compose = Some(parse_on_off_xml_element(child_node)?),
                "personalReply" => instance.personal_reply = Some(parse_on_off_xml_element(child_node)?),
                "rsid" => {
                    instance.revision_id = Some(LongHexNumber::from_str_radix(child_node.get_val_attribute()?, 16)?)
                }
                "pPr" => instance.paragraph_properties = Some(PPrGeneral::from_xml_element(child_node)?),
                "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                "tblPr" => instance.table_properties = Some(TblPrBase::from_xml_element(child_node)?),
                "trPr" => instance.table_row_properties = Some(TrPr::from_xml_element(child_node)?),
                "tcPr" => instance.table_cell_properties = Some(TcPr::from_xml_element(child_node)?),
                "tblStylePr" => instance
                    .table_style_properties_vec
                    .push(TblStylePr::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Styles {
    pub document_defaults: Option<DocDefaults>,
    pub latent_styles: Option<LatentStyles>,
    pub styles: Vec<Style>,
}

impl Styles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "docDefaults" => instance.document_defaults = Some(DocDefaults::from_xml_element(child_node)?),
                    "latentStyles" => instance.latent_styles = Some(LatentStyles::from_xml_element(child_node)?),
                    "style" => instance.styles.push(Style::from_xml_element(child_node)?),
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

    impl DocDefaults {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <rPrDefault>
                    {}
                </rPrDefault>
                <pPrDefault>
                    {}
                </pPrDefault>
            </{node_name}>"#,
                RPr::test_xml("rPr"),
                PPr::test_xml("pPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties_default: Some(RPrDefault(Some(RPr::test_instance()))),
                paragraph_properties_default: Some(PPrDefault(Some(PPr::test_instance()))),
            }
        }
    }

    #[test]
    pub fn test_doc_defaults_from_xml() {
        let xml = DocDefaults::test_xml("docDefaults");
        assert_eq!(
            DocDefaults::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocDefaults::test_instance()
        );
    }

    impl LsdException {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#""<{node_name} w:name="Normal" w:locked="false" w:uiPriority="1" w:semiHidden="false"
                w:unhideWhenUsed="false" w:qFormat="false"></{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: String::from("Normal"),
                locked: Some(false),
                ui_priority: Some(1),
                semi_hidden: Some(false),
                unhide_when_used: Some(false),
                primary_style: Some(false),
            }
        }
    }

    #[test]
    pub fn test_lsd_exception_from_xml() {
        let xml = LsdException::test_xml("lsdException");
        assert_eq!(
            LsdException::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            LsdException::test_instance()
        );
    }

    impl LatentStyles {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:defLockedState="false" w:defUIPriority="1" w:defSemiHidden="false"
                w:defUnhideWhenUsed="false" w:defQFormat="false" w:count="1">
                {}
            </{node_name}>"#,
                LsdException::test_xml("lsdException"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                lsd_exceptions: vec![LsdException::test_instance()],
                default_locked_state: Some(false),
                default_ui_priority: Some(1),
                default_semi_hidden: Some(false),
                default_unhide_when_used: Some(false),
                default_primary_style: Some(false),
                count: Some(1),
            }
        }
    }

    #[test]
    pub fn test_latent_styles_from_xml() {
        let xml = LatentStyles::test_xml("latentStyles");
        assert_eq!(
            LatentStyles::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            LatentStyles::test_instance()
        );
    }

    impl TblStylePr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="wholeTable">
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                PPrGeneral::test_xml("pPr"),
                RPr::test_xml("rPr"),
                TblPrBase::test_xml("tblPr"),
                TrPr::test_xml("trPr"),
                TcPr::test_xml("tcPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                paragraph_properties: Some(PPrGeneral::test_instance()),
                run_properties: Some(RPr::test_instance()),
                table_properties: Some(TblPrBase::test_instance()),
                table_row_properties: Some(TrPr::test_instance()),
                table_cell_properties: Some(TcPr::test_instance()),
                override_type: TblStyleOverrideType::WholeTable,
            }
        }
    }

    #[test]
    pub fn test_tbl_style_pr_from_xml() {
        let xml = TblStylePr::test_xml("tblStylePr");
        assert_eq!(
            TblStylePr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblStylePr::test_instance()
        );
    }

    impl Style {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:type="paragraph" w:styleId="Title" w:default="false" w:customStyle="false">
                <name w:val="Title" />
                <aliases w:val="Title" />
                <basedOn w:val="Normal" />
                <next w:val="Normal" />
                <link w:val="Heading1Char" />
                <autoRedefine w:val="false" />
                <hidden w:val="false" />
                <uiPriority w:val="1" />
                <semiHidden w:val="false" />
                <unhideWhenUsed w:val="false" />
                <qFormat w:val="true" />
                <locked w:val="false" />
                <personal w:val="false" />
                <personalCompose w:val="false" />
                <personalReply w:val="false" />
                <rsid w:val="ffffffff" />
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                PPrGeneral::test_xml("pPr"),
                RPr::test_xml("rPr"),
                TblPrBase::test_xml("tblPr"),
                TrPr::test_xml("trPr"),
                TcPr::test_xml("tcPr"),
                TblStylePr::test_xml("tblStylePr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: Some(String::from("Title")),
                aliases: Some(String::from("Title")),
                based_on: Some(String::from("Normal")),
                next: Some(String::from("Normal")),
                link: Some(String::from("Heading1Char")),
                auto_redefine: Some(false),
                hidden: Some(false),
                ui_priority: Some(1),
                semi_hidden: Some(false),
                unhide_when_used: Some(false),
                primary_style: Some(true),
                locked: Some(false),
                personal: Some(false),
                personal_compose: Some(false),
                personal_reply: Some(false),
                revision_id: Some(0xffffffff),
                paragraph_properties: Some(PPrGeneral::test_instance()),
                run_properties: Some(RPr::test_instance()),
                table_properties: Some(TblPrBase::test_instance()),
                table_row_properties: Some(TrPr::test_instance()),
                table_cell_properties: Some(TcPr::test_instance()),
                table_style_properties_vec: vec![TblStylePr::test_instance()],
                style_type: Some(StyleType::Paragraph),
                style_id: Some(String::from("Title")),
                is_default: Some(false),
                custom_style: Some(false),
            }
        }
    }

    #[test]
    pub fn test_style_from_xml() {
        let xml = Style::test_xml("style");
        assert_eq!(
            Style::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Style::test_instance()
        );
    }

    impl Styles {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                DocDefaults::test_xml("docDefaults"),
                LatentStyles::test_xml("latentStyles"),
                Style::test_xml("style"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                document_defaults: Some(DocDefaults::test_instance()),
                latent_styles: Some(LatentStyles::test_instance()),
                styles: vec![Style::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_styles_from_xml() {
        let xml = Styles::test_xml("styles");
        assert_eq!(
            Styles::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Styles::test_instance()
        );
    }
}
