use super::{
    document::{ChapterSep, DecimalNumberOrPercent, EdnProps, FtnProps, Language, NumberFormat, Rel},
    simpletypes::{parse_on_off_xml_element, DecimalNumber, LongHexNumber, UnsignedDecimalNumber},
    util::XmlNodeExt,
};
use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError},
    shared::{
        drawingml::simpletypes::Lang,
        relationship::RelationshipId,
        sharedtypes::{OnOff, TwipsMeasure},
    },
    xml::{parse_xml_bool, XmlNode},
};
use log::info;

pub type Base64Binary = String;
pub type DocType = String;
pub type MailMergeDataType = String;
pub type PixelsMeasure = UnsignedDecimalNumber;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Password {
    pub algorithm_name: Option<String>,
    pub hash_value: Option<Base64Binary>,
    pub salt_value: Option<Base64Binary>,
    pub spin_count: Option<DecimalNumber>,
}

impl Password {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing Password");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_attribute)
    }

    pub fn try_update_from_xml_attribute(mut self, (attr, value): (&String, &String)) -> Result<Self> {
        match attr.as_ref() {
            "w:algorithmName" => self.algorithm_name = Some(value.clone()),
            "w:hashValue" => self.hash_value = Some(value.clone()),
            "w:saltValue" => self.salt_value = Some(value.clone()),
            "w:spinCount" => self.spin_count = Some(value.parse()?),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WriteProtection {
    pub recommended: Option<OnOff>,
    pub password: Password,
}

impl WriteProtection {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing WriteProtection");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:recommended" => instance.recommended = Some(parse_xml_bool(value)?),
                    _ => instance.password = instance.password.try_update_from_xml_attribute((attr, value))?,
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum View {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "print")]
    Print,
    #[strum(serialize = "outline")]
    Outline,
    #[strum(serialize = "masterPages")]
    MasterPages,
    #[strum(serialize = "normal")]
    Normal,
    #[strum(serialize = "web")]
    Web,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ZoomType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "fullPage")]
    FullPage,
    #[strum(serialize = "bestFit")]
    BestFit,
    #[strum(serialize = "textFit")]
    TextFit,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Zoom {
    pub value: Option<ZoomType>,
    pub percent: DecimalNumberOrPercent,
}

impl Zoom {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing Zoom");

        let mut value = None;
        let mut percent = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => value = Some(attr_value.parse()?),
                "w:percent" => percent = Some(attr_value.parse()?),
                _ => (),
            }
        }

        let percent = percent.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "percent"))?;

        Ok(Self { value, percent })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WritingStyle {
    pub language: Lang,
    pub vendor_id: String,
    pub dll_version: String,
    pub natural_language_check: Option<OnOff>,
    pub check_style: OnOff,
    pub app_name: String,
}

impl WritingStyle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing WritingStyle");

        let mut language = None;
        let mut vendor_id = None;
        let mut dll_version = None;
        let mut natural_language_check = None;
        let mut check_style = None;
        let mut app_name = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:lang" => language = Some(value.clone()),
                "w:vendorID" => vendor_id = Some(value.clone()),
                "w:dllVersion" => dll_version = Some(value.clone()),
                "w:nlCheck" => natural_language_check = Some(parse_xml_bool(value)?),
                "w:checkStyle" => check_style = Some(parse_xml_bool(value)?),
                "w:appName" => app_name = Some(value.clone()),
                _ => (),
            }
        }

        let language = language.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "lang"))?;
        let vendor_id = vendor_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "vendorID"))?;
        let dll_version = dll_version.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "dllVersion"))?;
        let check_style = check_style.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "checkStyle"))?;
        let app_name = app_name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "appName"))?;

        Ok(Self {
            language,
            vendor_id,
            dll_version,
            natural_language_check,
            check_style,
            app_name,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct StylePaneFilter {
    pub all_styles: Option<OnOff>,
    pub custom_styles: Option<OnOff>,
    pub latent_styles: Option<OnOff>,
    pub styles_in_use: Option<OnOff>,
    pub heading_styles: Option<OnOff>,
    pub numbering_styles: Option<OnOff>,
    pub table_styles: Option<OnOff>,
    pub direct_formatting_on_runs: Option<OnOff>,
    pub direct_formatting_on_paragraphs: Option<OnOff>,
    pub direct_formatting_on_numbering: Option<OnOff>,
    pub direct_formatting_on_tables: Option<OnOff>,
    pub clear_formatting: Option<OnOff>,
    pub top_three_heading_styles: Option<OnOff>,
    pub visible_styles: Option<OnOff>,
    pub alternate_style_names: Option<OnOff>,
}

impl StylePaneFilter {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing StylePaneFilter");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:allStyles" => instance.all_styles = Some(parse_xml_bool(value)?),
                    "w:customStyles" => instance.custom_styles = Some(parse_xml_bool(value)?),
                    "w:latentStyles" => instance.latent_styles = Some(parse_xml_bool(value)?),
                    "w:stylesInUse" => instance.styles_in_use = Some(parse_xml_bool(value)?),
                    "w:headingStyles" => instance.heading_styles = Some(parse_xml_bool(value)?),
                    "w:numberingStyles" => instance.numbering_styles = Some(parse_xml_bool(value)?),
                    "w:tableStyles" => instance.table_styles = Some(parse_xml_bool(value)?),
                    "w:directFormattingOnRuns" => instance.direct_formatting_on_runs = Some(parse_xml_bool(value)?),
                    "w:directFormattingOnParagraphs" => {
                        instance.direct_formatting_on_paragraphs = Some(parse_xml_bool(value)?)
                    }
                    "w:directFormattingOnNumbering" => {
                        instance.direct_formatting_on_numbering = Some(parse_xml_bool(value)?)
                    }
                    "w:directFormattingOnTables" => instance.direct_formatting_on_tables = Some(parse_xml_bool(value)?),
                    "w:clearFormatting" => instance.clear_formatting = Some(parse_xml_bool(value)?),
                    "w:top3HeadingStyles" => instance.top_three_heading_styles = Some(parse_xml_bool(value)?),
                    "w:visibleStyles" => instance.visible_styles = Some(parse_xml_bool(value)?),
                    "w:alternateStyleNames" => instance.alternate_style_names = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum StyleSort {
    #[strum(serialize = "name")]
    Name,
    #[strum(serialize = "priority")]
    Priority,
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "font")]
    Font,
    #[strum(serialize = "basedOn")]
    BasedOn,
    #[strum(serialize = "type")]
    Type,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ProofType {
    #[strum(serialize = "clean")]
    Clean,
    #[strum(serialize = "dirty")]
    Dirty,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Proof {
    pub spelling: Option<ProofType>,
    pub grammar: Option<ProofType>,
}

impl Proof {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing Proof");

        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:spelling" => instance.spelling = Some(value.parse()?),
                    "w:grammar" => instance.grammar = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum MailMergeDocType {
    #[strum(serialize = "catalog")]
    Catalog,
    #[strum(serialize = "envelopes")]
    Envelopes,
    #[strum(serialize = "mailingLabels")]
    MailingLabels,
    #[strum(serialize = "formLetters")]
    FormLetter,
    #[strum(serialize = "email")]
    Email,
    #[strum(serialize = "fax")]
    Fax,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum MailMergeDest {
    #[strum(serialize = "newDocument")]
    NewDocument,
    #[strum(serialize = "printer")]
    Printer,
    #[strum(serialize = "email")]
    Email,
    #[strum(serialize = "fax")]
    Fax,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum MailMergeSourceType {
    #[strum(serialize = "database")]
    Database,
    #[strum(serialize = "addressBook")]
    AddressBook,
    #[strum(serialize = "document1")]
    Document1,
    #[strum(serialize = "document2")]
    Document2,
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "email")]
    Email,
    #[strum(serialize = "native")]
    Native,
    #[strum(serialize = "legacy")]
    Legacy,
    #[strum(serialize = "master")]
    Master,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum MailMergeOdsoFMDFieldType {
    #[strum(serialize = "null")]
    Null,
    #[strum(serialize = "dbColumn")]
    DatabaseColumn,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct OdsoFieldMapData {
    pub field_type: Option<MailMergeOdsoFMDFieldType>,
    pub name: Option<String>,
    pub mapped_name: Option<String>,
    pub column: Option<DecimalNumber>,
    pub language_id: Option<Lang>,
    pub dynamic_address: Option<OnOff>,
}

impl OdsoFieldMapData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing OdsoFieldMapData");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "type" => instance.field_type = Some(child_node.get_val_attribute()?.parse()?),
                    "name" => instance.name = Some(child_node.get_val_attribute()?.clone()),
                    "mappedName" => instance.mapped_name = Some(child_node.get_val_attribute()?.clone()),
                    "column" => instance.column = Some(child_node.get_val_attribute()?.parse()?),
                    "lid" => instance.language_id = Some(child_node.get_val_attribute()?.clone()),
                    "dynamicAddress" => instance.dynamic_address = Some(parse_on_off_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Odso {
    pub udl: Option<String>,
    pub table: Option<String>,
    pub source: Option<Rel>,
    pub column_delimiter: Option<DecimalNumber>,
    pub mail_merge_source_type: Option<MailMergeSourceType>,
    pub first_header: Option<OnOff>,
    pub field_map_datas: Vec<OdsoFieldMapData>,
    pub recipient_datas: Vec<Rel>,
}

impl Odso {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "udl" => instance.udl = Some(child_node.get_val_attribute()?.clone()),
                    "table" => instance.table = Some(child_node.get_val_attribute()?.clone()),
                    "src" => instance.source = Some(Rel::from_xml_element(child_node)?),
                    "colDelim" => instance.column_delimiter = Some(child_node.get_val_attribute()?.parse()?),
                    "type" => instance.mail_merge_source_type = Some(child_node.get_val_attribute()?.parse()?),
                    "fHdr" => instance.first_header = Some(parse_on_off_xml_element(child_node)?),
                    "fieldMapData" => instance
                        .field_map_datas
                        .push(OdsoFieldMapData::from_xml_element(child_node)?),
                    "recipientData" => instance.recipient_datas.push(Rel::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct MailMerge {
    pub main_document_type: MailMergeDocType,
    pub link_to_query: Option<OnOff>,
    pub data_type: MailMergeDataType,
    pub connect_string: Option<String>,
    pub query: Option<String>,
    pub data_source: Option<Rel>,
    pub header_source: Option<Rel>,
    pub do_not_suppress_blank_lines: Option<OnOff>,
    pub destination: Option<MailMergeDest>,
    pub address_field_name: Option<String>,
    pub mail_subject: Option<String>,
    pub mail_as_attachment: Option<OnOff>,
    pub view_merged_data: Option<OnOff>,
    pub active_record: Option<DecimalNumber>,
    pub check_errors: Option<DecimalNumber>,
    pub odso: Option<Odso>,
}

impl MailMerge {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut main_document_type = None;
        let mut link_to_query = None;
        let mut data_type = None;
        let mut connect_string = None;
        let mut query = None;
        let mut data_source = None;
        let mut header_source = None;
        let mut do_not_suppress_blank_lines = None;
        let mut destination = None;
        let mut address_field_name = None;
        let mut mail_subject = None;
        let mut mail_as_attachment = None;
        let mut view_merged_data = None;
        let mut active_record = None;
        let mut check_errors = None;
        let mut odso = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "mainDocumentType" => main_document_type = Some(child_node.get_val_attribute()?.parse()?),
                "linkToQuery" => link_to_query = Some(parse_on_off_xml_element(child_node)?),
                "dataType" => data_type = Some(child_node.get_val_attribute()?.clone()),
                "connectString" => connect_string = Some(child_node.get_val_attribute()?.clone()),
                "query" => query = Some(child_node.get_val_attribute()?.clone()),
                "dataSource" => data_source = Some(Rel::from_xml_element(child_node)?),
                "headerSource" => header_source = Some(Rel::from_xml_element(child_node)?),
                "doNotSuppressBlankLines" => do_not_suppress_blank_lines = Some(parse_on_off_xml_element(child_node)?),
                "destination" => destination = Some(child_node.get_val_attribute()?.parse()?),
                "addressFieldName" => address_field_name = Some(child_node.get_val_attribute()?.clone()),
                "mailSubject" => mail_subject = Some(child_node.get_val_attribute()?.clone()),
                "mailAsAttachment" => mail_as_attachment = Some(parse_on_off_xml_element(child_node)?),
                "viewMergedData" => view_merged_data = Some(parse_on_off_xml_element(child_node)?),
                "activeRecord" => active_record = Some(child_node.get_val_attribute()?.parse()?),
                "checkErrors" => check_errors = Some(child_node.get_val_attribute()?.parse()?),
                "odso" => odso = Some(Odso::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let main_document_type =
            main_document_type.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "mainDocumentType"))?;

        let data_type = data_type.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "dataType"))?;

        Ok(Self {
            main_document_type,
            link_to_query,
            data_type,
            connect_string,
            query,
            data_source,
            header_source,
            do_not_suppress_blank_lines,
            destination,
            address_field_name,
            mail_subject,
            mail_as_attachment,
            view_merged_data,
            active_record,
            check_errors,
            odso,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct TrackChangesView {
    pub markup: Option<OnOff>,
    pub comments: Option<OnOff>,
    pub display_content_revisions: Option<OnOff>,
    pub formatting: Option<OnOff>,
    pub ink_annotations: Option<OnOff>,
}

impl TrackChangesView {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:markup" => instance.markup = Some(parse_xml_bool(value)?),
                    "w:comments" => instance.comments = Some(parse_xml_bool(value)?),
                    "w:insDel" => instance.display_content_revisions = Some(parse_xml_bool(value)?),
                    "w:formatting" => instance.formatting = Some(parse_xml_bool(value)?),
                    "w:inkAnnotations" => instance.ink_annotations = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum DocProtectType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "readOnly")]
    ReadOnly,
    #[strum(serialize = "comments")]
    Comments,
    #[strum(serialize = "trackedChanges")]
    TrackedChanges,
    #[strum(serialize = "forms")]
    Forms,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocProtect {
    pub edit: Option<DocProtectType>,
    pub formatting: Option<OnOff>,
    pub enforcement: Option<OnOff>,
    pub password: Password,
}

impl DocProtect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w:edit" => instance.edit = Some(value.parse()?),
                    "w:formatting" => instance.formatting = Some(value.parse()?),
                    "w:enforcement" => instance.enforcement = Some(value.parse()?),
                    _ => instance.password = instance.password.try_update_from_xml_attribute((attr, value))?,
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum CharacterSpacing {
    #[strum(serialize = "doNotCompress")]
    DoNotCompress,
    #[strum(serialize = "compressPunctuation")]
    CompressPunctuation,
    #[strum(serialize = "compressPunctuationAndJapaneseKana")]
    CompressPunctuationAndJapaneseKana,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Kinsoku {
    pub language: Lang,
    pub value: String,
}

impl Kinsoku {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut language = None;
        let mut value = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:lang" => language = Some(attr_value.clone()),
                "w:val" => value = Some(attr_value.clone()),
                _ => (),
            }
        }

        let language = language.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:lang"))?;
        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:val"))?;

        Ok(Self { language, value })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SaveThroughXslt {
    pub rel_id: Option<RelationshipId>,
    pub solution_id: Option<String>,
}

impl SaveThroughXslt {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        xml_node
            .attributes
            .iter()
            .fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "r:id" => instance.rel_id = Some(value.clone()),
                    "w:solutionID" => instance.solution_id = Some(value.clone()),
                    _ => (),
                }

                instance
            })
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FtnEndSepRef {
    pub id: DecimalNumber,
}

impl FtnEndSepRef {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let id = xml_node
            .attributes
            .get("w:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:id"))?
            .parse()?;

        Ok(Self { id })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FtnDocProps {
    pub base: FtnProps,
    pub footnotes: Vec<FtnEndSepRef>,
}

impl FtnDocProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing FtnDocProps");

        let fold_result: Result<Self> =
            xml_node
                .child_nodes
                .iter()
                .try_fold(Default::default(), |mut instance: Self, child_node| {
                    match child_node.local_name() {
                        "footnote" => instance.footnotes.push(FtnEndSepRef::from_xml_element(child_node)?),
                        _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
                    }

                    Ok(instance)
                });

        let instance = fold_result?;

        match instance.footnotes.len() {
            0..=3 => Ok(instance),
            len => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "footnote",
                0,
                MaxOccurs::Value(3),
                len as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct EdnDocProps {
    pub base: EdnProps,
    pub endnotes: Vec<FtnEndSepRef>,
}

impl EdnDocProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing EdnDocProps");

        let fold_result: Result<Self> =
            xml_node
                .child_nodes
                .iter()
                .try_fold(Default::default(), |mut instance: Self, child_node| {
                    match child_node.local_name() {
                        "endnote" => instance.endnotes.push(FtnEndSepRef::from_xml_element(child_node)?),
                        _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
                    }

                    Ok(instance)
                });

        let instance = fold_result?;

        match instance.endnotes.len() {
            0..=3 => Ok(instance),
            len => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "endnote",
                0,
                MaxOccurs::Value(3),
                len as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct CompatSetting {
    pub name: Option<String>,
    pub uri: Option<String>,
    pub value: Option<String>,
}

impl CompatSetting {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        xml_node
            .attributes
            .iter()
            .fold(Default::default(), |mut instance: Self, (attr, attr_value)| {
                match attr.as_ref() {
                    "w:name" => instance.name = Some(attr_value.clone()),
                    "w:uri" => instance.uri = Some(attr_value.clone()),
                    "w:val" => instance.value = Some(attr_value.clone()),
                    _ => (),
                }

                instance
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Compat {
    pub space_for_underline: Option<OnOff>,
    pub balance_single_byte_double_byte_width: Option<OnOff>,
    pub do_not_leave_backslash_alone: Option<OnOff>,
    pub underline_trail_space: Option<OnOff>,
    pub do_not_expand_shift_return: Option<OnOff>,
    pub adjust_line_height_in_table: Option<OnOff>,
    pub apply_breaking_rules: Option<OnOff>,
    pub compatibility_settings: Vec<CompatSetting>,
}

impl Compat {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "spaceForUL" => instance.space_for_underline = Some(parse_on_off_xml_element(child_node)?),
                    "balanceSingleByteDoubleByteWidth" => {
                        instance.balance_single_byte_double_byte_width = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "doNotLeaveBackslashAlone" => {
                        instance.do_not_leave_backslash_alone = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "ulTrailSpace" => instance.underline_trail_space = Some(parse_on_off_xml_element(child_node)?),
                    "doNotExpandShiftReturn" => {
                        instance.do_not_expand_shift_return = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "adjustLineHeightInTable" => {
                        instance.adjust_line_height_in_table = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "applyBreakingRules" => instance.apply_breaking_rules = Some(parse_on_off_xml_element(child_node)?),
                    "compatSetting" => instance
                        .compatibility_settings
                        .push(CompatSetting::from_xml_element(child_node)),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/*
<xsd:complexType name="CT_DocVar">
    <xsd:attribute name="name" type="s:ST_String" use="required"/>
    <xsd:attribute name="val" type="s:ST_String" use="required"/>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct DocVar {
    pub name: String,
    pub value: String,
}

impl DocVar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut value = None;
        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:name" => name = Some(attr_value.clone()),
                "w:val" => value = Some(attr_value.clone()),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:name"))?;
        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:val"))?;

        Ok(Self { name, value })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocVars(pub Vec<DocVar>);

impl DocVars {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let doc_vars = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "docVar")
            .map(DocVar::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(doc_vars))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocRsids {
    pub revision_id_root: Option<LongHexNumber>,
    pub revision_ids: Vec<LongHexNumber>,
}

impl DocRsids {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "rsidRoot" => {
                        instance.revision_id_root =
                            Some(LongHexNumber::from_str_radix(child_node.get_val_attribute()?, 16)?)
                    }
                    "rsid" => instance
                        .revision_ids
                        .push(LongHexNumber::from_str_radix(child_node.get_val_attribute()?, 16)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum WmlColorSchemeIndex {
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
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColorSchemeMapping {
    pub background1: WmlColorSchemeIndex,
    pub text1: WmlColorSchemeIndex,
    pub background2: WmlColorSchemeIndex,
    pub text2: WmlColorSchemeIndex,
    pub accent1: WmlColorSchemeIndex,
    pub accent2: WmlColorSchemeIndex,
    pub accent3: WmlColorSchemeIndex,
    pub accent4: WmlColorSchemeIndex,
    pub accent5: WmlColorSchemeIndex,
    pub accent6: WmlColorSchemeIndex,
    pub hyperlink: WmlColorSchemeIndex,
    pub followed_hyperlink: WmlColorSchemeIndex,
}

impl ColorSchemeMapping {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut background1 = None;
        let mut text1 = None;
        let mut background2 = None;
        let mut text2 = None;
        let mut accent1 = None;
        let mut accent2 = None;
        let mut accent3 = None;
        let mut accent4 = None;
        let mut accent5 = None;
        let mut accent6 = None;
        let mut hyperlink = None;
        let mut followed_hyperlink = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:bg1" => background1 = Some(value.parse()?),
                "w:t1" => text1 = Some(value.parse()?),
                "w:bg2" => background2 = Some(value.parse()?),
                "w:t2" => text2 = Some(value.parse()?),
                "w:accent1" => accent1 = Some(value.parse()?),
                "w:accent2" => accent2 = Some(value.parse()?),
                "w:accent3" => accent3 = Some(value.parse()?),
                "w:accent4" => accent4 = Some(value.parse()?),
                "w:accent5" => accent5 = Some(value.parse()?),
                "w:accent6" => accent6 = Some(value.parse()?),
                "w:hyperlink" => hyperlink = Some(value.parse()?),
                "w:followedHyperlink" => followed_hyperlink = Some(value.parse()?),
                _ => (),
            }
        }

        let background1 = background1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:bg1"))?;
        let text1 = text1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:t1"))?;
        let background2 = background2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:bg2"))?;
        let text2 = text2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:t2"))?;
        let accent1 = accent1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent1"))?;
        let accent2 = accent2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent2"))?;
        let accent3 = accent3.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent3"))?;
        let accent4 = accent4.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent4"))?;
        let accent5 = accent5.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent5"))?;
        let accent6 = accent6.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:accent6"))?;
        let hyperlink = hyperlink.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:hyperlink"))?;
        let followed_hyperlink = followed_hyperlink
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:followedHyperlink"))?;

        Ok(Self {
            background1,
            text1,
            background2,
            text2,
            accent1,
            accent2,
            accent3,
            accent4,
            accent5,
            accent6,
            hyperlink,
            followed_hyperlink,
        })
    }
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum CaptionPos {
    #[strum(serialize = "above")]
    Above,
    #[strum(serialize = "below")]
    Below,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Caption {
    pub name: String,
    pub position: Option<CaptionPos>,
    pub chapter_number: Option<OnOff>,
    pub heading: Option<DecimalNumber>,
    pub no_label: Option<OnOff>,
    pub numbering_format: Option<NumberFormat>,
    pub separator: Option<ChapterSep>,
}

impl Caption {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut position = None;
        let mut chapter_number = None;
        let mut heading = None;
        let mut no_label = None;
        let mut numbering_format = None;
        let mut separator = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:name" => name = Some(value.clone()),
                "w:pos" => position = Some(value.parse()?),
                "w:chapNum" => chapter_number = Some(parse_xml_bool(value)?),
                "w:heading" => heading = Some(value.parse()?),
                "w:noLabel" => no_label = Some(parse_xml_bool(value)?),
                "w:numFmt" => numbering_format = Some(value.parse()?),
                "w:sep" => separator = Some(value.parse()?),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:name"))?;

        Ok(Self {
            name,
            position,
            chapter_number,
            heading,
            no_label,
            numbering_format,
            separator,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AutoCaption {
    pub name: String,
    pub caption: String,
}

impl AutoCaption {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut caption = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:name" => name = Some(value.clone()),
                "w:caption" => caption = Some(value.clone()),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:name"))?;
        let caption = caption.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:caption"))?;

        Ok(Self { name, caption })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct AutoCaptions(pub Vec<AutoCaption>);

impl AutoCaptions {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("Parsing AutoCaptions");

        let auto_captions = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "autoCaption")
            .map(AutoCaption::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(auto_captions))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Captions {
    pub captions: Vec<Caption>,
    pub auto_captions: Option<AutoCaptions>,
}

impl Captions {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fold_result: Result<Self> =
            xml_node
                .child_nodes
                .iter()
                .try_fold(Default::default(), |mut instance: Self, child_node| {
                    match child_node.local_name() {
                        "caption" => instance.captions.push(Caption::from_xml_element(child_node)?),
                        "autoCaptions" => instance.auto_captions = Some(AutoCaptions::from_xml_element(child_node)?),
                        _ => (),
                    }

                    Ok(instance)
                });

        let instance = fold_result?;
        if !instance.captions.is_empty() {
            Ok(instance)
        } else {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "caption",
                1,
                MaxOccurs::Unbounded,
                instance.captions.len() as u32,
            )))
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ReadingModeInkLockDown {
    pub use_actual_pages: OnOff,
    pub width: PixelsMeasure,
    pub height: PixelsMeasure,
    pub font_size: DecimalNumberOrPercent,
}

impl ReadingModeInkLockDown {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut use_actual_pages = None;
        let mut width = None;
        let mut height = None;
        let mut font_size = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:actualPg" => use_actual_pages = Some(parse_xml_bool(value)?),
                "w:w" => width = Some(value.parse()?),
                "w:h" => height = Some(value.parse()?),
                "w:fontSz" => font_size = Some(value.parse()?),
                _ => (),
            }
        }

        let use_actual_pages =
            use_actual_pages.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:actualPg"))?;
        let width = width.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:w"))?;
        let height = height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:h"))?;
        let font_size = font_size.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:fontSz"))?;

        Ok(Self {
            use_actual_pages,
            width,
            height,
            font_size,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SmartTagType {
    pub namespaceuri: String,
    pub name: String,
    pub url: String,
}

impl SmartTagType {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut namespaceuri = None;
        let mut name = None;
        let mut url = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:namespaceuri" => namespaceuri = Some(value.clone()),
                "w:name" => name = Some(value.clone()),
                "w:url" => url = Some(value.clone()),
                _ => (),
            }
        }

        let namespaceuri =
            namespaceuri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:namespaceuri"))?;
        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:name"))?;
        let url = url.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "w:url"))?;

        Ok(Self {
            namespaceuri,
            name,
            url,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Settings {
    pub write_protection: Option<WriteProtection>,
    pub view: Option<View>,
    pub zoom: Option<Zoom>,
    pub remove_personal_information: Option<OnOff>,
    pub remove_date_and_time: Option<OnOff>,
    pub do_not_display_page_boundaries: Option<OnOff>,
    pub display_background_shape: Option<OnOff>,
    pub print_post_script_over_text: Option<OnOff>,
    pub print_fractional_character_width: Option<OnOff>,
    pub print_forms_data: Option<OnOff>,
    pub embed_true_type_fonts: Option<OnOff>,
    pub embed_system_fonts: Option<OnOff>,
    pub save_subset_fonts: Option<OnOff>,
    pub save_forms_data: Option<OnOff>,
    pub mirror_margins: Option<OnOff>,
    pub align_borders_and_edges: Option<OnOff>,
    pub borders_do_not_surround_header: Option<OnOff>,
    pub borders_do_not_surround_footer: Option<OnOff>,
    pub gutter_at_top: Option<OnOff>,
    pub hide_spelling_errors: Option<OnOff>,
    pub hide_grammatical_errors: Option<OnOff>,
    pub active_writing_styles: Vec<WritingStyle>,
    pub proof_state: Option<Proof>,
    pub forms_design: Option<OnOff>,
    pub attached_template: Option<Rel>,
    pub link_styles: Option<OnOff>,
    pub style_pane_format_filter: Option<StylePaneFilter>,
    pub style_pane_sort_method: Option<StyleSort>,
    pub document_type: Option<DocType>,
    pub mail_merge: Option<MailMerge>,
    pub revision_view: Option<TrackChangesView>,
    pub track_revisions: Option<OnOff>,
    pub do_not_track_moves: Option<OnOff>,
    pub do_not_track_formatting: Option<OnOff>,
    pub document_protection: Option<DocProtect>,
    pub auto_format_override: Option<OnOff>,
    pub style_lock_theme: Option<OnOff>,
    pub style_lock_set: Option<OnOff>,
    pub default_tab_stop: Option<TwipsMeasure>,
    pub auto_hyphenation: Option<OnOff>,
    pub consecutive_hyphen_limit: Option<DecimalNumber>,
    pub hyphenation_zone: Option<TwipsMeasure>,
    pub do_not_hyphenate_capitals: Option<OnOff>,
    pub show_envelope: Option<OnOff>,
    pub summary_length: Option<DecimalNumberOrPercent>,
    pub click_and_type_style: Option<String>,
    pub default_table_style: Option<String>,
    pub even_and_odd_headers: Option<OnOff>,
    pub book_fold_revision_printing: Option<OnOff>,
    pub book_fold_printing: Option<OnOff>,
    pub book_fold_printing_sheets: Option<OnOff>,
    pub drawing_grid_horizontal_spacing: Option<TwipsMeasure>,
    pub drawing_grid_vertical_spacing: Option<TwipsMeasure>,
    pub display_horizontal_drawing_grid_every: Option<DecimalNumber>,
    pub display_vertical_drawing_grid_every: Option<DecimalNumber>,
    pub do_not_use_margins_for_drawing_grid_origin: Option<OnOff>,
    pub drawing_grid_horizontal_origin: Option<TwipsMeasure>,
    pub drawing_grid_vertical_origin: Option<TwipsMeasure>,
    pub do_not_shade_form_data: Option<OnOff>,
    pub no_punctuation_kerning: Option<OnOff>,
    pub character_spacing_control: Option<CharacterSpacing>,
    pub print_two_on_one: Option<OnOff>,
    pub strict_first_and_last_chars: Option<OnOff>,
    pub no_line_breaks_after: Option<Kinsoku>,
    pub no_line_breaks_before: Option<Kinsoku>,
    pub save_preview_picture: Option<OnOff>,
    pub do_not_validate_against_schema: Option<OnOff>,
    pub save_invalid_xml: Option<OnOff>,
    pub ignore_mixed_content: Option<OnOff>,
    pub always_show_placeholder_text: Option<OnOff>,
    pub do_not_demarcate_invalid_xml: Option<OnOff>,
    pub save_xml_data_only: Option<OnOff>,
    pub use_xslt_when_saving: Option<OnOff>,
    pub save_through_xslt: Option<SaveThroughXslt>,
    pub show_xml_tags: Option<OnOff>,
    pub always_merge_empty_namespace: Option<OnOff>,
    pub update_fields: Option<OnOff>,
    pub footnote_properties: Option<FtnDocProps>,
    pub endnote_properties: Option<EdnDocProps>,
    pub compatibility: Option<Compat>,
    pub document_variables: Option<DocVars>,
    pub revision_ids: Option<DocRsids>,
    //       <xsd:element ref="m:mathPr" minOccurs="0" maxOccurs="1"/>
    // TODO(kalmar.robert): Implement
    pub attached_schemas: Vec<String>,
    pub theme_font_lang: Option<Language>,
    pub color_scheme_mapping: Option<ColorSchemeMapping>,
    pub do_not_include_subdocs_in_stats: Option<OnOff>,
    pub do_not_auto_compress_pictures: Option<OnOff>,
    pub force_upgrade: bool,
    pub captions: Option<Captions>,
    pub read_move_ink_lock_down: Option<ReadingModeInkLockDown>,
    pub smart_tag_types: Vec<SmartTagType>,
    //       <xsd:element ref="sl:schemaLibrary" minOccurs="0" maxOccurs="1"/>
    // TODO(kalmar.robert): Implement
    pub do_not_embed_smart_tags: Option<OnOff>,
    pub decimal_symbol: Option<String>,
    pub list_separator: Option<String>,
}

impl Settings {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "writeProtection" => {
                        instance.write_protection = Some(WriteProtection::from_xml_element(child_node)?)
                    }
                    "view" => instance.view = Some(child_node.get_val_attribute()?.parse()?),
                    "zoom" => instance.zoom = Some(Zoom::from_xml_element(child_node)?),
                    "removePersonalInformation" => {
                        instance.remove_personal_information = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "removeDateAndTime" => instance.remove_date_and_time = Some(parse_on_off_xml_element(child_node)?),
                    "doNotDisplayPageBoundaries" => {
                        instance.do_not_display_page_boundaries = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "displayBackgroundShape" => {
                        instance.display_background_shape = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "printPostScriptOverText" => {
                        instance.print_post_script_over_text = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "printFractionalCharacterWidth" => {
                        instance.print_fractional_character_width = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "printFormsData" => instance.print_forms_data = Some(parse_on_off_xml_element(child_node)?),
                    "embedTrueTypeFonts" => {
                        instance.embed_true_type_fonts = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "embedSystemFonts" => instance.embed_system_fonts = Some(parse_on_off_xml_element(child_node)?),
                    "saveSubsetFonts" => instance.save_subset_fonts = Some(parse_on_off_xml_element(child_node)?),
                    "saveFormsData" => instance.save_forms_data = Some(parse_on_off_xml_element(child_node)?),
                    "mirrorMargins" => instance.mirror_margins = Some(parse_on_off_xml_element(child_node)?),
                    "alignBordersAndEdges" => {
                        instance.align_borders_and_edges = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "bordersDoNotSurroundHeader" => {
                        instance.borders_do_not_surround_header = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "bordersDoNotSurroundFooter" => {
                        instance.borders_do_not_surround_footer = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "gutterAtTop" => instance.gutter_at_top = Some(parse_on_off_xml_element(child_node)?),
                    "hideSpellingErrors" => instance.hide_spelling_errors = Some(parse_on_off_xml_element(child_node)?),
                    "hideGrammaticalErrors" => {
                        instance.hide_grammatical_errors = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "activeWritingStyle" => instance
                        .active_writing_styles
                        .push(WritingStyle::from_xml_element(child_node)?),
                    "proofState" => instance.proof_state = Some(Proof::from_xml_element(child_node)?),
                    "formsDesign" => instance.forms_design = Some(parse_on_off_xml_element(child_node)?),
                    "attachedTemplate" => instance.attached_template = Some(Rel::from_xml_element(child_node)?),
                    "linkStyles" => instance.link_styles = Some(parse_on_off_xml_element(child_node)?),
                    "stylePaneFormatFilter" => {
                        instance.style_pane_format_filter = Some(StylePaneFilter::from_xml_element(child_node)?)
                    }
                    "stylePaneSortMethod" => {
                        instance.style_pane_sort_method = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "documentType" => instance.document_type = Some(child_node.get_val_attribute()?.clone()),
                    "mailMerge" => instance.mail_merge = Some(MailMerge::from_xml_element(child_node)?),
                    "revisionView" => instance.revision_view = Some(TrackChangesView::from_xml_element(child_node)?),
                    "trackRevisions" => instance.track_revisions = Some(parse_on_off_xml_element(child_node)?),
                    "doNotTrackMoves" => instance.do_not_track_moves = Some(parse_on_off_xml_element(child_node)?),
                    "doNotTrackFormatting" => {
                        instance.do_not_track_formatting = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "documentProtection" => {
                        instance.document_protection = Some(DocProtect::from_xml_element(child_node)?)
                    }
                    "autoFormatOverride" => instance.auto_format_override = Some(parse_on_off_xml_element(child_node)?),
                    "styleLockTheme" => instance.style_lock_theme = Some(parse_on_off_xml_element(child_node)?),
                    "styleLockQFSet" => instance.style_lock_set = Some(parse_on_off_xml_element(child_node)?),
                    "defaultTabStop" => instance.default_tab_stop = Some(child_node.get_val_attribute()?.parse()?),
                    "autoHyphenation" => instance.auto_hyphenation = Some(parse_on_off_xml_element(child_node)?),
                    "consecutiveHyphenLimit" => {
                        instance.consecutive_hyphen_limit = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "hyphenationZone" => instance.hyphenation_zone = Some(child_node.get_val_attribute()?.parse()?),
                    "doNotHyphenateCaps" => {
                        instance.do_not_hyphenate_capitals = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "showEnvelope" => instance.show_envelope = Some(parse_on_off_xml_element(child_node)?),
                    "summaryLength" => instance.summary_length = Some(child_node.get_val_attribute()?.parse()?),
                    "clickAndTypeStyle" => {
                        instance.click_and_type_style = Some(child_node.get_val_attribute()?.clone())
                    }
                    "defaultTableStyle" => instance.default_table_style = Some(child_node.get_val_attribute()?.clone()),
                    "evenAndOddHeaders" => instance.even_and_odd_headers = Some(parse_on_off_xml_element(child_node)?),
                    "bookFoldRevPrinting" => {
                        instance.book_fold_revision_printing = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "bookFoldPrinting" => instance.book_fold_printing = Some(parse_on_off_xml_element(child_node)?),
                    "bookFoldPrintingSheets" => {
                        instance.book_fold_printing_sheets = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "drawingGridHorizontalSpacing" => {
                        instance.drawing_grid_horizontal_spacing = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "drawingGridVerticalSpacing" => {
                        instance.drawing_grid_vertical_spacing = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "displayHorizontalDrawingGridEvery" => {
                        instance.display_horizontal_drawing_grid_every = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "displayVerticalDrawingGridEvery" => {
                        instance.display_vertical_drawing_grid_every = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "doNotUseMarginsForDrawingGridOrigin" => {
                        instance.do_not_use_margins_for_drawing_grid_origin =
                            Some(parse_on_off_xml_element(child_node)?)
                    }
                    "drawingGridHorizontalOrigin" => {
                        instance.drawing_grid_horizontal_origin = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "drawingGridVerticalOrigin" => {
                        instance.drawing_grid_vertical_origin = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "doNotShadeFormData" => {
                        instance.do_not_shade_form_data = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "noPunctuationKerning" => {
                        instance.no_punctuation_kerning = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "characterSpacingControl" => {
                        instance.character_spacing_control = Some(child_node.get_val_attribute()?.parse()?)
                    }
                    "printTwoOnOne" => instance.print_two_on_one = Some(parse_on_off_xml_element(child_node)?),
                    "strictFirstAndLastChars" => {
                        instance.strict_first_and_last_chars = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "noLineBreaksAfter" => instance.no_line_breaks_after = Some(Kinsoku::from_xml_element(child_node)?),
                    "noLineBreaksBefore" => {
                        instance.no_line_breaks_before = Some(Kinsoku::from_xml_element(child_node)?)
                    }
                    "savePreviewPicture" => instance.save_preview_picture = Some(parse_on_off_xml_element(child_node)?),
                    "doNotValidateAgainstSchema" => {
                        instance.do_not_validate_against_schema = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "saveInvalidXml" => instance.save_invalid_xml = Some(parse_on_off_xml_element(child_node)?),
                    "ignoreMixedContent" => instance.ignore_mixed_content = Some(parse_on_off_xml_element(child_node)?),
                    "alwaysShowPlaceholderText" => {
                        instance.always_show_placeholder_text = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "doNotDemarcateInvalidXml" => {
                        instance.do_not_demarcate_invalid_xml = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "saveXmlDataOnly" => instance.save_xml_data_only = Some(parse_on_off_xml_element(child_node)?),
                    "useXSLTWhenSaving" => instance.use_xslt_when_saving = Some(parse_on_off_xml_element(child_node)?),
                    "saveThroughXslt" => {
                        instance.save_through_xslt = Some(SaveThroughXslt::from_xml_element(child_node))
                    }
                    "showXMLTags" => instance.show_xml_tags = Some(parse_on_off_xml_element(child_node)?),
                    "alwaysMergeEmptyNamespace" => {
                        instance.always_merge_empty_namespace = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "updateFields" => instance.update_fields = Some(parse_on_off_xml_element(child_node)?),
                    "footnotePr" => instance.footnote_properties = Some(FtnDocProps::from_xml_element(child_node)?),
                    "endnotePr" => instance.endnote_properties = Some(EdnDocProps::from_xml_element(child_node)?),
                    "compat" => instance.compatibility = Some(Compat::from_xml_element(child_node)?),
                    "docVars" => instance.document_variables = Some(DocVars::from_xml_element(child_node)?),
                    "rsids" => instance.revision_ids = Some(DocRsids::from_xml_element(child_node)?),
                    "attachedSchema" => instance.attached_schemas.push(child_node.get_val_attribute()?.clone()),
                    "themeFontLang" => instance.theme_font_lang = Some(Language::from_xml_element(child_node)),
                    "clrSchemeMapping" => {
                        instance.color_scheme_mapping = Some(ColorSchemeMapping::from_xml_element(child_node)?)
                    }
                    "doNotIncludeSubdocsInStats" => {
                        instance.do_not_include_subdocs_in_stats = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "doNotAutoCompressPictures" => {
                        instance.do_not_auto_compress_pictures = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "forceUpgrade" => instance.force_upgrade = true,
                    "captions" => instance.captions = Some(Captions::from_xml_element(child_node)?),
                    "readModeInkLockDown" => {
                        instance.read_move_ink_lock_down = Some(ReadingModeInkLockDown::from_xml_element(child_node)?)
                    }
                    "smartTagType" => instance
                        .smart_tag_types
                        .push(SmartTagType::from_xml_element(child_node)?),
                    "doNotEmbedSmartTags" => {
                        instance.do_not_embed_smart_tags = Some(parse_on_off_xml_element(child_node)?)
                    }
                    "decimalSymbol" => instance.decimal_symbol = Some(child_node.get_val_attribute()?.clone()),
                    "listSeparator" => instance.list_separator = Some(child_node.get_val_attribute()?.clone()),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shared::sharedtypes::Percentage;
    use std::str::FromStr;

    impl Password {
        pub const TEST_ATTRIBUTES: &'static str =
            r#"w:algorithmName="MD5" w:hashValue="Some hash" w:saltValue="Some salt" w:spinCount="1""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}></{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                algorithm_name: Some(String::from("MD5")),
                hash_value: Some(Base64Binary::from("Some hash")),
                salt_value: Some(Base64Binary::from("Some salt")),
                spin_count: Some(1),
            }
        }
    }

    #[test]
    pub fn test_password_from_xml() {
        let xml = Password::test_xml("password");
        assert_eq!(
            Password::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Password::test_instance()
        );
    }

    impl WriteProtection {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:recommended="true" {}></{node_name}>"#,
                Password::TEST_ATTRIBUTES,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                recommended: Some(true),
                password: Password::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_write_propection_from_xml() {
        let xml = WriteProtection::test_xml("writeProtection");
        assert_eq!(
            WriteProtection::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WriteProtection::test_instance()
        );
    }

    impl Zoom {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="fullPage" w:percent="100%"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(ZoomType::FullPage),
                percent: DecimalNumberOrPercent::Percentage(Percentage(100.0)),
            }
        }
    }

    #[test]
    pub fn test_zoom_from_xml() {
        let xml = Zoom::test_xml("zoom");
        assert_eq!(
            Zoom::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Zoom::test_instance()
        );
    }

    impl WritingStyle {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:lang="en-US" w:vendorID="64" w:dllVersion="131078" w:nlCheck="true" w:checkStyle="true" w:appName="testApp">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                language: Lang::from("en-US"),
                vendor_id: String::from("64"),
                dll_version: String::from("131078"),
                natural_language_check: Some(true),
                check_style: true,
                app_name: String::from("testApp"),
            }
        }
    }

    #[test]
    pub fn test_writing_style_from_xml() {
        let xml = WritingStyle::test_xml("writingStyle");
        assert_eq!(
            WritingStyle::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WritingStyle::test_instance()
        );
    }

    impl StylePaneFilter {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:allStyles="true" w:customStyles="true" w:latentStyles="true" w:stylesInUse="true"
                w:headingStyles="true" w:numberingStyles="true" w:tableStyles="true" w:directFormattingOnRuns="true"
                w:directFormattingOnParagraphs="true" w:directFormattingOnNumbering="true" w:directFormattingOnTables="true"
                w:clearFormatting="true" w:top3HeadingStyles="true" w:visibleStyles="true" w:alternateStyleNames="true">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                all_styles: Some(true),
                custom_styles: Some(true),
                latent_styles: Some(true),
                styles_in_use: Some(true),
                heading_styles: Some(true),
                numbering_styles: Some(true),
                table_styles: Some(true),
                direct_formatting_on_runs: Some(true),
                direct_formatting_on_paragraphs: Some(true),
                direct_formatting_on_numbering: Some(true),
                direct_formatting_on_tables: Some(true),
                clear_formatting: Some(true),
                top_three_heading_styles: Some(true),
                visible_styles: Some(true),
                alternate_style_names: Some(true),
            }
        }
    }

    #[test]
    pub fn test_style_pane_filter_from_xml() {
        let xml = StylePaneFilter::test_xml("stylePaneFilter");
        assert_eq!(
            StylePaneFilter::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            StylePaneFilter::test_instance()
        );
    }

    impl Proof {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:spelling="clean" w:grammar="dirty"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                spelling: Some(ProofType::Clean),
                grammar: Some(ProofType::Dirty),
            }
        }
    }

    #[test]
    pub fn test_proof_from_xml() {
        let xml = Proof::test_xml("proof");
        assert_eq!(
            Proof::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Proof::test_instance()
        );
    }

    impl OdsoFieldMapData {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <type w:val="dbColumn" />
                <name w:val="first" />
                <mappedName w:val="First Name" />
                <column w:val="0" />
                <lid w:val="en-US" />
                <dynamicAddress />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                field_type: Some(MailMergeOdsoFMDFieldType::DatabaseColumn),
                name: Some(String::from("first")),
                mapped_name: Some(String::from("First Name")),
                column: Some(0),
                language_id: Some(Lang::from("en-US")),
                dynamic_address: Some(true),
            }
        }
    }

    #[test]
    pub fn test_odso_field_map_data_from_xml() {
        let xml = OdsoFieldMapData::test_xml("odsoFieldMapData");
        assert_eq!(
            OdsoFieldMapData::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            OdsoFieldMapData::test_instance()
        );
    }

    impl Odso {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <udl w:val="Provider=Example" />
                <table w:val="Some table" />
                <src r:id="rId1" />
                <colDelim w:val="44" />
                <type w:val="database" />
                <fHdr />
                {}
                <recipientData r:id="rId2" />
            </{node_name}>"#,
                OdsoFieldMapData::test_xml("fieldMapData"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                udl: Some(String::from("Provider=Example")),
                table: Some(String::from("Some table")),
                source: Some(Rel {
                    rel_id: String::from("rId1"),
                }),
                column_delimiter: Some(44),
                mail_merge_source_type: Some(MailMergeSourceType::Database),
                first_header: Some(true),
                field_map_datas: vec![OdsoFieldMapData::test_instance()],
                recipient_datas: vec![Rel {
                    rel_id: String::from("rId2"),
                }],
            }
        }
    }

    #[test]
    pub fn test_odso_from_xml() {
        let xml = Odso::test_xml("odso");
        assert_eq!(
            Odso::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Odso::test_instance()
        );
    }

    impl MailMerge {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <mainDocumentType w:val="catalog" />
                <linkToQuery />
                <dataType w:val="database" />
                <connectString w:val="Provider=Example" />
                <query w:val="SELECT * FROM Documentation" />
                <dataSource r:id="rId1" />
                <headerSource r:id="rId2" />
                <doNotSuppressBlankLines />
                <destination w:val="newDocument" />
                <addressFieldName w:val="Alternate Email Address" />
                <mailSubject w:val="John Smith" />
                <mailAsAttachment />
                <viewMergedData />
                <activeRecord w:val="2" />
                <checkErrors w:val="1" />
                {}
            </{node_name}>"#,
                Odso::test_xml("odso"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                main_document_type: MailMergeDocType::Catalog,
                link_to_query: Some(true),
                data_type: MailMergeDataType::from("database"),
                connect_string: Some(String::from("Provider=Example")),
                query: Some(String::from("SELECT * FROM Documentation")),
                data_source: Some(Rel {
                    rel_id: String::from("rId1"),
                }),
                header_source: Some(Rel {
                    rel_id: String::from("rId2"),
                }),
                do_not_suppress_blank_lines: Some(true),
                destination: Some(MailMergeDest::NewDocument),
                address_field_name: Some(String::from("Alternate Email Address")),
                mail_subject: Some(String::from("John Smith")),
                mail_as_attachment: Some(true),
                view_merged_data: Some(true),
                active_record: Some(2),
                check_errors: Some(1),
                odso: Some(Odso::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_mail_merge_from_xml() {
        let xml = MailMerge::test_xml("mailMerge");
        assert_eq!(
            MailMerge::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            MailMerge::test_instance()
        );
    }

    impl TrackChangesView {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:markup="true" w:comments="true" w:insDel="true" w:formatting="true" w:inkAnnotations="true">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                markup: Some(true),
                comments: Some(true),
                display_content_revisions: Some(true),
                formatting: Some(true),
                ink_annotations: Some(true),
            }
        }
    }

    #[test]
    pub fn test_track_changes_view_from_xml() {
        let xml = TrackChangesView::test_xml("trackChangesView");
        assert_eq!(
            TrackChangesView::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TrackChangesView::test_instance()
        );
    }

    impl DocProtect {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:edit="none" w:formatting="true" w:enforcement="true" {}></{node_name}>"#,
                Password::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                edit: Some(DocProtectType::None),
                formatting: Some(true),
                enforcement: Some(true),
                password: Password::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_doc_protect_from_xml() {
        let xml = DocProtect::test_xml("docProtect");
        assert_eq!(
            DocProtect::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocProtect::test_instance()
        );
    }

    impl Kinsoku {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:lang="ja-JP" w:val="$"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                language: Lang::from("ja-JP"),
                value: String::from("$"),
            }
        }
    }

    #[test]
    pub fn test_kinsoku_from_xml() {
        let xml = Kinsoku::test_xml("kinsoku");
        assert_eq!(
            Kinsoku::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Kinsoku::test_instance()
        );
    }

    impl SaveThroughXslt {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} r:id="rId1" w:solutionID="Some solution"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                rel_id: Some(RelationshipId::from("rId1")),
                solution_id: Some(String::from("Some solution")),
            }
        }
    }

    #[test]
    pub fn test_save_through_xslt_from_xml() {
        let xml = SaveThroughXslt::test_xml("saveThroughXslt");
        assert_eq!(
            SaveThroughXslt::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()),
            SaveThroughXslt::test_instance()
        );
    }

    impl FtnEndSepRef {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:id="0"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { id: 0 }
        }
    }

    #[test]
    pub fn test_ftn_edn_sep_ref_from_xml() {
        let xml = FtnEndSepRef::test_xml("ftnEndSepRef");
        assert_eq!(
            FtnEndSepRef::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FtnEndSepRef::test_instance()
        );
    }

    impl FtnDocProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                FtnProps::test_extension_xml(),
                FtnEndSepRef::test_xml("footnote"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: FtnProps::test_instance(),
                footnotes: vec![FtnEndSepRef::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_ftn_doc_props_from_xml() {
        let xml = FtnDocProps::test_xml("ftnDocProps");
        assert_eq!(
            FtnDocProps::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            FtnDocProps::test_instance()
        );
    }

    impl EdnDocProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                EdnProps::test_extension_xml(),
                FtnEndSepRef::test_xml("endnote"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: EdnProps::test_instance(),
                endnotes: vec![FtnEndSepRef::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_edn_doc_props_from_xml() {
        let xml = EdnDocProps::test_xml("ednDocProps");
        assert_eq!(
            EdnDocProps::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            EdnDocProps::test_instance()
        );
    }

    impl CompatSetting {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:name="cooper" w:uri="http://www.example.com/exampleSetting" w:val="1">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: Some(String::from("cooper")),
                uri: Some(String::from("http://www.example.com/exampleSetting")),
                value: Some(String::from("1")),
            }
        }
    }

    #[test]
    pub fn test_compat_setting_from_xml() {
        let xml = CompatSetting::test_xml("compatSetting");
        assert_eq!(
            CompatSetting::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()),
            CompatSetting::test_instance()
        );
    }

    impl Compat {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <spaceForUL />
                <balanceSingleByteDoubleByteWidth />
                <doNotLeaveBackslashAlone />
                <ulTrailSpace />
                <doNotExpandShiftReturn />
                <adjustLineHeightInTable />
                <applyBreakingRules />
                {}
            </{node_name}>"#,
                CompatSetting::test_xml("compatSetting"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                space_for_underline: Some(true),
                balance_single_byte_double_byte_width: Some(true),
                do_not_leave_backslash_alone: Some(true),
                underline_trail_space: Some(true),
                do_not_expand_shift_return: Some(true),
                adjust_line_height_in_table: Some(true),
                apply_breaking_rules: Some(true),
                compatibility_settings: vec![CompatSetting::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_compat_from_xml() {
        let xml = Compat::test_xml("compat");
        assert_eq!(
            Compat::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Compat::test_instance()
        );
    }

    impl DocVar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:name="Example name" w:val="Example value"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: String::from("Example name"),
                value: String::from("Example value"),
            }
        }
    }

    #[test]
    pub fn test_doc_var_from_xml() {
        let xml = DocVar::test_xml("docVar");
        assert_eq!(
            DocVar::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocVar::test_instance()
        );
    }

    impl DocVars {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{doc_var}{doc_var}</{node_name}>"#,
                doc_var = DocVar::test_xml("docVar"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![DocVar::test_instance(), DocVar::test_instance()])
        }
    }

    #[test]
    pub fn test_doc_vars_from_xml() {
        let xml = DocVars::test_xml("docVars");
        assert_eq!(
            DocVars::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocVars::test_instance()
        );
    }

    impl DocRsids {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <rsidRoot w:val="ffffffff" />
                <rsid w:val="ffffffff" />
                <rsid w:val="ffffffff" />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                revision_id_root: Some(0xffffffff),
                revision_ids: vec![0xffffffff, 0xffffffff],
            }
        }
    }

    #[test]
    pub fn test_doc_rsids_from_xml() {
        let xml = DocRsids::test_xml("docRsids");
        assert_eq!(
            DocRsids::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            DocRsids::test_instance()
        );
    }

    impl ColorSchemeMapping {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1"
                w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6"
                w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                background1: WmlColorSchemeIndex::Light1,
                text1: WmlColorSchemeIndex::Dark1,
                background2: WmlColorSchemeIndex::Light2,
                text2: WmlColorSchemeIndex::Dark2,
                accent1: WmlColorSchemeIndex::Accent1,
                accent2: WmlColorSchemeIndex::Accent2,
                accent3: WmlColorSchemeIndex::Accent3,
                accent4: WmlColorSchemeIndex::Accent4,
                accent5: WmlColorSchemeIndex::Accent5,
                accent6: WmlColorSchemeIndex::Accent6,
                hyperlink: WmlColorSchemeIndex::Hyperlink,
                followed_hyperlink: WmlColorSchemeIndex::FollowedHyperlink,
            }
        }
    }

    #[test]
    pub fn test_color_scheme_mapping_from_xml() {
        let xml = ColorSchemeMapping::test_xml("colorSchemeMapping");
        assert_eq!(
            ColorSchemeMapping::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ColorSchemeMapping::test_instance()
        );
    }

    impl Caption {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:name="Example name" w:pos="below" w:chapNum="true" w:heading="0" w:noLabel="true"
                w:numFmt="decimal" w:sep="hyphen">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: String::from("Example name"),
                position: Some(CaptionPos::Below),
                chapter_number: Some(true),
                heading: Some(0),
                no_label: Some(true),
                numbering_format: Some(NumberFormat::Decimal),
                separator: Some(ChapterSep::Hyphen),
            }
        }
    }

    #[test]
    pub fn test_caption_from_xml() {
        let xml = Caption::test_xml("caption");
        assert_eq!(
            Caption::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Caption::test_instance()
        );
    }

    impl AutoCaption {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:name="Example name" w:caption="Example caption"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: String::from("Example name"),
                caption: String::from("Example caption"),
            }
        }
    }

    #[test]
    pub fn test_auto_caption_from_xml() {
        let xml = AutoCaption::test_xml("autoCaption");
        assert_eq!(
            AutoCaption::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            AutoCaption::test_instance()
        );
    }

    impl AutoCaptions {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{auto_caption}{auto_caption}</{node_name}>"#,
                auto_caption = AutoCaption::test_xml("autoCaption"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![AutoCaption::test_instance(), AutoCaption::test_instance()])
        }
    }

    #[test]
    pub fn test_auto_captions_from_xml() {
        let xml = AutoCaptions::test_xml("autoCaptions");
        assert_eq!(
            AutoCaptions::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            AutoCaptions::test_instance()
        );
    }

    impl Captions {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                Caption::test_xml("caption"),
                AutoCaptions::test_xml("autoCaptions"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                captions: vec![Caption::test_instance()],
                auto_captions: Some(AutoCaptions::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_captions_from_xml() {
        let xml = Captions::test_xml("captions");
        assert_eq!(
            Captions::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Captions::test_instance()
        );
    }

    impl ReadingModeInkLockDown {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:actualPg="true" w:w="100" w:h="100" w:fontSz="100%"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                use_actual_pages: true,
                width: 100,
                height: 100,
                font_size: DecimalNumberOrPercent::Percentage(Percentage(100.0)),
            }
        }
    }

    #[test]
    pub fn test_reading_mode_ink_lock_down_from_xml() {
        let xml = ReadingModeInkLockDown::test_xml("readingModeInkLockDown");
        assert_eq!(
            ReadingModeInkLockDown::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            ReadingModeInkLockDown::test_instance()
        );
    }

    impl SmartTagType {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:namespaceuri="urn:smartTagExample" w:name="Example name"
                w:url="http://www.example.com/smartTag">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                namespaceuri: String::from("urn:smartTagExample"),
                name: String::from("Example name"),
                url: String::from("http://www.example.com/smartTag"),
            }
        }
    }

    #[test]
    pub fn test_smart_tag_type_from_xml() {
        let xml = SmartTagType::test_xml("smartTagType");
        assert_eq!(
            SmartTagType::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SmartTagType::test_instance()
        );
    }

    impl Settings {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                <view w:val="none" />
                {}
                <removePersonalInformation />
                <removeDateAndTime />
                <doNotDisplayPageBoundaries />
                <displayBackgroundShape />
                <printPostScriptOverText />
                <printFractionalCharacterWidth />
                <printFormsData />
                <embedTrueTypeFonts />
                <embedSystemFonts />
                <saveSubsetFonts />
                <saveFormsData />
                <mirrorMargins />
                <alignBordersAndEdges />
                <bordersDoNotSurroundHeader />
                <bordersDoNotSurroundFooter />
                <gutterAtTop />
                <hideSpellingErrors />
                <hideGrammaticalErrors />
                {}
                {}
                <formsDesign />
                <attachedTemplate r:id="rId1" />
                <linkStyles />
                {}
                <stylePaneSortMethod w:val="default" />
                <documentType w:val="Example type" />
                {}
                {}
                <trackRevisions />
                <doNotTrackMoves />
                <doNotTrackFormatting />
                {}
                <autoFormatOverride />
                <styleLockTheme />
                <styleLockQFSet />
                <defaultTabStop w:val="100" />
                <autoHyphenation />
                <consecutiveHyphenLimit w:val="1" />
                <hyphenationZone w:val="10" />
                <doNotHyphenateCaps />
                <showEnvelope />
                <summaryLength w:val="100%" />
                <clickAndTypeStyle w:val="Some style" />
                <defaultTableStyle w:val="Some table style" />
                <evenAndOddHeaders />
                <bookFoldRevPrinting />
                <bookFoldPrinting />
                <bookFoldPrintingSheets />
                <drawingGridHorizontalSpacing w:val="100" />
                <drawingGridVerticalSpacing w:val="100" />
                <displayHorizontalDrawingGridEvery w:val="2" />
                <displayVerticalDrawingGridEvery w:val="2" />
                <doNotUseMarginsForDrawingGridOrigin />
                <drawingGridHorizontalOrigin w:val="10" />
                <drawingGridVerticalOrigin w:val="10" />
                <doNotShadeFormData />
                <noPunctuationKerning />
                <characterSpacingControl w:val="doNotCompress" />
                <printTwoOnOne />
                <strictFirstAndLastChars />
                {}
                {}
                <savePreviewPicture />
                <doNotValidateAgainstSchema />
                <saveInvalidXml />
                <ignoreMixedContent />
                <alwaysShowPlaceholderText />
                <doNotDemarcateInvalidXml />
                <saveXmlDataOnly />
                <useXSLTWhenSaving />
                {}
                <showXMLTags />
                <alwaysMergeEmptyNamespace />
                <updateFields />
                {}
                {}
                {}
                {}
                {}
                <attachedSchema w:val="Some schema" />
                {}
                {}
                <doNotIncludeSubdocsInStats />
                <doNotAutoCompressPictures />
                <forceUpgrade />
                {}
                {}
                {}
                <doNotEmbedSmartTags />
                <decimalSymbol w:val="." />
                <listSeparator w:val="," />
            </{node_name}>"#,
                WriteProtection::test_xml("writeProtection"),
                Zoom::test_xml("zoom"),
                WritingStyle::test_xml("activeWritingStyle"),
                Proof::test_xml("proofState"),
                StylePaneFilter::test_xml("stylePaneFormatFilter"),
                MailMerge::test_xml("mailMerge"),
                TrackChangesView::test_xml("revisionView"),
                DocProtect::test_xml("documentProtection"),
                Kinsoku::test_xml("noLineBreaksAfter"),
                Kinsoku::test_xml("noLineBreaksBefore"),
                SaveThroughXslt::test_xml("saveThroughXslt"),
                FtnDocProps::test_xml("footnotePr"),
                EdnDocProps::test_xml("endnotePr"),
                Compat::test_xml("compat"),
                DocVars::test_xml("docVars"),
                DocRsids::test_xml("rsids"),
                Language::test_xml("themeFontLang"),
                ColorSchemeMapping::test_xml("clrSchemeMapping"),
                Captions::test_xml("captions"),
                ReadingModeInkLockDown::test_xml("readModeInkLockDown"),
                SmartTagType::test_xml("smartTagType"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                write_protection: Some(WriteProtection::test_instance()),
                view: Some(View::None),
                zoom: Some(Zoom::test_instance()),
                remove_personal_information: Some(true),
                remove_date_and_time: Some(true),
                do_not_display_page_boundaries: Some(true),
                display_background_shape: Some(true),
                print_post_script_over_text: Some(true),
                print_fractional_character_width: Some(true),
                print_forms_data: Some(true),
                embed_true_type_fonts: Some(true),
                embed_system_fonts: Some(true),
                save_subset_fonts: Some(true),
                save_forms_data: Some(true),
                mirror_margins: Some(true),
                align_borders_and_edges: Some(true),
                borders_do_not_surround_header: Some(true),
                borders_do_not_surround_footer: Some(true),
                gutter_at_top: Some(true),
                hide_spelling_errors: Some(true),
                hide_grammatical_errors: Some(true),
                active_writing_styles: vec![WritingStyle::test_instance()],
                proof_state: Some(Proof::test_instance()),
                forms_design: Some(true),
                attached_template: Some(Rel {
                    rel_id: RelationshipId::from("rId1"),
                }),
                link_styles: Some(true),
                style_pane_format_filter: Some(StylePaneFilter::test_instance()),
                style_pane_sort_method: Some(StyleSort::Default),
                document_type: Some(DocType::from("Example type")),
                mail_merge: Some(MailMerge::test_instance()),
                revision_view: Some(TrackChangesView::test_instance()),
                track_revisions: Some(true),
                do_not_track_moves: Some(true),
                do_not_track_formatting: Some(true),
                document_protection: Some(DocProtect::test_instance()),
                auto_format_override: Some(true),
                style_lock_theme: Some(true),
                style_lock_set: Some(true),
                default_tab_stop: Some(TwipsMeasure::Decimal(100)),
                auto_hyphenation: Some(true),
                consecutive_hyphen_limit: Some(1),
                hyphenation_zone: Some(TwipsMeasure::Decimal(10)),
                do_not_hyphenate_capitals: Some(true),
                show_envelope: Some(true),
                summary_length: Some(DecimalNumberOrPercent::Percentage(Percentage(100.0))),
                click_and_type_style: Some(String::from("Some style")),
                default_table_style: Some(String::from("Some table style")),
                even_and_odd_headers: Some(true),
                book_fold_revision_printing: Some(true),
                book_fold_printing: Some(true),
                book_fold_printing_sheets: Some(true),
                drawing_grid_horizontal_spacing: Some(TwipsMeasure::Decimal(100)),
                drawing_grid_vertical_spacing: Some(TwipsMeasure::Decimal(100)),
                display_horizontal_drawing_grid_every: Some(2),
                display_vertical_drawing_grid_every: Some(2),
                do_not_use_margins_for_drawing_grid_origin: Some(true),
                drawing_grid_horizontal_origin: Some(TwipsMeasure::Decimal(10)),
                drawing_grid_vertical_origin: Some(TwipsMeasure::Decimal(10)),
                do_not_shade_form_data: Some(true),
                no_punctuation_kerning: Some(true),
                character_spacing_control: Some(CharacterSpacing::DoNotCompress),
                print_two_on_one: Some(true),
                strict_first_and_last_chars: Some(true),
                no_line_breaks_after: Some(Kinsoku::test_instance()),
                no_line_breaks_before: Some(Kinsoku::test_instance()),
                save_preview_picture: Some(true),
                do_not_validate_against_schema: Some(true),
                save_invalid_xml: Some(true),
                ignore_mixed_content: Some(true),
                always_show_placeholder_text: Some(true),
                do_not_demarcate_invalid_xml: Some(true),
                save_xml_data_only: Some(true),
                use_xslt_when_saving: Some(true),
                save_through_xslt: Some(SaveThroughXslt::test_instance()),
                show_xml_tags: Some(true),
                always_merge_empty_namespace: Some(true),
                update_fields: Some(true),
                footnote_properties: Some(FtnDocProps::test_instance()),
                endnote_properties: Some(EdnDocProps::test_instance()),
                compatibility: Some(Compat::test_instance()),
                document_variables: Some(DocVars::test_instance()),
                revision_ids: Some(DocRsids::test_instance()),
                attached_schemas: vec![String::from("Some schema")],
                theme_font_lang: Some(Language::test_instance()),
                color_scheme_mapping: Some(ColorSchemeMapping::test_instance()),
                do_not_include_subdocs_in_stats: Some(true),
                do_not_auto_compress_pictures: Some(true),
                force_upgrade: true,
                captions: Some(Captions::test_instance()),
                read_move_ink_lock_down: Some(ReadingModeInkLockDown::test_instance()),
                smart_tag_types: vec![SmartTagType::test_instance()],
                do_not_embed_smart_tags: Some(true),
                decimal_symbol: Some(String::from(".")),
                list_separator: Some(String::from(",")),
            }
        }
    }

    #[test]
    pub fn test_settings_from_xml() {
        let xml = Settings::test_xml("settings");
        assert_eq!(
            Settings::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Settings::test_instance()
        );
    }
}
