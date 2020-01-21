use super::{
    document::{
        BlockLevelElts, Border, Cnf, CustomXmlPr, HAnchor, HeightRule, Markup, MeasurementOrPercent,
        RangeMarkupElements, RunLevelElts, SdtEndPr, SdtPr, Shd, SignedTwipsMeasure, TextDirection, TrackChange,
        VAnchor, VerticalJc,
    },
    simpletypes::{parse_on_off_xml_element, DecimalNumber, LongHexNumber},
    util::XmlNodeExt,
};
use log::info;
use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    shared::sharedtypes::{OnOff, TwipsMeasure, XAlign, XmlName, YAlign},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblOverlap {
    #[strum(serialize = "never")]
    Never,
    #[strum(serialize = "overlap")]
    Overlap,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblWidthType {
    #[strum(serialize = "nil")]
    NoWidth,
    #[strum(serialize = "pct")]
    Percent,
    #[strum(serialize = "dxa")]
    TwentiethsOfPoint,
    #[strum(serialize = "auto")]
    Auto,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPPr {
    pub left_from_text: Option<TwipsMeasure>,
    pub right_from_text: Option<TwipsMeasure>,
    pub top_from_text: Option<TwipsMeasure>,
    pub bottom_from_text: Option<TwipsMeasure>,
    pub vertical_anchor: Option<VAnchor>,
    pub horizontal_anchor: Option<HAnchor>,
    pub horizontal_alignment: Option<XAlign>,
    pub horizontal_distance: Option<SignedTwipsMeasure>,
    pub vertical_alignment: Option<YAlign>,
    pub vertical_distance: Option<SignedTwipsMeasure>,
}

impl TblPPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPPr");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:leftFromText" => instance.left_from_text = Some(value.parse()?),
                "w:rightFromText" => instance.right_from_text = Some(value.parse()?),
                "w:topFromText" => instance.top_from_text = Some(value.parse()?),
                "w:bottomFromText" => instance.bottom_from_text = Some(value.parse()?),
                "w:vertAnchor" => instance.vertical_anchor = Some(value.parse()?),
                "w:horzAnchor" => instance.horizontal_anchor = Some(value.parse()?),
                "w:tblpXSpec" => instance.horizontal_alignment = Some(value.parse()?),
                "w:tblpX" => instance.horizontal_distance = Some(value.parse()?),
                "w:tblpYSpec" => instance.vertical_alignment = Some(value.parse()?),
                "w:tblpY" => instance.vertical_distance = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblWidth {
    pub width: Option<MeasurementOrPercent>,
    pub width_type: Option<TblWidthType>,
}

impl TblWidth {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblWidth");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:w" => instance.width = Some(value.parse()?),
                "w:type" => instance.width_type = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum JcTable {
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "start")]
    Start,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblBorders {
    pub top: Option<Border>,
    pub start: Option<Border>,
    pub bottom: Option<Border>,
    pub end: Option<Border>,
    pub inside_horizontal: Option<Border>,
    pub inside_vertical: Option<Border>,
}

impl TblBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblBorders");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "start" => instance.start = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "end" => instance.end = Some(Border::from_xml_element(child_node)?),
                "insideH" => instance.inside_horizontal = Some(Border::from_xml_element(child_node)?),
                "insideV" => instance.inside_vertical = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblLayoutType {
    #[strum(serialize = "fixed")]
    Fixed,
    #[strum(serialize = "autofit")]
    Autofit,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblCellMar {
    pub top: Option<TblWidth>,
    pub start: Option<TblWidth>,
    pub bottom: Option<TblWidth>,
    pub end: Option<TblWidth>,
}

impl TblCellMar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblCellMar");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(TblWidth::from_xml_element(child_node)?),
                "start" => instance.start = Some(TblWidth::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(TblWidth::from_xml_element(child_node)?),
                "end" => instance.end = Some(TblWidth::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblLook {
    pub first_row: Option<OnOff>,
    pub last_row: Option<OnOff>,
    pub first_column: Option<OnOff>,
    pub last_column: Option<OnOff>,
    pub no_horizontal_band: Option<OnOff>,
    pub no_vertical_band: Option<OnOff>,
}

impl TblLook {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblLook");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:firstRow" => instance.first_row = Some(parse_xml_bool(value)?),
                "w:lastRow" => instance.last_row = Some(parse_xml_bool(value)?),
                "w:firstColumn" => instance.first_column = Some(parse_xml_bool(value)?),
                "w:lastColumn" => instance.last_column = Some(parse_xml_bool(value)?),
                "w:noHBand" => instance.no_horizontal_band = Some(parse_xml_bool(value)?),
                "w:noVBand" => instance.no_vertical_band = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrBase {
    pub style: Option<String>,
    pub paragraph_properties: Option<TblPPr>,
    pub overlap: Option<TblOverlap>,
    pub bidirectional_visual: Option<OnOff>,
    pub style_row_band_size: Option<DecimalNumber>,
    pub style_column_band_size: Option<DecimalNumber>,
    pub width: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub cell_spacing: Option<TblWidth>,
    pub indent: Option<TblWidth>,
    pub borders: Option<TblBorders>,
    pub shading: Option<Shd>,
    pub layout: Option<TblLayoutType>,
    pub cell_margin: Option<TblCellMar>,
    pub look: Option<TblLook>,
    pub caption: Option<String>,
    pub description: Option<String>,
}

impl TblPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPrBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tblStyle" => self.style = Some(xml_node.get_val_attribute()?.clone()),
            "tblpPr" => self.paragraph_properties = Some(TblPPr::from_xml_element(xml_node)?),
            "tblOverlap" => self.overlap = Some(xml_node.get_val_attribute()?.parse()?),
            "bidiVisual" => self.bidirectional_visual = Some(parse_on_off_xml_element(xml_node)?),
            "tblStyleRowBandSize" => self.style_row_band_size = Some(xml_node.get_val_attribute()?.parse()?),
            "tblStyleColBandSize" => self.style_column_band_size = Some(xml_node.get_val_attribute()?.parse()?),
            "tblW" => self.width = Some(TblWidth::from_xml_element(xml_node)?),
            "jc" => self.alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "tblCellSpacing" => self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?),
            "tblInd" => self.indent = Some(TblWidth::from_xml_element(xml_node)?),
            "tblBorders" => self.borders = Some(TblBorders::from_xml_element(xml_node)?),
            "shd" => self.shading = Some(Shd::from_xml_element(xml_node)?),
            "tblLayout" => self.layout = Some(xml_node.get_val_attribute()?.parse()?),
            "tblCellMar" => self.cell_margin = Some(TblCellMar::from_xml_element(xml_node)?),
            "tblLook" => self.look = Some(TblLook::from_xml_element(xml_node)?),
            "tblCaption" => self.caption = Some(xml_node.get_val_attribute()?.clone()),
            "tblDescription" => self.description = Some(xml_node.get_val_attribute()?.clone()),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblPrChange {
    pub base: TrackChange,
    pub properties: TblPrBase,
}

impl TblPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tblPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblPr").into())
            .and_then(TblPrBase::from_xml_element)?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPr {
    pub base: TblPrBase,
    pub change: Option<TblPrChange>,
}

impl TblPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPrChange" => instance.change = Some(TblPrChange::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGridCol {
    pub width: Option<TwipsMeasure>,
}

impl TblGridCol {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblGridCol");

        let width = xml_node.attributes.get("w:w").map(|value| value.parse()).transpose()?;

        Ok(Self { width })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblGridChange {
    pub base: Markup,
    pub grid: TblGridBase,
}

impl TblGridChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblGridChange");

        let base = Markup::from_xml_element(xml_node)?;
        let grid = TblGridBase::from_xml_element(xml_node)?;

        Ok(Self { base, grid })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGridBase {
    pub columns: Vec<TblGridCol>,
}

impl TblGridBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblGridBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        if xml_node.local_name() == "gridCol" {
            self.columns.push(TblGridCol::from_xml_element(xml_node)?);
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGrid {
    pub base: TblGridBase,
    pub change: Option<TblGridChange>,
}

impl TblGrid {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblGrid");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblGridChange" => instance.change = Some(TblGridChange::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrExBase {
    pub width: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub cell_spacing: Option<TblWidth>,
    pub indent: Option<TblWidth>,
    pub borders: Option<TblBorders>,
    pub shading: Option<Shd>,
    pub layout: Option<TblLayoutType>,
    pub cell_margin: Option<TblCellMar>,
    pub look: Option<TblLook>,
}

impl TblPrExBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPrExBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tblW" => self.width = Some(TblWidth::from_xml_element(xml_node)?),
            "jc" => self.alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "tblCellSpacing" => self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?),
            "tblInd" => self.indent = Some(TblWidth::from_xml_element(xml_node)?),
            "tblBorders" => self.borders = Some(TblBorders::from_xml_element(xml_node)?),
            "shd" => self.shading = Some(Shd::from_xml_element(xml_node)?),
            "tblLayout" => self.layout = Some(xml_node.get_val_attribute()?.parse()?),
            "tblCellMar" => self.cell_margin = Some(TblCellMar::from_xml_element(xml_node)?),
            "tblLook" => self.look = Some(TblLook::from_xml_element(xml_node)?),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblPrExChange {
    pub base: TrackChange,
    pub properties_ex: TblPrExBase,
}

impl TblPrExChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPrExChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let properties_ex = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tblPrEx")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblPrEx").into())
            .and_then(TblPrExBase::from_xml_element)?;

        Ok(Self { base, properties_ex })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrEx {
    pub base: TblPrExBase,
    pub change: Option<TblPrExChange>,
}

impl TblPrEx {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TblPrEx");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPrExChange" => instance.change = Some(TblPrExChange::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TrPrBase {
    pub conditional_formatting: Option<Cnf>,
    pub div_id: Option<DecimalNumber>,
    pub grid_column_before_first_cell: Option<DecimalNumber>,
    pub grid_column_after_last_cell: Option<DecimalNumber>,
    pub width_before_row: Option<TblWidth>,
    pub width_after_row: Option<TblWidth>,
    pub cant_split: Option<OnOff>,
    pub row_height: Option<Height>,
    pub header: Option<OnOff>,
    pub cell_spacing: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub hidden: Option<OnOff>,
}

impl TrPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TrPrBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "cnfStyle" => self.conditional_formatting = Some(Cnf::from_xml_element(xml_node)?),
            "divId" => self.div_id = Some(xml_node.get_val_attribute()?.parse()?),
            "gridBefore" => self.grid_column_before_first_cell = Some(xml_node.get_val_attribute()?.parse()?),
            "gridAfter" => self.grid_column_after_last_cell = Some(xml_node.get_val_attribute()?.parse()?),
            "wBefore" => self.width_before_row = Some(TblWidth::from_xml_element(xml_node)?),
            "wAfter" => self.width_after_row = Some(TblWidth::from_xml_element(xml_node)?),
            "cantSplit" => self.cant_split = Some(parse_on_off_xml_element(xml_node)?),
            "trHeight" => self.row_height = Some(Height::from_xml_element(xml_node)?),
            "tblHeader" => self.header = Some(parse_on_off_xml_element(xml_node)?),
            "tblCellSpacing" => self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?),
            "jc" => self.alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "hidden" => self.hidden = Some(parse_on_off_xml_element(xml_node)?),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TrPrChange {
    pub base: TrackChange,
    pub properties: TrPrBase,
}

impl TrPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TrPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "trPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "trPr").into())
            .and_then(TrPrBase::from_xml_element)?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TrPr {
    pub base: TrPrBase,
    pub inserted: Option<TrackChange>,
    pub deleted: Option<TrackChange>,
    pub change: Option<TrPrChange>,
}

impl TrPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TrPr");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "ins" => instance.inserted = Some(TrackChange::from_xml_element(child_node)?),
                "del" => instance.deleted = Some(TrackChange::from_xml_element(child_node)?),
                "trPrChange" => instance.change = Some(TrPrChange::from_xml_element(child_node)?),
                _ => instance.base = instance.base.try_update_from_xml_element(child_node)?,
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Merge {
    #[strum(serialize = "continue")]
    Continue,
    #[strum(serialize = "restart")]
    Restart,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcBorders {
    pub top: Option<Border>,
    pub start: Option<Border>,
    pub bottom: Option<Border>,
    pub end: Option<Border>,
    pub inside_horizontal: Option<Border>,
    pub inside_vertical: Option<Border>,
    pub top_left_to_bottom_right: Option<Border>,
    pub top_right_to_bottom_left: Option<Border>,
}

impl TcBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcBorders");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "start" => instance.start = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "end" => instance.end = Some(Border::from_xml_element(child_node)?),
                "insideH" => instance.inside_horizontal = Some(Border::from_xml_element(child_node)?),
                "insideV" => instance.inside_vertical = Some(Border::from_xml_element(child_node)?),
                "tl2br" => instance.top_left_to_bottom_right = Some(Border::from_xml_element(child_node)?),
                "tr2bl" => instance.top_right_to_bottom_left = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcMar {
    pub top: Option<TblWidth>,
    pub start: Option<TblWidth>,
    pub bottom: Option<TblWidth>,
    pub end: Option<TblWidth>,
}

impl TcMar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcMar");

        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(TblWidth::from_xml_element(child_node)?),
                "start" => instance.start = Some(TblWidth::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(TblWidth::from_xml_element(child_node)?),
                "end" => instance.end = Some(TblWidth::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Headers(pub Vec<String>);

impl Headers {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Headers");

        let headers = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "header")
            .map(|child_node| Ok(child_node.get_val_attribute()?.clone()))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(headers))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPrBase {
    pub conditional_formatting: Option<Cnf>,
    pub width: Option<TblWidth>,
    pub grid_span: Option<DecimalNumber>,
    pub vertical_merge: Option<Merge>,
    pub borders: Option<TcBorders>,
    pub shading: Option<Shd>,
    pub no_wrapping: Option<OnOff>,
    pub margin: Option<TcMar>,
    pub text_direction: Option<TextDirection>,
    pub fit_text: Option<OnOff>,
    pub vertical_alignment: Option<VerticalJc>,
    pub hide_marker: Option<OnOff>,
    pub headers: Option<Headers>,
}

impl TcPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcPrBase");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "cnfStyle" => self.conditional_formatting = Some(Cnf::from_xml_element(xml_node)?),
            "tcW" => self.width = Some(TblWidth::from_xml_element(xml_node)?),
            "gridSpan" => self.grid_span = Some(xml_node.get_val_attribute()?.parse()?),
            "vMerge" => {
                self.vertical_merge = Some(
                    xml_node
                        .attributes
                        .get("w:val")
                        .map(|value| value.parse())
                        .transpose()?
                        .unwrap_or(Merge::Continue),
                )
            }
            "tcBorders" => self.borders = Some(TcBorders::from_xml_element(xml_node)?),
            "shd" => self.shading = Some(Shd::from_xml_element(xml_node)?),
            "noWrap" => self.no_wrapping = Some(parse_on_off_xml_element(xml_node)?),
            "tcMar" => self.margin = Some(TcMar::from_xml_element(xml_node)?),
            "textDirection" => self.text_direction = Some(xml_node.get_val_attribute()?.parse()?),
            "tcFitText" => self.fit_text = Some(parse_on_off_xml_element(xml_node)?),
            "vAlign" => self.vertical_alignment = Some(xml_node.get_val_attribute()?.parse()?),
            "hideMark" => self.hide_marker = Some(parse_on_off_xml_element(xml_node)?),
            "headers" => self.headers = Some(Headers::from_xml_element(xml_node)?),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum AnnotationVMerge {
    #[strum(serialize = "cont")]
    Merge,
    #[strum(serialize = "rest")]
    Split,
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellMergeTrackChange {
    pub base: TrackChange,
    pub vertical_merge: Option<AnnotationVMerge>,
    pub vertical_merge_original: Option<AnnotationVMerge>,
}

impl CellMergeTrackChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CellMergeTrackChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let vertical_merge = xml_node
            .attributes
            .get("w:vMerge")
            .map(|value| value.parse())
            .transpose()?;

        let vertical_merge_original = xml_node
            .attributes
            .get("w:vMergeOrig")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self {
            base,
            vertical_merge,
            vertical_merge_original,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum CellMarkupElements {
    Insertion(TrackChange),
    Deletion(TrackChange),
    Merge(CellMergeTrackChange),
}

impl XsdType for CellMarkupElements {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "cellIns" => Ok(CellMarkupElements::Insertion(TrackChange::from_xml_element(xml_node)?)),
            "cellDel" => Ok(CellMarkupElements::Deletion(TrackChange::from_xml_element(xml_node)?)),
            "cellMerge" => Ok(CellMarkupElements::Merge(CellMergeTrackChange::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CellMarkupElements",
            ))),
        }
    }
}

impl XsdChoice for CellMarkupElements {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "cellIns" | "cellDel" | "cellMerge" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPrInner {
    pub base: TcPrBase,
    pub markup_element: Option<CellMarkupElements>,
}

impl TcPrInner {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcPrInner");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_element)
    }

    pub fn try_update_from_xml_element(mut self, xml_node: &XmlNode) -> Result<Self> {
        if CellMarkupElements::is_choice_member(xml_node.local_name()) {
            self.markup_element = Some(CellMarkupElements::from_xml_element(xml_node)?);
        } else {
            self.base = self.base.try_update_from_xml_element(xml_node)?;
        }

        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TcPrChange {
    pub base: TrackChange,
    pub properties: TcPrInner,
}

impl TcPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcPrChange");

        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tcPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tcPr").into())
            .and_then(TcPrInner::from_xml_element)?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPr {
    pub base: TcPrInner,
    pub change: Option<TcPrChange>,
}

impl TcPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing TcPr");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                if child_node.local_name() == "tcPrChange" {
                    instance.change = Some(TcPrChange::from_xml_element(child_node)?);
                } else {
                    instance.base = instance.base.try_update_from_xml_element(child_node)?;
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Tc {
    pub properties: Option<TcPr>,
    pub block_level_elements: Vec<BlockLevelElts>,
    pub id: Option<String>,
}

impl Tc {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Tc");

        let mut instance: Self = Default::default();

        instance.id = xml_node.attributes.get("w:id").cloned();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tcPr" => instance.properties = Some(TcPr::from_xml_element(child_node)?),
                node_name if BlockLevelElts::is_choice_member(node_name) => {
                    instance
                        .block_level_elements
                        .push(BlockLevelElts::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        if instance.block_level_elements.is_empty() {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "BlockLevelElts",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        } else {
            Ok(instance)
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlCell {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub contents: Vec<ContentCellContent>,
    pub uri: Option<String>,
    pub element: XmlName,
}

impl CustomXmlCell {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CustomXmlCell");

        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(value.clone()),
                "w:element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        let mut custom_xml_properties = None;
        let mut contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name if ContentCellContent::is_choice_member(node_name) => {
                    contents.push(ContentCellContent::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        Ok(Self {
            custom_xml_properties,
            contents,
            uri,
            element,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentCell {
    pub contents: Vec<ContentCellContent>,
}

impl SdtContentCell {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtContentCell");

        let contents = xml_node
            .child_nodes
            .iter()
            .filter_map(ContentCellContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtCell {
    pub properties: Option<SdtPr>,
    pub end_properties: Option<SdtEndPr>,
    pub content: Option<SdtContentCell>,
}

impl SdtCell {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtCell");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "sdtPr" => instance.properties = Some(SdtPr::from_xml_element(child_node)?),
                    "sdtEndPr" => instance.end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                    "sdtContent" => instance.content = Some(SdtContentCell::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ContentCellContent {
    Cell(Box<Tc>),
    CustomXml(CustomXmlCell),
    Sdt(Box<SdtCell>),
    RunLevelElement(RunLevelElts),
}

impl XsdType for ContentCellContent {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tc" => Ok(ContentCellContent::Cell(Box::new(Tc::from_xml_element(xml_node)?))),
            "customXml" => Ok(ContentCellContent::CustomXml(CustomXmlCell::from_xml_element(
                xml_node,
            )?)),
            "sdt" => Ok(ContentCellContent::Sdt(Box::new(SdtCell::from_xml_element(xml_node)?))),
            node_name if RunLevelElts::is_choice_member(node_name) => Ok(ContentCellContent::RunLevelElement(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentCellContent",
            ))),
        }
    }
}

impl XsdChoice for ContentCellContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "tc" | "customXml" | "sdt" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Row {
    pub property_exceptions: Option<TblPrEx>,
    pub properties: Option<TrPr>,
    pub contents: Vec<ContentCellContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub row_revision_id: Option<LongHexNumber>,
}

impl Row {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Row");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "w:rsidTr" => instance.row_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPrEx" => instance.property_exceptions = Some(TblPrEx::from_xml_element(child_node)?),
                "trPr" => instance.properties = Some(TrPr::from_xml_element(child_node)?),
                node_name if ContentCellContent::is_choice_member(node_name) => instance
                    .contents
                    .push(ContentCellContent::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlRow {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub contents: Vec<ContentRowContent>,
    pub uri: Option<String>,
    pub element: XmlName,
}

impl CustomXmlRow {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing CustomXmlRow");

        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:uri" => uri = Some(value.clone()),
                "w:element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        let mut custom_xml_properties = None;
        let mut contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name if ContentRowContent::is_choice_member(node_name) => {
                    contents.push(ContentRowContent::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        Ok(Self {
            custom_xml_properties,
            contents,
            uri,
            element,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentRow {
    pub contents: Vec<ContentRowContent>,
}

impl SdtContentRow {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtContentRow");

        let contents = xml_node
            .child_nodes
            .iter()
            .filter_map(ContentRowContent::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtRow {
    pub properties: Option<SdtPr>,
    pub end_properties: Option<SdtEndPr>,
    pub content: Option<SdtContentRow>,
}

impl SdtRow {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing SdtRow");

        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "sdtPr" => instance.properties = Some(SdtPr::from_xml_element(child_node)?),
                    "sdtEndPr" => instance.end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                    "sdtContent" => instance.content = Some(SdtContentRow::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ContentRowContent {
    Table(Box<Row>),
    CustomXml(CustomXmlRow),
    Sdt(Box<SdtRow>),
    RunLevelElements(RunLevelElts),
}

impl XsdType for ContentRowContent {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tr" => Ok(ContentRowContent::Table(Box::new(Row::from_xml_element(xml_node)?))),
            "customXml" => Ok(ContentRowContent::CustomXml(CustomXmlRow::from_xml_element(xml_node)?)),
            "sdt" => Ok(ContentRowContent::Sdt(Box::new(SdtRow::from_xml_element(xml_node)?))),
            node_name if RunLevelElts::is_choice_member(node_name) => Ok(ContentRowContent::RunLevelElements(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentRowContent",
            ))),
        }
    }
}

impl XsdChoice for ContentRowContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "tr" | "customXml" | "sdt" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Height {
    pub value: Option<TwipsMeasure>,
    pub height_rule: Option<HeightRule>,
}

impl Height {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Height");

        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w:val" => instance.value = Some(value.parse()?),
                "w:hRule" => instance.height_rule = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Tbl {
    pub range_markup_elements: Vec<RangeMarkupElements>,
    pub properties: TblPr,
    pub grid: TblGrid,
    pub row_contents: Vec<ContentRowContent>,
}

impl Tbl {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        info!("parsing Tbl");

        let mut range_markup_elements = Vec::new();
        let mut properties = None;
        let mut grid = None;
        let mut row_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPr" => properties = Some(TblPr::from_xml_element(child_node)?),
                "tblGrid" => grid = Some(TblGrid::from_xml_element(child_node)?),
                node_name => {
                    if RangeMarkupElements::is_choice_member(node_name) {
                        range_markup_elements.push(RangeMarkupElements::from_xml_element(child_node)?);
                    } else if ContentRowContent::is_choice_member(node_name) {
                        row_contents.push(ContentRowContent::from_xml_element(child_node)?);
                    }
                }
            }
        }

        let properties = properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblPr"))?;
        let grid = grid.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblGrid"))?;

        Ok(Self {
            range_markup_elements,
            properties,
            grid,
            row_contents,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::document::{Bookmark, ContentBlockContent, DecimalNumberOrPercent, ProofErr};
    use std::str::FromStr;

    impl TblPPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:leftFromText="10" w:rightFromText="10" w:topFromText="10" w:bottomFromText="10"
                w:vertAnchor="text" w:horzAnchor="text" w:tblpXSpec="left" w:tblpX="10" w:tblpYSpec="top" w:tblpY="10">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                left_from_text: Some(TwipsMeasure::Decimal(10)),
                right_from_text: Some(TwipsMeasure::Decimal(10)),
                top_from_text: Some(TwipsMeasure::Decimal(10)),
                bottom_from_text: Some(TwipsMeasure::Decimal(10)),
                vertical_anchor: Some(VAnchor::Text),
                horizontal_anchor: Some(HAnchor::Text),
                horizontal_alignment: Some(XAlign::Left),
                horizontal_distance: Some(SignedTwipsMeasure::Decimal(10)),
                vertical_alignment: Some(YAlign::Top),
                vertical_distance: Some(SignedTwipsMeasure::Decimal(10)),
            }
        }
    }

    #[test]
    pub fn test_tbl_p_pr_from_xml() {
        let xml = TblPPr::test_xml("tblPPr");
        assert_eq!(
            TblPPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPPr::test_instance(),
        );
    }

    impl TblWidth {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:w="100" w:type="auto"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(
                    100,
                ))),
                width_type: Some(TblWidthType::Auto),
            }
        }
    }

    #[test]
    pub fn test_tbl_width_from_xml() {
        let xml = TblWidth::test_xml("tblWidth");
        assert_eq!(
            TblWidth::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblWidth::test_instance(),
        );
    }

    impl TblBorders {
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
                Border::test_xml("start"),
                Border::test_xml("bottom"),
                Border::test_xml("end"),
                Border::test_xml("insideH"),
                Border::test_xml("insideV"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                start: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                end: Some(Border::test_instance()),
                inside_horizontal: Some(Border::test_instance()),
                inside_vertical: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_borders_from_xml() {
        let xml = TblBorders::test_xml("tblBorders");
        assert_eq!(
            TblBorders::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblBorders::test_instance(),
        );
    }

    impl TblCellMar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TblWidth::test_xml("top"),
                TblWidth::test_xml("start"),
                TblWidth::test_xml("bottom"),
                TblWidth::test_xml("end"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(TblWidth::test_instance()),
                start: Some(TblWidth::test_instance()),
                bottom: Some(TblWidth::test_instance()),
                end: Some(TblWidth::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_cell_mar_from_xml() {
        let xml = TblCellMar::test_xml("tblCellMar");
        assert_eq!(
            TblCellMar::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblCellMar::test_instance(),
        );
    }

    impl TblLook {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:firstRow="true" w:lastRow="true" w:firstColumn="true" w:lastColumn="true" w:noHBand="true" w:noVBand="true">
            </{node_name}>"#,
                node_name=node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first_row: Some(true),
                last_row: Some(true),
                first_column: Some(true),
                last_column: Some(true),
                no_horizontal_band: Some(true),
                no_vertical_band: Some(true),
            }
        }
    }

    #[test]
    pub fn test_tbl_look_from_xml() {
        let xml = TblLook::test_xml("tblLook");
        assert_eq!(
            TblLook::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblLook::test_instance(),
        );
    }

    impl TblPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"<tblStyle w:val="Normal" />
                {}
                <tblOverlap w:val="never" />
                <bidiVisual w:val="false" />
                <tblStyleRowBandSize w:val="100" />
                <tblStyleColBandSize w:val="100" />
                {}
                <jc w:val="center" />
                {}
                {}
                {}
                {}
                <tblLayout w:val="autofit" />
                {}
                {}
                <tblCaption w:val="Some caption" />
                <tblDescription w:val="Some description" />"#,
                TblPPr::test_xml("tblpPr"),
                TblWidth::test_xml("tblW"),
                TblWidth::test_xml("tblCellSpacing"),
                TblWidth::test_xml("tblInd"),
                TblBorders::test_xml("tblBorders"),
                Shd::test_xml("shd"),
                TblCellMar::test_xml("tblCellMar"),
                TblLook::test_xml("tblLook"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                style: Some(String::from("Normal")),
                paragraph_properties: Some(TblPPr::test_instance()),
                overlap: Some(TblOverlap::Never),
                bidirectional_visual: Some(false),
                style_row_band_size: Some(100),
                style_column_band_size: Some(100),
                width: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                cell_spacing: Some(TblWidth::test_instance()),
                indent: Some(TblWidth::test_instance()),
                borders: Some(TblBorders::test_instance()),
                shading: Some(Shd::test_instance()),
                layout: Some(TblLayoutType::Autofit),
                cell_margin: Some(TblCellMar::test_instance()),
                look: Some(TblLook::test_instance()),
                caption: Some(String::from("Some caption")),
                description: Some(String::from("Some description")),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_base_from_xml() {
        let xml = TblPrBase::test_xml("tblPrBase");
        assert_eq!(
            TblPrBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPrBase::test_instance(),
        );
    }

    impl TblPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TblPrBase::test_xml("tblPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: TblPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_change_from_xml() {
        let xml = TblPrChange::test_xml("tblPrChange");
        assert_eq!(
            TblPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPrChange::test_instance(),
        );
    }

    impl TblPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblPrBase::test_extension_xml(),
                TblPrChange::test_xml("tblPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblPrBase::test_instance(),
                change: Some(TblPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_from_xml() {
        let xml = TblPr::test_xml("tblPr");
        assert_eq!(
            TblPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPr::test_instance(),
        );
    }

    impl TblGridCol {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w:w="100"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_col_from_xml() {
        let xml = TblGridCol::test_xml("tblGridCol");
        assert_eq!(
            TblGridCol::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblGridCol::test_instance(),
        );
    }

    impl TblGridBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_extension_xml() -> String {
            format!("{grid_col}{grid_col}", grid_col = TblGridCol::test_xml("gridCol"))
        }

        pub fn test_instance() -> Self {
            Self {
                columns: vec![TblGridCol::test_instance(), TblGridCol::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_base_from_xml() {
        let xml = TblGridBase::test_xml("tblGridBase");
        assert_eq!(
            TblGridBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblGridBase::test_instance(),
        );
    }

    impl TblGridChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="0">{}</{node_name}>"#,
                TblGridBase::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Markup::test_instance(),
                grid: TblGridBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_change_from_xml() {
        let xml = TblGridChange::test_xml("tblGridChange");
        assert_eq!(
            TblGridChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblGridChange::test_instance(),
        );
    }

    impl TblGrid {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblGridBase::test_extension_xml(),
                TblGridChange::test_xml("tblGridChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblGridBase::test_instance(),
                change: Some(TblGridChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_from_xml() {
        let xml = TblGrid::test_xml("tblGrid");
        assert_eq!(
            TblGrid::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblGrid::test_instance(),
        );
    }

    impl TblPrExBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"
                {}
                <jc w:val="center" />
                {}
                {}
                {}
                {}
                <tblLayout w:val="autofit" />
                {}
                {}"#,
                TblWidth::test_xml("tblW"),
                TblWidth::test_xml("tblCellSpacing"),
                TblWidth::test_xml("tblInd"),
                TblBorders::test_xml("tblBorders"),
                Shd::test_xml("shd"),
                TblCellMar::test_xml("tblCellMar"),
                TblLook::test_xml("tblLook"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                cell_spacing: Some(TblWidth::test_instance()),
                indent: Some(TblWidth::test_instance()),
                borders: Some(TblBorders::test_instance()),
                shading: Some(Shd::test_instance()),
                layout: Some(TblLayoutType::Autofit),
                cell_margin: Some(TblCellMar::test_instance()),
                look: Some(TblLook::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_base_from_xml() {
        let xml = TblPrExBase::test_xml("tblPrExBase");
        assert_eq!(
            TblPrExBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPrExBase::test_instance(),
        );
    }

    impl TblPrExChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TblPrExBase::test_xml("tblPrEx"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties_ex: TblPrExBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_change_from_xml() {
        let xml = TblPrExChange::test_xml("tblPrExChange");
        assert_eq!(
            TblPrExChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPrExChange::test_instance(),
        );
    }

    impl TblPrEx {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblPrExBase::test_extension_xml(),
                TblPrExChange::test_xml("tblPrExChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblPrExBase::test_instance(),
                change: Some(TblPrExChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_from_xml() {
        let xml = TblPrEx::test_xml("tblPrEx");
        assert_eq!(
            TblPrEx::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TblPrEx::test_instance(),
        );
    }

    impl TrPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"{}
                <divId w:val="1" />
                <gridBefore w:val="1" />
                <gridAfter w:val="1" />
                {}
                {}
                <cantSplit w:val="false" />
                {}
                <tblHeader w:val="true" />
                {}
                <jc w:val="center" />
                <hidden w:val="false" />"#,
                Cnf::test_xml("cnfStyle"),
                TblWidth::test_xml("wBefore"),
                TblWidth::test_xml("wAfter"),
                Height::test_xml("trHeight"),
                TblWidth::test_xml("tblCellSpacing"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                conditional_formatting: Some(Cnf::test_instance()),
                div_id: Some(1),
                grid_column_before_first_cell: Some(1),
                grid_column_after_last_cell: Some(1),
                width_before_row: Some(TblWidth::test_instance()),
                width_after_row: Some(TblWidth::test_instance()),
                cant_split: Some(false),
                row_height: Some(Height::test_instance()),
                header: Some(true),
                cell_spacing: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                hidden: Some(false),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_base_from_xml() {
        let xml = TrPrBase::test_xml("trPrBase");
        assert_eq!(
            TrPrBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TrPrBase::test_instance(),
        );
    }

    impl TrPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>{}</{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TrPrBase::test_xml("trPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: TrPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_change_from_xml() {
        let xml = TrPrChange::test_xml("trPrChange");
        assert_eq!(
            TrPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TrPrChange::test_instance(),
        );
    }

    impl TrPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TrPrBase::test_extension_xml(),
                TrackChange::test_xml("ins"),
                TrackChange::test_xml("del"),
                TrPrChange::test_xml("trPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrPrBase::test_instance(),
                inserted: Some(TrackChange::test_instance()),
                deleted: Some(TrackChange::test_instance()),
                change: Some(TrPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_from_xml() {
        let xml = TrPr::test_xml("trPr");
        assert_eq!(
            TrPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TrPr::test_instance(),
        );
    }

    impl TcBorders {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Border::test_xml("top"),
                Border::test_xml("start"),
                Border::test_xml("bottom"),
                Border::test_xml("end"),
                Border::test_xml("insideH"),
                Border::test_xml("insideV"),
                Border::test_xml("tl2br"),
                Border::test_xml("tr2bl"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                start: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                end: Some(Border::test_instance()),
                inside_horizontal: Some(Border::test_instance()),
                inside_vertical: Some(Border::test_instance()),
                top_left_to_bottom_right: Some(Border::test_instance()),
                top_right_to_bottom_left: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tc_borders_from_xml() {
        let xml = TcBorders::test_xml("tcBorders");
        assert_eq!(
            TcBorders::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcBorders::test_instance(),
        );
    }

    impl TcMar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>",
                TblWidth::test_xml("top"),
                TblWidth::test_xml("start"),
                TblWidth::test_xml("bottom"),
                TblWidth::test_xml("end"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(TblWidth::test_instance()),
                start: Some(TblWidth::test_instance()),
                bottom: Some(TblWidth::test_instance()),
                end: Some(TblWidth::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tc_mar_from_xml() {
        let xml = TcMar::test_xml("tcMar");
        assert_eq!(
            TcMar::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcMar::test_instance(),
        );
    }

    impl Headers {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}><header w:val="Header1" /><header w:val="Header2" /></{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![String::from("Header1"), String::from("Header2")])
        }
    }

    #[test]
    pub fn test_headers_from_xml() {
        let xml = Headers::test_xml("headers");
        assert_eq!(
            Headers::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Headers::test_instance(),
        );
    }

    impl TcPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"{}
                {}
                <gridSpan w:val="1" />
                <vMerge w:val="continue" />
                {}
                {}
                <noWrap w:val="true" />
                {}
                <textDirection w:val="lr" />
                <tcFitText w:val="true" />
                <vAlign w:val="top" />
                <hideMark w:val="true" />
                {}"#,
                Cnf::test_xml("cnfStyle"),
                TblWidth::test_xml("tcW"),
                TcBorders::test_xml("tcBorders"),
                Shd::test_xml("shd"),
                TcMar::test_xml("tcMar"),
                Headers::test_xml("headers"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                conditional_formatting: Some(Cnf::test_instance()),
                width: Some(TblWidth::test_instance()),
                grid_span: Some(1),
                vertical_merge: Some(Merge::Continue),
                borders: Some(TcBorders::test_instance()),
                shading: Some(Shd::test_instance()),
                no_wrapping: Some(true),
                margin: Some(TcMar::test_instance()),
                text_direction: Some(TextDirection::LeftToRight),
                fit_text: Some(true),
                vertical_alignment: Some(VerticalJc::Top),
                hide_marker: Some(true),
                headers: Some(Headers::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tc_pr_base_from_xml() {
        let xml = TcPrBase::test_xml("tcPrBase");
        assert_eq!(
            TcPrBase::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcPrBase::test_instance(),
        );
    }

    impl CellMergeTrackChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} w:vMerge="cont" w:vMergeOrig="cont"></{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                vertical_merge: Some(AnnotationVMerge::Merge),
                vertical_merge_original: Some(AnnotationVMerge::Merge),
            }
        }
    }

    #[test]
    pub fn test_cell_merge_track_change_from_xml() {
        let xml = CellMergeTrackChange::test_xml("cellMergeTrackChange");
        assert_eq!(
            CellMergeTrackChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            CellMergeTrackChange::test_instance(),
        );
    }

    impl TcPrInner {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!("{}{}", TcPrBase::test_extension_xml(), TrackChange::test_xml("cellIns"),)
        }

        pub fn test_instance() -> Self {
            Self {
                base: TcPrBase::test_instance(),
                markup_element: Some(CellMarkupElements::Insertion(TrackChange::test_instance())),
            }
        }
    }

    #[test]
    pub fn test_tc_pr_inner_from_xml() {
        let xml = TcPrInner::test_xml("tcPrInner");
        assert_eq!(
            TcPrInner::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcPrInner::test_instance(),
        );
    }

    impl TcPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>{}</{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TcPrInner::test_xml("tcPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: TcPrInner::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tc_pr_change_from_xml() {
        let xml = TcPrChange::test_xml("tcPrChange");
        assert_eq!(
            TcPrChange::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcPrChange::test_instance(),
        );
    }

    impl TcPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TcPrInner::test_extension_xml(),
                TcPrChange::test_xml("tcPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TcPrInner::test_instance(),
                change: Some(TcPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tc_pr_from_xml() {
        let xml = TcPr::test_xml("tcPr");
        assert_eq!(
            TcPr::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TcPr::test_instance(),
        );
    }

    impl Tc {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:id="Some id">
                {}
                {}
            </{node_name}>"#,
                TcPr::test_xml("tcPr"),
                ProofErr::test_xml("proofErr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(TcPr::test_instance()),
                block_level_elements: vec![BlockLevelElts::Chunk(ContentBlockContent::RunLevelElement(
                    RunLevelElts::ProofError(ProofErr::test_instance()),
                ))],
                id: Some(String::from("Some id")),
            }
        }
    }

    #[test]
    pub fn test_tc_from_xml() {
        let xml = Tc::test_xml("tc");
        assert_eq!(
            Tc::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Tc::test_instance(),
        );
    }

    impl CustomXmlCell {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="https://some/uri" w:element="Xml name">
                {}
                {}
            </{node_name}>"#,
                CustomXmlPr::test_xml("customXmlPr"),
                Tc::test_xml("tc"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_xml_properties: Some(CustomXmlPr::test_instance()),
                contents: vec![ContentCellContent::Cell(Box::new(Tc::test_instance()))],
                uri: Some(String::from("https://some/uri")),
                element: XmlName::from("Xml name"),
            }
        }
    }

    #[test]
    pub fn test_custom_xml_cell_from_xml() {
        let xml = CustomXmlCell::test_xml("customXmlCell");
        assert_eq!(
            CustomXmlCell::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            CustomXmlCell::test_instance(),
        );
    }

    impl SdtContentCell {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {tc}
                {tc}
            </{node_name}>"#,
                tc = Tc::test_xml("tc"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                contents: vec![
                    ContentCellContent::Cell(Box::new(Tc::test_instance())),
                    ContentCellContent::Cell(Box::new(Tc::test_instance())),
                ],
            }
        }
    }

    #[test]
    pub fn test_std_content_cell_from_xml() {
        let xml = SdtContentCell::test_xml("sdtContentCell");
        assert_eq!(
            SdtContentCell::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtContentCell::test_instance(),
        );
    }

    impl SdtCell {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentCell::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(SdtPr::test_instance()),
                end_properties: Some(SdtEndPr::test_instance()),
                content: Some(SdtContentCell::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_std_cell_from_xml() {
        let xml = SdtCell::test_xml("sdtCell");
        assert_eq!(
            SdtCell::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtCell::test_instance(),
        );
    }

    impl Row {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:rsidRPr="ffffffff" w:rsidR="fefefefe" w:rsidDel="fdfdfdfd" w:rsidTr="fcfcfcfc">
                {}
                {}
                {}
            </{node_name}>"#,
                TblPrEx::test_xml("tblPrEx"),
                TrPr::test_xml("trPr"),
                Tc::test_xml("tc"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                property_exceptions: Some(TblPrEx::test_instance()),
                properties: Some(TrPr::test_instance()),
                contents: vec![ContentCellContent::Cell(Box::new(Tc::test_instance()))],
                run_properties_revision_id: Some(0xffffffff),
                run_revision_id: Some(0xfefefefe),
                deletion_revision_id: Some(0xfdfdfdfd),
                row_revision_id: Some(0xfcfcfcfc),
            }
        }
    }

    #[test]
    pub fn test_row_from_xml() {
        let xml = Row::test_xml("row");
        assert_eq!(
            Row::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Row::test_instance(),
        );
    }

    impl CustomXmlRow {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:uri="https://some/uri" w:element="Xml name">
                {}
                {}
            </{node_name}>"#,
                CustomXmlPr::test_xml("customXmlPr"),
                Row::test_xml("tr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_xml_properties: Some(CustomXmlPr::test_instance()),
                contents: vec![ContentRowContent::Table(Box::new(Row::test_instance()))],
                uri: Some(String::from("https://some/uri")),
                element: String::from("Xml name"),
            }
        }
    }

    #[test]
    pub fn test_custom_xml_row_from_xml() {
        let xml = CustomXmlRow::test_xml("customXmlRow");
        assert_eq!(
            CustomXmlRow::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            CustomXmlRow::test_instance(),
        );
    }

    impl SdtContentRow {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {tr}
                {tr}
            </{node_name}>"#,
                tr = Row::test_xml("tr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                contents: vec![
                    ContentRowContent::Table(Box::new(Row::test_instance())),
                    ContentRowContent::Table(Box::new(Row::test_instance())),
                ],
            }
        }
    }

    #[test]
    pub fn test_sdt_content_row_from_xml() {
        let xml = SdtContentRow::test_xml("sdtContentRow");
        assert_eq!(
            SdtContentRow::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtContentRow::test_instance(),
        );
    }

    impl SdtRow {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentRow::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(SdtPr::test_instance()),
                end_properties: Some(SdtEndPr::test_instance()),
                content: Some(SdtContentRow::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_std_row_from_xml() {
        let xml = SdtRow::test_xml("sdtRow");
        assert_eq!(
            SdtRow::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            SdtRow::test_instance(),
        );
    }

    impl Height {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w:val="100" w:hRule="auto"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(TwipsMeasure::Decimal(100)),
                height_rule: Some(HeightRule::Auto),
            }
        }
    }

    #[test]
    pub fn test_height_from_xml() {
        let xml = Height::test_xml("height");
        assert_eq!(
            Height::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Height::test_instance(),
        );
    }

    impl Tbl {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Bookmark::test_xml("bookmarkStart"),
                TblPr::test_xml("tblPr"),
                TblGrid::test_xml("tblGrid"),
                Row::test_xml("tr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                range_markup_elements: vec![RangeMarkupElements::BookmarkStart(Bookmark::test_instance())],
                properties: TblPr::test_instance(),
                grid: TblGrid::test_instance(),
                row_contents: vec![ContentRowContent::Table(Box::new(Row::test_instance()))],
            }
        }
    }

    #[test]
    pub fn test_tbl_from_xml() {
        let xml = Tbl::test_xml("tbl");
        assert_eq!(
            Tbl::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Tbl::test_instance(),
        );
    }
}
