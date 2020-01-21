use crate::{
    shared::{
        drawingml::{
            coordsys::{Point2D, PositiveSize2D, Transform2D},
            core::{
                GraphicalObject, GroupShapeProperties, NonVisualConnectorProperties, NonVisualContentPartProperties,
                NonVisualDrawingProps, NonVisualDrawingShapeProps, NonVisualGraphicFrameProperties,
                NonVisualGroupDrawingShapeProps, ShapeProperties, ShapeStyle,
            },
            diagrams::{BackgroundFormatting, WholeE2oFormatting},
            picture::Picture,
            simpletypes::{BlackWhiteMode, Coordinate},
            text::bodyformatting::TextBodyProperties,
        },
        relationship::RelationshipId,
    },
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::XsdChoice,
};

type Result<T> = ::std::result::Result<T, Box<dyn std::error::Error>>;

type PositionOffset = i32;
type WrapDistance = u32;

#[derive(Debug, Clone, PartialEq)]
pub struct EffectExtent {
    pub left: Coordinate,
    pub top: Coordinate,
    pub right: Coordinate,
    pub bottom: Coordinate,
}

impl EffectExtent {
    pub fn new(left: Coordinate, top: Coordinate, right: Coordinate, bottom: Coordinate) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut left = None;
        let mut top = None;
        let mut right = None;
        let mut bottom = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "l" => left = Some(value.parse()?),
                "t" => top = Some(value.parse()?),
                "r" => right = Some(value.parse()?),
                "b" => bottom = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            left: left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?,
            top: top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "t"))?,
            right: right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?,
            bottom: bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Inline {
    pub extent: PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub doc_properties: NonVisualDrawingProps,
    pub graphic_frame_properties: Option<NonVisualGraphicFrameProperties>,
    pub graphic: GraphicalObject,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl Inline {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut distance_top = None;
        let mut distance_bottom = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let mut extent = None;
        let mut effect_extent = None;
        let mut doc_properties = None;
        let mut graphic_frame_properties = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "extent" => extent = Some(PositiveSize2D::from_xml_element(child_node)?),
                "effectExtent" => effect_extent = Some(EffectExtent::from_xml_element(child_node)?),
                "docPr" => doc_properties = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvGraphicFramePr" => {
                    graphic_frame_properties = Some(NonVisualGraphicFrameProperties::from_xml_element(child_node)?)
                }
                "graphic" => graphic = Some(GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            extent: extent.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "extent"))?,
            effect_extent,
            doc_properties: doc_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPr"))?,
            graphic_frame_properties,
            graphic: graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?,
            distance_top,
            distance_bottom,
            distance_left,
            distance_right,
        })
    }
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum WrapText {
    #[strum(serialize = "bothSides")]
    BothSides,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "largest")]
    Largest,
}

#[derive(Debug, Clone, PartialEq)]
pub struct WrapPath {
    pub start: Point2D,
    pub line_to: Vec<Point2D>,

    pub edited: Option<bool>,
}

impl WrapPath {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let edited = xml_node
            .attributes
            .get("edited")
            .map(|value| value.parse())
            .transpose()?;

        let mut start = None;
        let mut line_to = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "start" => start = Some(Point2D::from_xml_element(child_node)?),
                "lineTo" => line_to.push(Point2D::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let start = start.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "start"))?;
        match line_to.len() {
            occurs if occurs >= 2 => Ok(Self { start, line_to, edited }),
            occurs => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "lineTo",
                2,
                MaxOccurs::Unbounded,
                occurs as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WrapSquare {
    pub effect_extent: Option<EffectExtent>,

    pub wrap_text: WrapText,
    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl WrapSquare {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut wrap_text = None;
        let mut distance_top = None;
        let mut distance_bottom = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "wrapText" => wrap_text = Some(value.parse()?),
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let effect_extent = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "effectExtent")
            .map(EffectExtent::from_xml_element)
            .transpose()?;

        Ok(Self {
            effect_extent,
            wrap_text: wrap_text.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wrapText"))?,
            distance_top,
            distance_bottom,
            distance_left,
            distance_right,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WrapTight {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl WrapTight {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut wrap_text = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "wrapText" => wrap_text = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let wrap_polygon = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "wrapPolygon")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "wrapPolygon").into())
            .and_then(WrapPath::from_xml_element)?;

        Ok(Self {
            wrap_polygon,
            wrap_text: wrap_text.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wrapText"))?,
            distance_left,
            distance_right,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WrapThrough {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl WrapThrough {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut wrap_text = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "wrapText" => wrap_text = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let wrap_polygon = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "wrapPolygon")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "wrapPolygon").into())
            .and_then(WrapPath::from_xml_element)?;

        Ok(Self {
            wrap_polygon,
            wrap_text: wrap_text.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wrapText"))?,
            distance_left,
            distance_right,
        })
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct WrapTopBottom {
    pub effect_extent: Option<EffectExtent>,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
}

impl WrapTopBottom {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut distance_top = None;
        let mut distance_bottom = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                _ => (),
            }
        }

        let effect_extent = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "effectExtent")
            .map(EffectExtent::from_xml_element)
            .transpose()?;

        Ok(Self {
            effect_extent,
            distance_top,
            distance_bottom,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum WrapType {
    None,
    Square(WrapSquare),
    Tight(WrapTight),
    Through(WrapThrough),
    TopAndBottom(WrapTopBottom),
}

impl WrapType {
    pub fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "wrapNone" | "wrapSquare" | "wrapTight" | "wrapThrough" | "wrapTopAndBottom" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "wrapNone" => Ok(WrapType::None),
            "wrapSquare" => Ok(WrapType::Square(WrapSquare::from_xml_element(xml_node)?)),
            "wrapTight" => Ok(WrapType::Tight(WrapTight::from_xml_element(xml_node)?)),
            "wrapThrough" => Ok(WrapType::Through(WrapThrough::from_xml_element(xml_node)?)),
            "wrapTopAndBottom" => Ok(WrapType::TopAndBottom(WrapTopBottom::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "WrapType"))),
        }
    }
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum AlignH {
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum RelFromH {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "column")]
    Column,
    #[strum(serialize = "character")]
    Character,
    #[strum(serialize = "leftMargin")]
    LeftMargin,
    #[strum(serialize = "rightMargin")]
    RightMargin,
    #[strum(serialize = "insideMargin")]
    InsideMargin,
    #[strum(serialize = "outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PosHChoice {
    Align(AlignH),
    PositionOffset(PositionOffset),
}

impl PosHChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "align" | "posOffset" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "align" => {
                let alignment = xml_node
                    .text
                    .as_ref()
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
                    .parse()?;

                Ok(PosHChoice::Align(alignment))
            }
            "posOffset" => {
                let offset = xml_node
                    .text
                    .as_ref()
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
                    .parse()?;

                Ok(PosHChoice::PositionOffset(offset))
            }
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "PosHChoice"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PosH {
    pub align_or_offset: PosHChoice,
    pub relative_from: RelFromH,
}

impl PosH {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let relative_from = xml_node
            .attributes
            .get("relativeFrom")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "relativeFrom"))?
            .parse()?;

        let align_or_offset = xml_node
            .child_nodes
            .iter()
            .find(|child_node| PosHChoice::is_choice_member(child_node.local_name()))
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "align|posOffset").into())
            .and_then(PosHChoice::from_xml_element)?;

        Ok(Self {
            align_or_offset,
            relative_from,
        })
    }
}
#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum AlignV {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "bottom")]
    Bottom,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum RelFromV {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "paragraph")]
    Paragraph,
    #[strum(serialize = "line")]
    Line,
    #[strum(serialize = "topMargin")]
    TopMargin,
    #[strum(serialize = "bottomMargin")]
    BottomMargin,
    #[strum(serialize = "insideMargin")]
    InsideMargin,
    #[strum(serialize = "outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PosVChoice {
    Align(AlignV),
    PositionOffset(PositionOffset),
}

impl PosVChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "align" | "posOffset" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "align" => {
                let alignment = xml_node
                    .text
                    .as_ref()
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
                    .parse()?;

                Ok(PosVChoice::Align(alignment))
            }
            "posOffset" => {
                let offset = xml_node
                    .text
                    .as_ref()
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
                    .parse()?;

                Ok(PosVChoice::PositionOffset(offset))
            }
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "PosVChoice"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PosV {
    pub align_or_offset: PosVChoice,
    pub relative_from: RelFromV,
}

impl PosV {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let relative_from = xml_node
            .attributes
            .get("relativeFrom")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "relativeFrom"))?
            .parse()?;

        let align_or_offset = xml_node
            .child_nodes
            .iter()
            .find(|child_node| PosVChoice::is_choice_member(child_node.local_name()))
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "align|posOffset").into())
            .and_then(PosVChoice::from_xml_element)?;

        Ok(Self {
            align_or_offset,
            relative_from,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Anchor {
    pub simple_position: Point2D,
    pub horizontal_position: PosH,
    pub vertical_position: PosV,
    pub extent: PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub wrap_type: WrapType,
    pub document_properties: NonVisualDrawingProps,
    pub graphic_frame_properties: Option<NonVisualGraphicFrameProperties>,
    pub graphic: GraphicalObject,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
    pub use_simple_position: Option<bool>,
    pub relative_height: u32,
    pub behind_document_text: bool,
    pub locked: bool,
    pub layout_in_cell: bool,
    pub hidden: Option<bool>,
    pub allow_overlap: bool,
}

impl Anchor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut distance_top = None;
        let mut distance_bottom = None;
        let mut distance_left = None;
        let mut distance_right = None;
        let mut use_simple_position = None;
        let mut relative_height = None;
        let mut behind_document_text = None;
        let mut locked = None;
        let mut layout_in_cell = None;
        let mut hidden = None;
        let mut allow_overlap = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                "simplePos" => use_simple_position = Some(parse_xml_bool(value)?),
                "relativeHeight" => relative_height = Some(value.parse()?),
                "behindDoc" => behind_document_text = Some(parse_xml_bool(value)?),
                "locked" => locked = Some(parse_xml_bool(value)?),
                "layoutInCell" => layout_in_cell = Some(parse_xml_bool(value)?),
                "hidden" => hidden = Some(parse_xml_bool(value)?),
                "allowOverlap" => allow_overlap = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut simple_position = None;
        let mut horizontal_position = None;
        let mut vertical_position = None;
        let mut extent = None;
        let mut effect_extent = None;
        let mut wrap_type = None;
        let mut document_properties = None;
        let mut graphic_frame_properties = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "simplePos" => simple_position = Some(Point2D::from_xml_element(child_node)?),
                "positionH" => horizontal_position = Some(PosH::from_xml_element(child_node)?),
                "positionV" => vertical_position = Some(PosV::from_xml_element(child_node)?),
                "extent" => extent = Some(PositiveSize2D::from_xml_element(child_node)?),
                "effectExtent" => effect_extent = Some(EffectExtent::from_xml_element(child_node)?),
                node_name if WrapType::is_choice_member(node_name) => {
                    wrap_type = Some(WrapType::from_xml_element(child_node)?)
                }
                "docPr" => document_properties = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvGraphicFramePr" => {
                    graphic_frame_properties = Some(NonVisualGraphicFrameProperties::from_xml_element(child_node)?)
                }
                "graphic" => graphic = Some(GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let simple_position =
            simple_position.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "simplePos"))?;
        let horizontal_position =
            horizontal_position.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "positionH"))?;
        let vertical_position =
            vertical_position.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "positionV"))?;
        let extent = extent.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "extent"))?;
        let wrap_type = wrap_type.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "WrapType"))?;
        let document_properties =
            document_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPr"))?;
        let graphic = graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?;
        let relative_height =
            relative_height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "relativeHeight"))?;
        let behind_document_text =
            behind_document_text.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "behindDoc"))?;
        let locked = locked.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "locked"))?;
        let layout_in_cell =
            layout_in_cell.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "layoutInCell"))?;
        let allow_overlap =
            allow_overlap.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "allowOverlap"))?;

        Ok(Self {
            simple_position,
            horizontal_position,
            vertical_position,
            extent,
            effect_extent,
            wrap_type,
            document_properties,
            graphic_frame_properties,
            graphic,
            distance_top,
            distance_bottom,
            distance_left,
            distance_right,
            use_simple_position,
            relative_height,
            behind_document_text,
            locked,
            layout_in_cell,
            hidden,
            allow_overlap,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TxbxContent {
    pub block_level_elements: Vec<super::document::BlockLevelElts>,
}

impl TxbxContent {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let block_level_elements = xml_node
            .child_nodes
            .iter()
            .filter_map(super::document::BlockLevelElts::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        if block_level_elements.is_empty() {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "BlockLevelElts",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        } else {
            Ok(Self { block_level_elements })
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextboxInfo {
    pub textbox_content: TxbxContent,
    pub id: Option<u16>, // default=0,
}

impl TextboxInfo {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let id = xml_node.attributes.get("id").map(|value| value.parse()).transpose()?;

        let textbox_content = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "txbxContent")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "txbxContent").into())
            .and_then(TxbxContent::from_xml_element)?;

        Ok(Self { textbox_content, id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct LinkedTextboxInformation {
    pub id: u16,
    pub sequence: u16,
}

impl LinkedTextboxInformation {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut sequence = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "id" => id = Some(value.parse()?),
                "seq" => sequence = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            id: id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?,
            sequence: sequence.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "seq"))?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum WordprocessingShapePropertiesChoice {
    ShapeProperties(NonVisualDrawingShapeProps),
    Connector(NonVisualConnectorProperties),
}

#[derive(Debug, Clone, PartialEq)]
pub enum WordprocessingShapeTextboxInfoChoice {
    Textbox(TextboxInfo),
    LinkedTextbox(LinkedTextboxInformation),
}

#[derive(Debug, Clone, PartialEq)]
pub struct WordprocessingShape {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub properties: WordprocessingShapePropertiesChoice,
    pub shape_properties: ShapeProperties,
    pub style: Option<ShapeStyle>,
    pub text_box_info: Option<WordprocessingShapeTextboxInfoChoice>,
    pub text_body_properties: TextBodyProperties,

    pub normal_east_asian_flow: Option<bool>, // default=false
}

impl WordprocessingShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let normal_east_asian_flow = xml_node
            .attributes
            .get("normalEastAsianFlow")
            .map(parse_xml_bool)
            .transpose()?;

        let mut non_visual_drawing_props = None;
        let mut properties = None;
        let mut shape_properties = None;
        let mut style = None;
        let mut text_box_info = None;
        let mut text_body_properties = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => non_visual_drawing_props = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvSpPr" => {
                    properties = Some(WordprocessingShapePropertiesChoice::ShapeProperties(
                        NonVisualDrawingShapeProps::from_xml_element(child_node)?,
                    ))
                }
                "cNvCnPr" => {
                    properties = Some(WordprocessingShapePropertiesChoice::Connector(
                        NonVisualConnectorProperties::from_xml_element(child_node)?,
                    ))
                }
                "spPr" => shape_properties = Some(ShapeProperties::from_xml_element(child_node)?),
                "style" => style = Some(ShapeStyle::from_xml_element(child_node)?),
                "txbx" => {
                    text_box_info = Some(WordprocessingShapeTextboxInfoChoice::Textbox(
                        TextboxInfo::from_xml_element(child_node)?,
                    ))
                }
                "linkedTxbx" => {
                    text_box_info = Some(WordprocessingShapeTextboxInfoChoice::LinkedTextbox(
                        LinkedTextboxInformation::from_xml_element(child_node)?,
                    ))
                }
                "bodyPr" => text_body_properties = Some(TextBodyProperties::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let properties =
            properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvSpPr|cNvCnPr"))?;
        let shape_properties =
            shape_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;
        let text_body_properties =
            text_body_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bodyPr"))?;

        Ok(Self {
            non_visual_drawing_props,
            properties,
            shape_properties,
            style,
            text_box_info,
            text_body_properties,
            normal_east_asian_flow,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GraphicFrame {
    pub non_visual_drawing_props: NonVisualDrawingProps,
    pub non_visual_props: NonVisualGraphicFrameProperties,
    pub transform: Transform2D,
    pub graphic: GraphicalObject,
}

impl GraphicFrame {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_drawing_props = None;
        let mut non_visual_props = None;
        let mut transform = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => non_visual_drawing_props = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvFrPr" => non_visual_props = Some(NonVisualGraphicFrameProperties::from_xml_element(child_node)?),
                "xfrm" => transform = Some(Transform2D::from_xml_element(child_node)?),
                "graphic" => graphic = Some(GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_drawing_props =
            non_visual_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvFrPr"))?;
        let transform = transform.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "xfrm"))?;
        let graphic = graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?;

        Ok(Self {
            non_visual_drawing_props,
            non_visual_props,
            transform,
            graphic,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WordprocessingContentPartNonVisual {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub non_visual_props: Option<NonVisualContentPartProperties>,
}

impl WordprocessingContentPartNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "cNvPr" => {
                        instance.non_visual_drawing_props = Some(NonVisualDrawingProps::from_xml_element(child_node)?)
                    }
                    "cNvContentPartPr" => {
                        instance.non_visual_props = Some(NonVisualContentPartProperties::from_xml_element(child_node)?)
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct WordprocessingContentPart {
    pub properties: Option<WordprocessingContentPartNonVisual>,
    pub transform: Option<Transform2D>,
    pub black_and_white_mode: Option<BlackWhiteMode>,
    pub relationship_id: RelationshipId,
}

impl WordprocessingContentPart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut black_and_white_mode = None;
        let mut relationship_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "bwMode" => black_and_white_mode = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        let mut properties = None;
        let mut transform = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvContentPartPr" => {
                    properties = Some(WordprocessingContentPartNonVisual::from_xml_element(child_node)?)
                }
                "xfrm" => transform = Some(Transform2D::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            properties,
            transform,
            black_and_white_mode,
            relationship_id,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum WordprocessingShapeChoice {
    Shape(Box<WordprocessingShape>),
    Group(Box<WordprocessingGroup>),
    GraphicFrame(GraphicFrame),
    Picture(Box<Picture>),
    ContentPart(WordprocessingContentPart),
}

#[derive(Debug, Clone, PartialEq)]
pub struct WordprocessingGroup {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub non_visual_drawing_shape_props: NonVisualGroupDrawingShapeProps,
    pub group_shape_props: GroupShapeProperties,
    pub shapes: Vec<WordprocessingShapeChoice>,
}

impl WordprocessingGroup {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_drawing_props = None;
        let mut non_visual_drawing_shape_props = None;
        let mut group_shape_props = None;
        let mut shapes = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => non_visual_drawing_props = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvGrpSpPr" => {
                    non_visual_drawing_shape_props =
                        Some(NonVisualGroupDrawingShapeProps::from_xml_element(child_node)?)
                }
                "grpSpPr" => group_shape_props = Some(GroupShapeProperties::from_xml_element(child_node)?),
                "wsp" => shapes.push(WordprocessingShapeChoice::Shape(Box::new(
                    WordprocessingShape::from_xml_element(child_node)?,
                ))),
                "grpSp" => shapes.push(WordprocessingShapeChoice::Group(Box::new(
                    WordprocessingGroup::from_xml_element(child_node)?,
                ))),
                "graphicFrame" => shapes.push(WordprocessingShapeChoice::GraphicFrame(GraphicFrame::from_xml_element(
                    child_node,
                )?)),
                "pic" => shapes.push(WordprocessingShapeChoice::Picture(Box::new(Picture::from_xml_element(
                    child_node,
                )?))),
                "contentPart" => shapes.push(WordprocessingShapeChoice::ContentPart(
                    WordprocessingContentPart::from_xml_element(child_node)?,
                )),
                _ => (),
            }
        }

        let non_visual_drawing_shape_props = non_visual_drawing_shape_props
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGrpSpPr"))?;
        let group_shape_props =
            group_shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "grpSpPr"))?;

        Ok(Self {
            non_visual_drawing_props,
            non_visual_drawing_shape_props,
            group_shape_props,
            shapes,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WordprocessingCanvas {
    pub background_formatting: Option<BackgroundFormatting>,
    pub whole_formatting: Option<WholeE2oFormatting>,
    pub shapes: Vec<WordprocessingShapeChoice>,
}

impl WordprocessingCanvas {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "bg" => instance.background_formatting = Some(BackgroundFormatting::from_xml_element(child_node)?),
                    "whole" => instance.whole_formatting = Some(WholeE2oFormatting::from_xml_element(child_node)?),
                    "wsp" => instance.shapes.push(WordprocessingShapeChoice::Shape(Box::new(
                        WordprocessingShape::from_xml_element(child_node)?,
                    ))),
                    "pic" => {
                        instance
                            .shapes
                            .push(WordprocessingShapeChoice::Picture(Box::new(Picture::from_xml_element(
                                child_node,
                            )?)))
                    }
                    "contentPart" => instance.shapes.push(WordprocessingShapeChoice::ContentPart(
                        WordprocessingContentPart::from_xml_element(child_node)?,
                    )),
                    "wgp" => instance.shapes.push(WordprocessingShapeChoice::Group(Box::new(
                        WordprocessingGroup::from_xml_element(child_node)?,
                    ))),
                    "graphicFrame" => {
                        instance
                            .shapes
                            .push(WordprocessingShapeChoice::GraphicFrame(GraphicFrame::from_xml_element(
                                child_node,
                            )?))
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
    use crate::shared::drawingml::core::{GraphicalObjectData, Hyperlink, Locking, ShapeLocking};
    use std::str::FromStr;

    const TEST_LOCKING_ATTRIBUTES: &'static str = r#"noGrp="false" noSelect="false" noRot="false" noChangeAspect="false"
        noMove="false" noResize="false" noEditPoints="false" noAdjustHandles="false" noChangeArrowheads="false" noChangeShapeType="false""#;

    fn test_locking_instance() -> Locking {
        Locking {
            no_grouping: Some(false),
            no_select: Some(false),
            no_rotate: Some(false),
            no_change_aspect_ratio: Some(false),
            no_move: Some(false),
            no_resize: Some(false),
            no_edit_points: Some(false),
            no_adjust_handles: Some(false),
            no_change_arrowheads: Some(false),
            no_change_shape_type: Some(false),
        }
    }

    fn test_shape_locking_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} {} noTextEdit="false"></{node_name}>"#,
            TEST_LOCKING_ATTRIBUTES,
            node_name = node_name
        )
    }

    fn test_shape_locking_instance() -> ShapeLocking {
        ShapeLocking {
            locking: test_locking_instance(),
            no_text_edit: Some(false),
        }
    }

    fn test_non_visual_drawing_shape_props_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} txBox="true">
            {}
        </{node_name}>"#,
            test_shape_locking_xml("spLocks"),
            node_name = node_name,
        )
    }

    fn test_non_visual_drawing_shape_props_instance() -> NonVisualDrawingShapeProps {
        NonVisualDrawingShapeProps {
            shape_locks: Some(test_shape_locking_instance()),
            is_text_box: Some(true),
        }
    }

    fn test_graphical_object_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}><graphicData uri="http://some/url" /></{node_name}>"#,
            node_name = node_name
        )
    }

    fn test_graphical_object_instance() -> GraphicalObject {
        GraphicalObject {
            graphic_data: GraphicalObjectData {
                uri: String::from("http://some/url"),
            },
        }
    }

    impl EffectExtent {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} l="0" t="0" r="100" b="100"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                left: 0,
                top: 0,
                right: 100,
                bottom: 100,
            }
        }
    }

    #[test]
    pub fn test_effect_extent_from_xml() {
        let xml = EffectExtent::test_xml("effectExtent");
        let effect_extent = EffectExtent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(effect_extent, EffectExtent::test_instance());
    }

    fn test_non_visual_drawing_props_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="1" name="Object name" descr="Some description" title="Title of the object">
            <a:hlinkClick r:id="rId2" tooltip="Some Sample Text"/>
            <a:hlinkHover r:id="rId2" tooltip="Some Sample Text"/>
        </{node_name}>"#,
            node_name = node_name,
        )
    }

    fn test_non_visual_drawing_props_instance() -> NonVisualDrawingProps {
        NonVisualDrawingProps {
            id: 1,
            name: String::from("Object name"),
            description: Some(String::from("Some description")),
            hidden: None,
            title: Some(String::from("Title of the object")),
            hyperlink_click: Some(Box::new(Hyperlink {
                relationship_id: Some(String::from("rId2")),
                tooltip: Some(String::from("Some Sample Text")),
                ..Default::default()
            })),
            hyperlink_hover: Some(Box::new(Hyperlink {
                relationship_id: Some(String::from("rId2")),
                tooltip: Some(String::from("Some Sample Text")),
                ..Default::default()
            })),
        }
    }

    impl Inline {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} distT="0" distB="100" distL="0" distR="100">
                <extent cx="10000" cy="10000" />
                {}
                {}
                {}
            </{node_name}>"#,
                EffectExtent::test_xml("effectExtent"),
                test_non_visual_drawing_props_xml("docPr"),
                test_graphical_object_xml("a:graphic"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                extent: PositiveSize2D::new(10000, 10000),
                effect_extent: Some(EffectExtent::test_instance()),
                doc_properties: test_non_visual_drawing_props_instance(),
                graphic_frame_properties: None,
                graphic: test_graphical_object_instance(),
                distance_top: Some(0),
                distance_bottom: Some(100),
                distance_left: Some(0),
                distance_right: Some(100),
            }
        }
    }

    #[test]
    pub fn test_inline_from_xml_element() {
        let xml = Inline::test_xml("inline");
        assert_eq!(
            Inline::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            Inline::test_instance()
        );
    }

    impl WrapPath {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} edited="true">
            <start x="0" y="0" />
            <lineTo x="50" y="50" />
            <lineTo x="100" y="100" />
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start: Point2D::new(0, 0),
                line_to: vec![Point2D::new(50, 50), Point2D::new(100, 100)],
                edited: Some(true),
            }
        }
    }

    #[test]
    pub fn test_wrap_path_from_xml() {
        let xml = WrapPath::test_xml("wrapPath");

        let wrap_path = WrapPath::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_path, WrapPath::test_instance());
    }

    impl WrapSquare {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} wrapText="bothSides" distT="0" distB="100" distL="0" distR="100">
            {}
        </{node_name}>"#,
                EffectExtent::test_xml("effectExtent"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            WrapSquare {
                effect_extent: Some(EffectExtent::test_instance()),
                wrap_text: WrapText::BothSides,
                distance_top: Some(0),
                distance_bottom: Some(100),
                distance_left: Some(0),
                distance_right: Some(100),
            }
        }
    }

    #[test]
    pub fn test_wrap_square_from_xml() {
        let xml = WrapSquare::test_xml("wrapSquare");
        let wrap_square = WrapSquare::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();

        assert_eq!(wrap_square, WrapSquare::test_instance());
    }

    impl WrapTight {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} wrapText="bothSides", distL="0" distR="0">
            {}
        </{node_name}>"#,
                WrapPath::test_xml("wrapPolygon"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                wrap_polygon: WrapPath::test_instance(),
                wrap_text: WrapText::BothSides,
                distance_left: Some(0),
                distance_right: Some(0),
            }
        }
    }

    #[test]
    pub fn test_wrap_tight_from_xml() {
        let xml = WrapTight::test_xml("wrapTight");
        let wrap_tight = WrapTight::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_tight, WrapTight::test_instance());
    }

    impl WrapThrough {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} wrapText="bothSides" distL="0" distR="0">
            {}
        </{node_name}>"#,
                WrapPath::test_xml("wrapPolygon"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                wrap_polygon: WrapPath::test_instance(),
                wrap_text: WrapText::BothSides,
                distance_left: Some(0),
                distance_right: Some(0),
            }
        }
    }

    #[test]
    pub fn test_wrap_through_from_xml() {
        let xml = WrapThrough::test_xml("wrapThrough");
        let wrap_through = WrapThrough::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_through, WrapThrough::test_instance());
    }

    impl WrapTopBottom {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} distT="0" distB="0">
            {}
        </{node_name}>"#,
                EffectExtent::test_xml("effectExtent"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                effect_extent: Some(EffectExtent::test_instance()),
                distance_top: Some(0),
                distance_bottom: Some(0),
            }
        }
    }

    #[test]
    pub fn test_wrap_top_bottom_from_xml() {
        let xml = WrapTopBottom::test_xml("wrapTopAndBottom");
        let wrap_top_bottom = WrapTopBottom::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_top_bottom, WrapTopBottom::test_instance());
    }

    #[test]
    pub fn test_wrap_type_none_from_xml() {
        let xml = r#"<wrapNone></wrapNone>"#;
        let wrap_type = WrapType::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(wrap_type, WrapType::None);
    }

    #[test]
    pub fn test_wrap_type_square() {
        let xml = WrapSquare::test_xml("wrapSquare");
        let wrap_type = WrapType::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_type, WrapType::Square(WrapSquare::test_instance()));
    }

    #[test]
    pub fn test_wrap_type_tight() {
        let xml = WrapTight::test_xml("wrapTight");
        let wrap_type = WrapType::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_type, WrapType::Tight(WrapTight::test_instance()));
    }

    #[test]
    pub fn test_wrap_type_through() {
        let xml = WrapThrough::test_xml("wrapThrough");
        let wrap_type = WrapType::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_type, WrapType::Through(WrapThrough::test_instance()));
    }

    #[test]
    pub fn test_wrap_type_top_and_bottom() {
        let xml = WrapTopBottom::test_xml("wrapTopAndBottom");
        let wrap_type = WrapType::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(wrap_type, WrapType::TopAndBottom(WrapTopBottom::test_instance()));
    }

    impl PosH {
        pub fn test_xml_with_align(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} relativeFrom="margin">
            <align>left</align>
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_xml_with_offset(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} relativeFrom="margin">
            <posOffset>50</posOffset>
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance_with_align() -> Self {
            Self {
                align_or_offset: PosHChoice::Align(AlignH::Left),
                relative_from: RelFromH::Margin,
            }
        }

        pub fn test_instance_with_offset() -> Self {
            Self {
                align_or_offset: PosHChoice::PositionOffset(50),
                relative_from: RelFromH::Margin,
            }
        }
    }

    #[test]
    pub fn test_pos_h_with_align_from_xml() {
        let xml = PosH::test_xml_with_align("posH");
        let pos_h = PosH::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pos_h, PosH::test_instance_with_align());
    }

    #[test]
    pub fn test_pos_h_with_offset_from_xml() {
        let xml = PosH::test_xml_with_offset("posH");
        let pos_h = PosH::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pos_h, PosH::test_instance_with_offset());
    }

    impl PosV {
        pub fn test_xml_with_align(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} relativeFrom="margin">
            <align>top</align>
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_xml_with_offset(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} relativeFrom="margin">
            <posOffset>50</posOffset>
        </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance_with_align() -> Self {
            Self {
                align_or_offset: PosVChoice::Align(AlignV::Top),
                relative_from: RelFromV::Margin,
            }
        }

        pub fn test_instance_with_offset() -> Self {
            Self {
                align_or_offset: PosVChoice::PositionOffset(50),
                relative_from: RelFromV::Margin,
            }
        }
    }

    #[test]
    pub fn test_pos_v_with_align_from_xml() {
        let xml = PosV::test_xml_with_align("posV");
        let pos_v = PosV::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pos_v, PosV::test_instance_with_align());
    }

    #[test]
    pub fn test_pos_v_with_offset_from_xml() {
        let xml = PosV::test_xml_with_offset("posV");
        let pos_h = PosV::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(pos_h, PosV::test_instance_with_offset());
    }

    impl Anchor {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} distT="0" distB="100" distL="0" distR="100" simplePos="false" relativeHeight="100" behindDoc="false" locked="false" layoutInCell="false" hidden="false" allowOverlap="false">
            <simplePos x="0" y="0" />
            {}
            {}
            <extent cx="100" cy="100" />
            {}
            {}
            {}
            {}
        </{node_name}>"#,
            PosH::test_xml_with_align("positionH"),
            PosV::test_xml_with_align("positionV"),
            EffectExtent::test_xml("effectExtent"),
            WrapSquare::test_xml("wrapSquare"),
            test_non_visual_drawing_props_xml("docPr"),
            test_graphical_object_xml("a:graphic"),
            node_name=node_name,
        )
        }

        pub fn test_instance() -> Self {
            Self {
                simple_position: Point2D::new(0, 0),
                horizontal_position: PosH::test_instance_with_align(),
                vertical_position: PosV::test_instance_with_align(),
                extent: PositiveSize2D::new(100, 100),
                effect_extent: Some(EffectExtent::test_instance()),
                wrap_type: WrapType::Square(WrapSquare::test_instance()),
                document_properties: test_non_visual_drawing_props_instance(),
                graphic_frame_properties: None,
                graphic: test_graphical_object_instance(),
                distance_top: Some(0),
                distance_bottom: Some(100),
                distance_left: Some(0),
                distance_right: Some(100),
                use_simple_position: Some(false),
                relative_height: 100,
                behind_document_text: false,
                locked: false,
                layout_in_cell: false,
                hidden: Some(false),
                allow_overlap: false,
            }
        }
    }

    #[test]
    pub fn test_anchor_from_xml() {
        let xml = Anchor::test_xml("anchor");
        let anchor = Anchor::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap();
        assert_eq!(anchor, Anchor::test_instance());
    }

    impl TxbxContent {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                super::super::document::P::test_xml("p"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            use super::super::document::{BlockLevelElts, ContentBlockContent, P};

            Self {
                block_level_elements: vec![BlockLevelElts::Chunk(ContentBlockContent::Paragraph(Box::new(
                    P::test_instance(),
                )))],
            }
        }
    }

    #[test]
    pub fn test_txbx_content_from_xml() {
        let xml = TxbxContent::test_xml("txbxContent");
        assert_eq!(
            TxbxContent::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TxbxContent::test_instance()
        );
    }

    impl TextboxInfo {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} id="1">
                {}
            </{node_name}>"#,
                TxbxContent::test_xml("txbxContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                textbox_content: TxbxContent::test_instance(),
                id: Some(1),
            }
        }
    }

    #[test]
    pub fn test_textbox_info_from_xml() {
        let xml = TextboxInfo::test_xml("textboxInfo");
        assert_eq!(
            TextboxInfo::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            TextboxInfo::test_instance(),
        );
    }

    impl LinkedTextboxInformation {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} id="1" seq="1"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { id: 1, sequence: 1 }
        }
    }

    #[test]
    pub fn test_linked_textbox_information_from_xml() {
        let xml = LinkedTextboxInformation::test_xml("linkedTextboxInformation");
        assert_eq!(
            LinkedTextboxInformation::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            LinkedTextboxInformation::test_instance(),
        );
    }

    impl WordprocessingShape {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} normalEastAsianFlow="false">
                {}
                <spPr />
                {}
                <bodyPr />
            </{node_name}>"#,
                test_non_visual_drawing_shape_props_xml("cNvSpPr"),
                TextboxInfo::test_xml("txbx"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                non_visual_drawing_props: None,
                properties: WordprocessingShapePropertiesChoice::ShapeProperties(
                    test_non_visual_drawing_shape_props_instance(),
                ),
                shape_properties: Default::default(),
                style: None,
                text_box_info: Some(WordprocessingShapeTextboxInfoChoice::Textbox(
                    TextboxInfo::test_instance(),
                )),
                text_body_properties: Default::default(),
                normal_east_asian_flow: Some(false),
            }
        }
    }

    #[test]
    pub fn test_wordprocessing_shape_from_xml() {
        let xml = WordprocessingShape::test_xml("wordprocessingShape");
        assert_eq!(
            WordprocessingShape::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WordprocessingShape::test_instance(),
        );
    }

    impl GraphicFrame {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                <cNvFrPr />
                <xfrm />
                {}
            </{node_name}>"#,
                test_non_visual_drawing_props_xml("cNvPr"),
                test_graphical_object_xml("a:graphic"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                non_visual_drawing_props: test_non_visual_drawing_props_instance(),
                non_visual_props: Default::default(),
                transform: Default::default(),
                graphic: test_graphical_object_instance(),
            }
        }
    }

    #[test]
    pub fn test_graphic_frame_from_xml() {
        let xml = GraphicFrame::test_xml("graphicFrame");
        assert_eq!(
            GraphicFrame::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            GraphicFrame::test_instance(),
        );
    }

    impl WordprocessingContentPart {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} bwMode="auto" r:id="rId1">
                <nvContentPartPr />
                <xfrm />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(Default::default()),
                transform: Some(Default::default()),
                black_and_white_mode: Some(BlackWhiteMode::Auto),
                relationship_id: RelationshipId::from("rId1"),
            }
        }
    }

    #[test]
    pub fn test_wordprocessing_content_part_from_xml() {
        let xml = WordprocessingContentPart::test_xml("wordprocessingContentPart");
        assert_eq!(
            WordprocessingContentPart::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WordprocessingContentPart::test_instance(),
        );
    }

    impl WordprocessingGroup {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <cNvGrpSpPr />
                <grpSpPr />
                {}
            </{node_name}>"#,
                WordprocessingContentPart::test_xml("contentPart"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                non_visual_drawing_props: None,
                non_visual_drawing_shape_props: Default::default(),
                group_shape_props: Default::default(),
                shapes: vec![WordprocessingShapeChoice::ContentPart(
                    WordprocessingContentPart::test_instance(),
                )],
            }
        }
    }

    #[test]
    pub fn test_wordprocessing_group_from_xml() {
        let xml = WordprocessingGroup::test_xml("wordprocessingGroup");
        assert_eq!(
            WordprocessingGroup::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WordprocessingGroup::test_instance(),
        );
    }

    impl WordprocessingCanvas {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                WordprocessingContentPart::test_xml("contentPart"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                background_formatting: None,
                whole_formatting: None,
                shapes: vec![WordprocessingShapeChoice::ContentPart(
                    WordprocessingContentPart::test_instance(),
                )],
            }
        }
    }

    #[test]
    pub fn test_wordprocessing_canvas_from_xml() {
        let xml = WordprocessingCanvas::test_xml("wordprocessingCanvas");
        assert_eq!(
            WordprocessingCanvas::from_xml_element(&XmlNode::from_str(xml.as_str()).unwrap()).unwrap(),
            WordprocessingCanvas::test_instance(),
        );
    }
}
