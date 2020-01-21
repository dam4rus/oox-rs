use super::{
    audiovideo::EmbeddedWAVAudioFile,
    coordsys::{GroupTransform2D, Transform2D},
    shapedefs::Geometry,
    shapeprops::{
        EffectProperties, FillProperties, LineDashProperties, LineEndProperties, LineFillProperties, LineJoinProperties,
    },
    simpletypes::{
        AnimationChartBuildType, AnimationDgmBuildType, BlackWhiteMode, ChartBuildStep, CompoundLine, DgmBuildStep,
        DrawingElementId, Guid, LineCap, LineWidth, PenAlignment,
    },
    styles::{FontReference, StyleMatrixReference},
    text::{bodyformatting::TextBodyProperties, bullet::TextListStyle, paragraphs::TextParagraph},
};
use crate::{
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    shared::relationship::RelationshipId,
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, PartialEq)]
pub enum AnimationGraphicalObjectBuildProperties {
    /// This element specifies how to build the animation for a diagram.
    ///
    /// # Xml example
    ///
    /// Consider having a diagram appear as on entity as opposed to by section. The bldDgm element should
    /// be used as follows:
    /// ```xml
    /// <p:bdldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldDgm bld="one"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildDiagram(AnimationDgmBuildProperties),

    /// This element specifies how to build the animation for a diagram.
    ///
    /// # Xml example
    ///
    /// Consider the following example where a chart is specified to be animated by category rather than as
    /// one entity. Thus, the bldChart element should be used as follows:
    /// ```xml
    /// <p:bdldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildChart(AnimationChartBuildProperties),
}

impl XsdType for AnimationGraphicalObjectBuildProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldDgm" => Ok(AnimationGraphicalObjectBuildProperties::BuildDiagram(
                AnimationDgmBuildProperties::from_xml_element(xml_node)?,
            )),
            "bldChart" => Ok(AnimationGraphicalObjectBuildProperties::BuildChart(
                AnimationChartBuildProperties::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_AnimationGraphicalObjectBuildProperties",
            ))),
        }
    }
}

impl XsdChoice for AnimationGraphicalObjectBuildProperties {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "bldDgm" | "bldChart" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct AnimationDgmBuildProperties {
    /// Specifies how the chart is built. The animation animates the sub-elements in the
    /// container in the particular order defined by this attribute.
    ///
    /// Defaults to AnimationDgmBuildType::AllAtOnce
    pub build_type: Option<AnimationDgmBuildType>,

    /// Specifies whether the animation of the objects in this diagram should be reversed or not.
    /// If this attribute is not specified, a value of false is assumed.
    ///
    /// Defaults to false
    pub reverse: Option<bool>,
}

impl AnimationDgmBuildProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "bld" => instance.build_type = Some(value.parse()?),
                    "rev" => instance.reverse = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct AnimationChartBuildProperties {
    /// Specifies how the chart is built. The animation animates the sub-elements in the
    /// container in the particular order defined by this attribute.
    ///
    /// Defaults to AnimationChartBuildType::AllAtOnce
    pub build_type: Option<AnimationChartBuildType>,

    /// Specifies whether or not the chart background elements should be animated as well.
    ///
    /// Defaults to true
    ///
    /// # Note
    ///
    /// An example of background elements are grid lines and the chart legend.
    pub animate_bg: Option<bool>,
}

impl AnimationChartBuildProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "bld" => instance.build_type = Some(value.parse()?),
                    "animBg" => instance.animate_bg = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum AnimationElementChoice {
    /// This element specifies a reference to a diagram that should be animated within a sequence of slide animations.
    /// In addition to simply acting as a reference to a diagram there is also animation build steps defined.
    Diagram(AnimationDgmElement),

    /// This element specifies a reference to a chart that should be animated within a sequence of slide animations. In
    /// addition to simply acting as a reference to a chart there is also animation build steps defined.
    Chart(AnimationChartElement),
}

impl XsdType for AnimationElementChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "dgm" => Ok(AnimationElementChoice::Diagram(AnimationDgmElement::from_xml_element(
                xml_node,
            )?)),
            "chart" => Ok(AnimationElementChoice::Chart(AnimationChartElement::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_AnimationElementChoice",
            ))),
        }
    }
}

impl XsdChoice for AnimationElementChoice {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "dgm" | "chart" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct AnimationDgmElement {
    /// Specifies the GUID of the shape for this build step in the animation.
    ///
    /// Defaults to {00000000-0000-0000-0000-000000000000}
    pub id: Option<Guid>,

    /// Specifies which step this part of the diagram should be built using. For instance the
    /// diagram can be built as one object meaning it is animated as a single graphic.
    /// Alternatively the diagram can be animated, or built as separate pieces.
    ///
    /// Defaults to DgmBuildStep::Shape
    pub build_step: Option<DgmBuildStep>,
}

impl AnimationDgmElement {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "id" => instance.id = Some(value.clone()),
                    "bldStep" => instance.build_step = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AnimationChartElement {
    /// Specifies the index of the series within the corresponding chart that should be animated.
    ///
    /// Defaults to -1
    pub series_index: Option<i32>,

    /// Specifies the index of the category within the corresponding chart that should be
    /// animated.
    ///
    /// Defaults to -1
    pub category_index: Option<i32>,

    /// Specifies which step this part of the chart should be built using. For instance the chart can
    /// be built as one object meaning it is animated as a single graphic. Alternatively the chart
    /// can be animated, or built as separate pieces.
    pub build_step: ChartBuildStep,
}

impl AnimationChartElement {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut series_index = None;
        let mut category_index = None;
        let mut build_step = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "seriesIdx" => series_index = Some(value.parse()?),
                "categoryIdx" => category_index = Some(value.parse()?),
                "bldStep" => build_step = Some(value.parse()?),
                _ => (),
            }
        }

        let build_step = build_step.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bldStep"))?;

        Ok(Self {
            series_index,
            category_index,
            build_step,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualConnectorProperties {
    /// This element specifies all locking properties for a connection shape. These properties inform the generating
    /// application about specific properties that have been previously locked and thus should not be changed.
    pub connector_locks: Option<ConnectorLocking>,

    /// This element specifies the starting connection that should be made by the corresponding connector shape. This
    /// connects the head of the connector to the first shape.
    pub start_connection: Option<Connection>,

    /// This element specifies the ending connection that should be made by the corresponding connector shape. This
    /// connects the end tail of the connector to the final destination shape.
    pub end_connection: Option<Connection>,
}

impl NonVisualConnectorProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "cxnSpLocks" => instance.connector_locks = Some(ConnectorLocking::from_xml_element(child_node)?),
                    "stCxn" => instance.start_connection = Some(Connection::from_xml_element(child_node)?),
                    "endCxn" => instance.end_connection = Some(Connection::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualGraphicFrameProperties {
    /// This element specifies all locking properties for a graphic frame. These properties inform the generating
    /// application about specific properties that have been previously locked and thus should not be changed.
    pub graphic_frame_locks: Option<GraphicalObjectFrameLocking>,
}

impl NonVisualGraphicFrameProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let graphic_frame_locks = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "graphicFrameLocks")
            .map(GraphicalObjectFrameLocking::from_xml_element)
            .transpose()?;

        Ok(Self { graphic_frame_locks })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ContentPartLocking {
    pub locking: Locking,
}

impl ContentPartLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let locking = Locking::from_xml_element(xml_node)?;
        Ok(Self { locking })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualContentPartProperties {
    pub locking: Option<ContentPartLocking>,
    pub is_comment: Option<bool>, // default=true
}

impl NonVisualContentPartProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_comment = xml_node.attributes.get("isComment").map(parse_xml_bool).transpose()?;

        let locking = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cpLocks")
            .map(ContentPartLocking::from_xml_element)
            .transpose()?;

        Ok(Self { locking, is_comment })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualGroupDrawingShapeProps {
    pub locks: Option<GroupLocking>,
}

impl NonVisualGroupDrawingShapeProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let locks = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "grpSpLocks")
            .map(GroupLocking::from_xml_element)
            .transpose()?;

        Ok(Self { locks })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualPictureProperties {
    /// Specifies if the user interface should show the resizing of the picture based on the
    /// picture's current size or its original size. If this attribute is set to true, then scaling is
    /// relative to the original picture size as opposed to the current picture size.
    ///
    /// Defaults to true
    ///
    /// # Example
    ///
    /// Consider the case where a picture has been resized within a document and is
    /// now 50% of the originally inserted picture size. Now if the user chooses to make a later
    /// adjustment to the size of this picture within the generating application, then the value of
    /// this attribute should be checked.
    ///
    /// If this attribute is set to true then a value of 50% is shown. Similarly, if this attribute is set
    /// to false, then a value of 100% should be shown because the picture has not yet been
    /// resized from its current (smaller) size.
    pub prefer_relative_resize: Option<bool>,
    pub picture_locks: Option<PictureLocking>,
}

impl NonVisualPictureProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let prefer_relative_resize = xml_node
            .attributes
            .get("preferRelativeResize")
            .map(parse_xml_bool)
            .transpose()?;

        let picture_locks = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "picLocks")
            .map(PictureLocking::from_xml_element)
            .transpose()?;

        Ok(Self {
            prefer_relative_resize,
            picture_locks,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct NonVisualDrawingShapeProps {
    pub shape_locks: Option<ShapeLocking>,

    /// Specifies that the corresponding shape is a text box and thus should be treated as such
    /// by the generating application. If this attribute is omitted then it is assumed that the
    /// corresponding shape is not specifically a text box.
    ///
    /// Defaults to false
    pub is_text_box: Option<bool>,
}

impl NonVisualDrawingShapeProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_text_box = xml_node.attributes.get("txBox").map(parse_xml_bool).transpose()?;

        let shape_locks = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "spLocks")
            .map(ShapeLocking::from_xml_element)
            .transpose()?;

        Ok(Self {
            is_text_box,
            shape_locks,
        })
    }
}

/// ```xml example
/// <docPr id="1" name="Object name" descr="Some description" title="Title of the object">
///     <a:hlinkClick r:id="rId2" tooltip="Some Sample Text"/>
///     <a:hlinkHover r:id="rId2" tooltip="Some Sample Text"/>
/// </docPr>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct NonVisualDrawingProps {
    /// Specifies a unique identifier for the current DrawingML object within the current
    /// document. This ID can be used to assist in uniquely identifying this object so that it can
    /// be referred to by other parts of the document.
    ///
    /// If multiple objects within the same document share the same id attribute value, then the
    /// document shall be considered non-conformant.
    ///
    /// # Example
    ///
    /// Consider a DrawingML object defined as follows:
    ///
    /// <… id="10" … >
    ///
    /// The id attribute has a value of 10, which is the unique identifier for this DrawingML
    /// object.
    pub id: DrawingElementId,

    /// Specifies the name of the object.
    ///
    /// # Note
    ///
    /// Typically, this is used to store the original file name of a picture object.
    ///
    /// # Example
    ///
    /// Consider a DrawingML object defined as follows:
    ///
    /// < … name="foo.jpg" >
    ///
    /// The name attribute has a value of foo.jpg, which is the name of this DrawingML object.
    pub name: String,

    /// Specifies alternative text for the current DrawingML object, for use by assistive
    /// technologies or applications which do not display the current object.
    ///
    /// If this element is omitted, then no alternative text is present for the parent object.
    ///
    /// # Example
    ///
    /// Consider a DrawingML object defined as follows:
    ///
    /// <… descr="A picture of a bowl of fruit">
    ///
    /// The descr attribute contains alternative text which can be used in place of the actual
    /// DrawingML object.
    pub description: Option<String>,

    /// Specifies whether this DrawingML object is displayed. When a DrawingML object is
    /// displayed within a document, that object can be hidden (i.e., present, but not visible).
    /// This attribute determines whether the object is rendered or made hidden. [Note: An
    /// application can have settings which allow this object to be viewed. end note]
    ///
    /// If this attribute is omitted, then the parent DrawingML object shall be displayed (i.e., not
    /// hidden).
    ///
    /// Defaults to false
    ///
    /// # Example
    ///
    /// Consider an inline DrawingML object which must be hidden within the
    /// document's content. This setting would be specified as follows:
    ///
    /// <… hidden="true" />
    ///
    /// The hidden attribute has a value of true, which specifies that the DrawingML object is
    /// hidden and not displayed when the document is displayed.
    pub hidden: Option<bool>,

    /// Specifies the title (caption) of the current DrawingML object.
    ///
    /// If this attribute is omitted, then no title text is present for the parent object.
    ///
    /// # Example
    ///
    /// Consider a DrawingML object defined as follows:
    ///
    /// <… title="Process Flow Diagram">
    pub title: Option<String>,

    /// Specifies the hyperlink information to be activated when the user click's over the corresponding object.
    pub hyperlink_click: Option<Box<Hyperlink>>,

    /// This element specifies the hyperlink information to be activated when the user's mouse is hovered over the
    /// corresponding object. The operation of the hyperlink is to have the specified action be activated when the
    /// mouse of the user hovers over the object. When this action is activated then additional attributes can be used to
    /// specify other tasks that should be performed along with the action.
    pub hyperlink_hover: Option<Box<Hyperlink>>,
}

impl NonVisualDrawingProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_id = None;
        let mut opt_name = None;
        let mut description = None;
        let mut hidden = None;
        let mut title = None;
        let mut hyperlink_click = None;
        let mut hyperlink_hover = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => opt_id = Some(value.parse()?),
                "name" => opt_name = Some(value.clone()),
                "descr" => description = Some(value.clone()),
                "hidden" => hidden = Some(parse_xml_bool(value)?),
                "title" => title = Some(value.clone()),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "hlinkClick" => hyperlink_click = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                "hlinkHover" => hyperlink_hover = Some(Box::new(Hyperlink::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let id = opt_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let name = opt_name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;

        Ok(Self {
            id,
            name,
            description,
            hidden,
            title,
            hyperlink_click,
            hyperlink_hover,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Locking {
    /// Specifies that the generating application should not allow shape grouping for the
    /// corresponding connection shape. That is it cannot be combined within other shapes to
    /// form a group of shapes. If this attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_grouping: Option<bool>,

    /// Specifies that the generating application should not allow selecting of the corresponding
    /// connection shape. That means also that no picture, shapes or text attached to this
    /// connection shape can be selected if this attribute has been specified. If this attribute is
    /// not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_select: Option<bool>,

    /// Specifies that the generating application should not allow shape rotation changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_rotate: Option<bool>,

    /// Specifies that the generating application should not allow aspect ratio changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_change_aspect_ratio: Option<bool>,

    /// Specifies that the generating application should not allow position changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_move: Option<bool>,

    /// Specifies that the generating application should not allow size changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_resize: Option<bool>,

    /// Specifies that the generating application should not allow shape point changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_edit_points: Option<bool>,

    /// Specifies that the generating application should not show adjust handles for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_adjust_handles: Option<bool>,

    /// Specifies that the generating application should not allow arrowhead changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_change_arrowheads: Option<bool>,

    /// Specifies that the generating application should not allow shape type changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_change_shape_type: Option<bool>,
}

impl Locking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), Self::try_update_from_xml_attribute)
    }

    pub fn try_update_from_xml_attribute(mut self, (attr, value): (&String, &String)) -> Result<Self> {
        match attr.as_ref() {
            "noGrp" => self.no_grouping = Some(parse_xml_bool(value)?),
            "noSelect" => self.no_select = Some(parse_xml_bool(value)?),
            "noRot" => self.no_rotate = Some(parse_xml_bool(value)?),
            "noChangeAspect" => self.no_change_aspect_ratio = Some(parse_xml_bool(value)?),
            "noMove" => self.no_move = Some(parse_xml_bool(value)?),
            "noResize" => self.no_resize = Some(parse_xml_bool(value)?),
            "noEditPoints" => self.no_edit_points = Some(parse_xml_bool(value)?),
            "noAdjustHandles" => self.no_adjust_handles = Some(parse_xml_bool(value)?),
            "noChangeArrowheads" => self.no_change_arrowheads = Some(parse_xml_bool(value)?),
            "noChangeShapeType" => self.no_change_shape_type = Some(parse_xml_bool(value)?),
            _ => (),
        }

        Ok(self)
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct ShapeLocking {
    pub locking: Locking,

    /// Specifies that the generating application should not allow editing of the shape text for
    /// the corresponding shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_text_edit: Option<bool>,
}

impl ShapeLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "noTextEdit" => instance.no_text_edit = Some(parse_xml_bool(value)?),
                    _ => instance.locking = instance.locking.try_update_from_xml_attribute((attr, value))?,
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct GroupLocking {
    /// Specifies that the corresponding group shape cannot be grouped. That is it cannot be
    /// combined within other shapes to form a group of shapes. If this attribute is not specified,
    /// then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_grouping: Option<bool>,

    /// Specifies that the generating application should not show adjust handles for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_ungrouping: Option<bool>,

    /// Specifies that the corresponding group shape cannot have any part of it be selected. That
    /// means that no picture, shapes or attached text can be selected either if this attribute has
    /// been specified. If this attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_select: Option<bool>,

    /// Specifies that the corresponding group shape cannot be rotated Objects that reside
    /// within the group can still be rotated unless they also have been locked. If this attribute is
    /// not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_rotate: Option<bool>,

    /// Specifies that the generating application should not allow aspect ratio changes for the
    /// corresponding connection shape. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_change_aspect_ratio: Option<bool>,

    /// Specifies that the corresponding graphic frame cannot be moved. Objects that reside
    /// within the graphic frame can still be moved unless they also have been locked. If this
    /// attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_move: Option<bool>,

    /// Specifies that the corresponding group shape cannot be resized. If this attribute is not
    /// specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_resize: Option<bool>,
}

impl GroupLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "noGrp" => instance.no_grouping = Some(parse_xml_bool(value)?),
                    "noUngrp" => instance.no_ungrouping = Some(parse_xml_bool(value)?),
                    "noSelect" => instance.no_select = Some(parse_xml_bool(value)?),
                    "noRot" => instance.no_rotate = Some(parse_xml_bool(value)?),
                    "noChangeAspect" => instance.no_change_aspect_ratio = Some(parse_xml_bool(value)?),
                    "noMove" => instance.no_move = Some(parse_xml_bool(value)?),
                    "noResize" => instance.no_resize = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct GraphicalObjectFrameLocking {
    /// Specifies that the generating application should not allow shape grouping for the
    /// corresponding graphic frame. That is it cannot be combined within other shapes to form
    /// a group of shapes. If this attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_grouping: Option<bool>,

    /// Specifies that the generating application should not allow selecting of objects within the
    /// corresponding graphic frame but allow selecting of the graphic frame itself. If this
    /// attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_drilldown: Option<bool>,

    /// Specifies that the generating application should not allow selecting of the corresponding
    /// picture. That means also that no picture, shapes or text attached to this picture can be
    /// selected if this attribute has been specified. If this attribute is not specified, then a value
    /// of false is assumed.
    ///
    /// Defaults to false
    ///
    /// # Note
    ///
    /// If this attribute is specified to be true then the graphic frame cannot be selected
    /// and the objects within the graphic frame cannot be selected as well. That is the entire
    /// graphic frame including all sub-parts are considered un-selectable.
    pub no_select: Option<bool>,

    /// Specifies that the generating application should not allow aspect ratio changes for the
    /// corresponding graphic frame. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_change_aspect: Option<bool>,

    /// Specifies that the corresponding graphic frame cannot be moved. Objects that reside
    /// within the graphic frame can still be moved unless they also have been locked. If this
    /// attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_move: Option<bool>,

    /// Specifies that the generating application should not allow size changes for the
    /// corresponding graphic frame. If this attribute is not specified, then a value of false is
    /// assumed.
    ///
    /// Defaults to false
    pub no_resize: Option<bool>,
}

impl GraphicalObjectFrameLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "noGrp" => instance.no_grouping = Some(parse_xml_bool(value)?),
                    "noDrilldown" => instance.no_drilldown = Some(parse_xml_bool(value)?),
                    "noSelect" => instance.no_select = Some(parse_xml_bool(value)?),
                    "noChangeAspect" => instance.no_change_aspect = Some(parse_xml_bool(value)?),
                    "noMove" => instance.no_move = Some(parse_xml_bool(value)?),
                    "noResize" => instance.no_resize = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct ConnectorLocking {
    pub locking: Locking,
}

impl ConnectorLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let locking = Locking::from_xml_element(xml_node)?;
        Ok(Self { locking })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct PictureLocking {
    pub locking: Locking,

    /// Specifies that the generating application should not allow cropping for the corresponding
    /// picture. If this attribute is not specified, then a value of false is assumed.
    ///
    /// Defaults to false
    pub no_crop: Option<bool>,
}

impl PictureLocking {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "noCrop" => instance.no_crop = Some(parse_xml_bool(value)?),
                    _ => instance.locking = instance.locking.try_update_from_xml_attribute((attr, value))?,
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Connection {
    /// Specifies the id of the shape to make the final connection to.
    pub id: DrawingElementId,

    /// Specifies the index into the connection site table of the final connection shape. That is
    /// there are many connection sites on a shape and it shall be specified which connection
    /// site the corresponding connector shape should connect to.
    pub shape_index: u32,
}

impl Connection {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut shape_index = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "idx" => shape_index = Some(value.parse()?),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let shape_index = shape_index.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?;

        Ok(Self { id, shape_index })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GraphicalObject {
    /// This element specifies the reference to a graphic object within the document. This graphic object is provided
    /// entirely by the document authors who choose to persist this data within the document.
    ///
    /// # Note
    ///
    /// Depending on the kind of graphical object used not every generating application that supports the
    /// OOXML framework has the ability to render the graphical object.
    pub graphic_data: GraphicalObjectData,
}

impl GraphicalObject {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let graphic_data = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "graphicData")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "graphicData")))
            .and_then(GraphicalObjectData::from_xml_element)?;

        Ok(Self { graphic_data })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GraphicalObjectData {
    // TODO implement
    //pub graphic_object: Vec<Any>,
    /// Specifies the URI, or uniform resource identifier that represents the data stored under
    /// this tag. The URI is used to identify the correct 'server' that can process the contents of
    /// this tag.
    pub uri: String,
}

impl GraphicalObjectData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let uri = xml_node
            .attributes
            .get("uri")
            .ok_or_else(|| Box::<dyn Error>::from(MissingAttributeError::new(xml_node.name.clone(), "uri")))?
            .clone();

        Ok(Self { uri })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct GroupShapeProperties {
    /// Specifies that the group shape should be rendered using only black and white coloring.
    /// That is the coloring information for the group shape should be converted to either black
    /// or white when rendering the corresponding shapes.
    ///
    /// No gray is to be used in rendering this image, only stark black and stark white.
    ///
    /// # Note
    ///
    /// This does not mean that the group shapes themselves are stored with only black
    /// and white color information. This attribute instead sets the rendering mode that the
    /// shapes use when rendering.
    pub black_and_white_mode: Option<BlackWhiteMode>,

    /// This element is nearly identical to the representation of 2-D transforms for ordinary shapes. The only
    /// addition is a member to represent the Child offset and the Child extents.
    pub transform: Option<Box<GroupTransform2D>>,

    /// Specifies the fill properties for this group shape.
    pub fill_properties: Option<FillProperties>,

    /// Specifies the effect that should be applied to this group shape.
    pub effect_properties: Option<EffectProperties>,
    // TODO implement
    //pub scene_3d: Option<Scene3D>,
}

impl GroupShapeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = xml_node
            .attributes
            .get("bwMode")
            .map(|value| value.parse())
            .transpose()?;

        xml_node.child_nodes.iter().try_fold(
            Self {
                black_and_white_mode,
                ..Default::default()
            },
            |mut instance, child_node| {
                match child_node.local_name() {
                    "xfrm" => instance.transform = Some(Box::new(GroupTransform2D::from_xml_element(child_node)?)),
                    child_name if FillProperties::is_choice_member(child_name) => {
                        instance.fill_properties = Some(FillProperties::from_xml_element(child_node)?)
                    }
                    child_name if EffectProperties::is_choice_member(child_name) => {
                        instance.effect_properties = Some(EffectProperties::from_xml_element(child_node)?)
                    }
                    _ => (),
                }

                Ok(instance)
            },
        )
    }
}

/// This element specifies an outline style that can be applied to a number of different objects such as shapes and
/// text. The line allows for the specifying of many different types of outlines including even line dashes and bevels.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct LineProperties {
    /// Specifies the width to be used for the underline stroke. If this attribute is omitted, then a
    /// value of 0 is assumed.
    pub width: Option<LineWidth>,

    /// Specifies the ending caps that should be used for this line. If this attribute is omitted, than a value of
    /// square is assumed.
    ///
    /// # Note
    ///
    /// Examples of cap types are rounded, flat, etc.
    pub cap: Option<LineCap>,

    /// Specifies the compound line type to be used for the underline stroke. If this attribute is
    /// omitted, then a value of CompoundLine::Single is assumed.
    pub compound: Option<CompoundLine>,

    /// Specifies the alignment to be used for the underline stroke.
    pub pen_alignment: Option<PenAlignment>,

    /// Specifies the fill properties for this line.
    pub fill_properties: Option<LineFillProperties>,

    /// Specifies the dash properties for this line.
    pub dash_properties: Option<LineDashProperties>,

    /// Specifies the join properties for this line.
    pub join_properties: Option<LineJoinProperties>,

    /// This element specifies decorations which can be added to the head of a line.
    pub head_end: Option<LineEndProperties>,

    /// This element specifies decorations which can be added to the tail of a line.
    pub tail_end: Option<LineEndProperties>,
}

impl LineProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineProperties> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w" => instance.width = Some(value.parse()?),
                    "cap" => instance.cap = Some(value.parse()?),
                    "cmpd" => instance.compound = Some(value.parse()?),
                    "algn" => instance.pen_alignment = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|instance| {
                xml_node
                    .child_nodes
                    .iter()
                    .try_fold(instance, |mut instance, child_node| {
                        match child_node.local_name() {
                            "headEnd" => instance.head_end = Some(LineEndProperties::from_xml_element(child_node)?),
                            "tailEnd" => instance.tail_end = Some(LineEndProperties::from_xml_element(child_node)?),
                            child_name if LineFillProperties::is_choice_member(child_name) => {
                                instance.fill_properties = Some(LineFillProperties::from_xml_element(child_node)?)
                            }
                            child_name if LineDashProperties::is_choice_member(child_name) => {
                                instance.dash_properties = Some(LineDashProperties::from_xml_element(child_node)?)
                            }
                            child_name if LineJoinProperties::is_choice_member(child_name) => {
                                instance.join_properties = Some(LineJoinProperties::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct ShapeProperties {
    /// Specifies that the picture should be rendered using only black and white coloring. That is
    /// the coloring information for the picture should be converted to either black or white
    /// when rendering the picture.
    ///
    /// No gray is to be used in rendering this image, only stark black and stark white.
    ///
    /// # Note
    ///
    /// This does not mean that the picture itself that is stored within the file is
    /// necessarily a black and white picture. This attribute instead sets the rendering mode that
    /// the picture has applied to when rendering.
    pub black_and_white_mode: Option<BlackWhiteMode>,

    /// This element represents 2-D transforms for ordinary shapes.
    pub transform: Option<Box<Transform2D>>,

    /// Specifies the geometry of this shape
    pub geometry: Option<Geometry>,

    /// Specifies the fill properties of this shape
    pub fill_properties: Option<FillProperties>,

    /// Specifies the outline properties of this shape.
    pub line_properties: Option<Box<LineProperties>>,

    /// Specifies the effect that should be applied to this shape.
    pub effect_properties: Option<EffectProperties>,
    // TODO implement
    //pub scene_3d: Option<Scene3D>,
    //pub shape_3d: Option<Shape3D>,
}

impl ShapeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = xml_node
            .attributes
            .get("bwMode")
            .map(|value| value.parse())
            .transpose()?;

        xml_node.child_nodes.iter().try_fold(
            Self {
                black_and_white_mode,
                ..Default::default()
            },
            |mut instance, child_node| {
                match child_node.local_name() {
                    "xfrm" => instance.transform = Some(Box::new(Transform2D::from_xml_element(child_node)?)),
                    "ln" => instance.line_properties = Some(Box::new(LineProperties::from_xml_element(child_node)?)),
                    child_name if Geometry::is_choice_member(child_name) => {
                        instance.geometry = Some(Geometry::from_xml_element(child_node)?)
                    }
                    child_name if FillProperties::is_choice_member(child_name) => {
                        instance.fill_properties = Some(FillProperties::from_xml_element(child_node)?)
                    }
                    child_name if EffectProperties::is_choice_member(child_name) => {
                        instance.effect_properties = Some(EffectProperties::from_xml_element(child_node)?)
                    }
                    _ => (),
                }

                Ok(instance)
            },
        )
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ShapeStyle {
    /// This element represents a reference to a line properties.
    pub line_reference: StyleMatrixReference,

    /// This element represents a reference to a fill properties.
    pub fill_reference: StyleMatrixReference,

    /// This element represents a reference to an effect properties.
    pub effect_reference: StyleMatrixReference,

    /// This element represents a reference to a themed font. When used it specifies which themed font to use along
    /// with a choice of color.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <fontRef idx="minor">
    ///   <schemeClr val="tx1"/>
    /// </fontRef>
    /// ```
    pub font_reference: FontReference,
}

impl ShapeStyle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut line_reference = None;
        let mut fill_reference = None;
        let mut effect_reference = None;
        let mut font_reference = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "lnRef" => line_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "fillRef" => fill_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "effectRef" => effect_reference = Some(StyleMatrixReference::from_xml_element(child_node)?),
                "fontRef" => font_reference = Some(FontReference::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let line_reference =
            line_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lnRef"))?;
        let fill_reference =
            fill_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fillRef"))?;
        let effect_reference =
            effect_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "effectRef"))?;
        let font_reference =
            font_reference.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fontRef"))?;

        Ok(Self {
            line_reference,
            fill_reference,
            effect_reference,
            font_reference,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextBody {
    /// Specifies the properties of this text body.
    pub body_properties: Box<TextBodyProperties>,

    /// Specifies the list style of this text body.
    pub list_style: Option<Box<TextListStyle>>,

    /// This element specifies the presence of a paragraph of text within the containing text body. The paragraph is the
    /// highest level text separation mechanism within a text body. A paragraph can contain text paragraph properties
    /// associated with the paragraph. If no properties are listed then properties specified in the defPPr element are
    /// used.
    ///
    /// # Xml example
    ///
    /// Consider the case where the user would like to describe a text body that contains two paragraphs.
    /// The requirement for these paragraphs is that one be right aligned and the other left aligned. The following
    /// DrawingML would specify a text body such as this.
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr algn="r">
    ///     </a:pPr>
    ///     …
    ///     <a:t>Some text</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr algn="l">
    ///     </a:pPr>
    ///     …
    ///     <a:t>Some text</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub paragraph_array: Vec<TextParagraph>,
}

impl TextBody {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut body_properties = None;
        let mut list_style = None;
        let mut paragraph_array = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bodyPr" => body_properties = Some(Box::new(TextBodyProperties::from_xml_element(child_node)?)),
                "lstStyle" => list_style = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "p" => paragraph_array.push(TextParagraph::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let body_properties =
            body_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bodyPr"))?;

        Ok(Self {
            body_properties,
            list_style,
            paragraph_array,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Hyperlink {
    /// Specifies the relationship id that when looked up in this slides relationship file contains
    /// the target of this hyperlink. This attribute cannot be omitted.
    pub relationship_id: Option<RelationshipId>,

    /// Specifies the URL when it has been determined by the generating application that the
    /// URL is invalid. That is the generating application can still store the URL but it is known
    /// that this URL is not correct.
    pub invalid_url: Option<String>,

    /// Specifies an action that is to be taken when this hyperlink is activated. This can be used to
    /// specify a slide to be navigated to or a script of code to be run.
    pub action: Option<String>,

    /// Specifies the target frame that is to be used when opening this hyperlink. When the
    /// hyperlink is activated this attribute is used to determine if a new window is launched for
    /// viewing or if an existing one can be used. If this attribute is omitted, than a new window
    /// is opened.
    pub target_frame: Option<String>,

    /// Specifies the tooltip that should be displayed when the hyperlink text is hovered over
    /// with the mouse. If this attribute is omitted, than the hyperlink text itself can be
    /// displayed.
    pub tooltip: Option<String>,

    /// Specifies whether to add this URI to the history when navigating to it. This allows for the
    /// viewing of this presentation without the storing of history information on the viewing
    /// machine. If this attribute is omitted, then a value of 1 or true is assumed.
    ///
    /// Defaults to true
    pub history: Option<bool>,

    /// Specifies if this attribute has already been used within this document. That is when a
    /// hyperlink has already been visited that this attribute would be utilized so the generating
    /// application can determine the color of this text. If this attribute is omitted, then a value
    /// of 0 or false is implied.
    ///
    /// Defaults to false
    pub highlight_click: Option<bool>,

    /// Specifies if the URL in question should stop all sounds that are playing when it is clicked.
    ///
    /// Defaults to false
    pub end_sound: Option<bool>,
    pub sound: Option<EmbeddedWAVAudioFile>,
}

impl Hyperlink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let sound = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "snd")
            .map(EmbeddedWAVAudioFile::from_xml_element)
            .transpose()?;

        xml_node.attributes.iter().try_fold(
            Self {
                sound,
                ..Default::default()
            },
            |mut instance, (attr, value)| {
                match attr.as_str() {
                    "r:id" => instance.relationship_id = Some(value.clone()),
                    "invalidUrl" => instance.invalid_url = Some(value.clone()),
                    "action" => instance.action = Some(value.clone()),
                    "tgtFrame" => instance.target_frame = Some(value.clone()),
                    "tooltip" => instance.tooltip = Some(value.clone()),
                    "history" => instance.history = Some(parse_xml_bool(value)?),
                    "highlightClick" => instance.highlight_click = Some(parse_xml_bool(value)?),
                    "endSnd" => instance.end_sound = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            },
        )
    }
}
