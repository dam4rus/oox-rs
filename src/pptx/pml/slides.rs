use crate::{
    shared::{
        drawingml::{
            audiovideo::{EmbeddedWAVAudioFile, Media},
            colors::ColorMappingOverride,
            coordsys::Transform2D,
            core::{
                GraphicalObject, GroupShapeProperties, NonVisualConnectorProperties, NonVisualDrawingProps,
                NonVisualDrawingShapeProps, NonVisualGraphicFrameProperties, NonVisualGroupDrawingShapeProps,
                NonVisualPictureProperties, ShapeProperties, ShapeStyle, TextBody,
            },
            shapeprops::{BlipFillProperties, EffectProperties, FillProperties},
            sharedstylesheet::ColorMapping,
            simpletypes::{BlackWhiteMode, PositiveCoordinate32, ShapeId},
            styles::StyleMatrixReference,
            text::bullet::TextListStyle,
        },
        relationship::RelationshipId,
    },
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::{error::Error, io::Read, str::FromStr};
use zip::read::ZipFile;

use super::{
    animation::{Build, TimeNodeGroup},
    presentation::{CustomerDataList, SlideLayoutIdList},
};

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

/// This simple type facilitates the storing of the content type a placeholder should contain.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PlaceholderType {
    /// Contains a slide title. Allowed for Slide, Slide Layout and Slide Master. Can be horizontal or vertical on Slide
    /// and Slide Layout.
    #[strum(serialize = "title")]
    Title,
    /// Contains body text. Allowed for Slide, Slide Layout, Slide Master, Notes, Notes Master. Can be horizontal
    /// or vertical on Slide and Slide Layout.
    #[strum(serialize = "body")]
    Body,
    /// Contains a title intended to be centered on the slide. Allowed for Slide and Slide Layout.
    #[strum(serialize = "ctrTitle")]
    CenteredTitle,
    /// Contains a subtitle. Allowed for Slide and Slide Layout.
    #[strum(serialize = "subTitle")]
    SubTitle,
    /// Contains the date and time. Allowed for Slide, Slide Layout, Slide Master, Notes, Notes Master, Handout Master
    #[strum(serialize = "dt")]
    DateTime,
    /// Contains the number of a slide. Allowed for Slide, Slide Layout, Slide Master, Notes, Notes Master, Handout
    /// Master
    #[strum(serialize = "sldNum")]
    SlideNumber,
    /// Contains text to be used as a footer in the document. Allowed for Slide, Slide Layout, Slide Master, Notes,
    /// Notes Master, Handout Master
    #[strum(serialize = "ftr")]
    Footer,
    /// Contains text to be used as a header for the document. Allowed for Notes, Notes Master, Handout Master.
    #[strum(serialize = "hdr")]
    Header,
    /// Contains any content type. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "obj")]
    Object,
    /// Contains a chart or graph. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "chart")]
    Chart,
    /// Contains a table. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "tbl")]
    Table,
    /// Contains a single clip art image. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "clipArt")]
    ClipArt,
    /// Contains a diagram. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "dgm")]
    Diagram,
    /// Contains multimedia content such as audio or a movie clip. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "media")]
    Media,
    /// Contains an image of the slide. Allowed for Notes and Notes Master.
    #[strum(serialize = "sldImg")]
    SlideImage,
    /// Contains a picture. Special type. Allowed for Slide and Slide Layout.
    #[strum(serialize = "pic")]
    Picture,
}

/// This simple type defines a direction of either horizontal or vertical.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum Direction {
    /// Defines a horizontal direction.
    #[strum(serialize = "horz")]
    Horizontal,
    /// Defines a vertical direction.
    #[strum(serialize = "vert")]
    Vertical,
}

/// This simple type facilitates the storing of the size of the placeholder. This size is described relative to the body
/// placeholder on the master.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PlaceholderSize {
    /// Specifies that the placeholder should take the full size of the body placeholder on the master.
    #[strum(serialize = "full")]
    Full,
    /// Specifies that the placeholder should take the half size of the body placeholder on the master. Half size
    /// vertically or horizontally? Needs a picture.
    #[strum(serialize = "half")]
    Half,
    /// Specifies that the placeholder should take a quarter of the size of the body placeholder on the master. Picture
    /// would be helpful
    #[strum(serialize = "quarter")]
    Quarter,
}

/// This simple type defines a set of slide transition directions.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TransitionSideDirectionType {
    /// Specifies that the transition direction is left
    #[strum(serialize = "l")]
    Left,
    /// Specifies that the transition direction is up
    #[strum(serialize = "u")]
    Up,
    /// Specifies that the transition direction is right
    #[strum(serialize = "r")]
    Right,
    /// Specifies that the transition direction is down
    #[strum(serialize = "d")]
    Down,
}

/// This simple type specifies diagonal directions for slide transitions.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TransitionCornerDirectionType {
    /// Specifies the slide transition direction of left-up
    #[strum(serialize = "lu")]
    LeftUp,
    /// Specifies the slide transition direction of right-up
    #[strum(serialize = "ru")]
    RightUp,
    /// Specifies the slide transition direction of left-down
    #[strum(serialize = "ld")]
    LeftDown,
    /// Specifies the slide transition direction of right-down
    #[strum(serialize = "rd")]
    RightDown,
}

/// This simple type specifies the direction of an animation.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TransitionEightDirectionType {
    /// Specifies that the transition direction is left
    #[strum(serialize = "l")]
    Left,
    /// Specifies that the transition direction is up
    #[strum(serialize = "u")]
    Up,
    /// Specifies that the transition direction is right
    #[strum(serialize = "r")]
    Right,
    /// Specifies that the transition direction is down
    #[strum(serialize = "d")]
    Down,
    /// Specifies the slide transition direction of left-up
    #[strum(serialize = "lu")]
    LeftUp,
    /// Specifies the slide transition direction of right-up
    #[strum(serialize = "ru")]
    RightUp,
    /// Specifies the slide transition direction of left-down
    #[strum(serialize = "ld")]
    LeftDown,
    /// Specifies the slide transition direction of right-down
    #[strum(serialize = "rd")]
    RightDown,
}

/// This simple type specifies if a slide transition should go in or out.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TransitionInOutDirectionType {
    /// Specifies the slide transition should go in
    #[strum(serialize = "in")]
    In,
    /// Specifies the slide transition should go out
    #[strum(serialize = "out")]
    Out,
}

/// This simple type defines the allowed transition speeds for transitioning from the current slide to the next.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TransitionSpeed {
    /// Slow slide transition.
    #[strum(serialize = "slow")]
    Slow,
    /// Medium slide transition.
    #[strum(serialize = "med")]
    Medium,
    /// Fast slide transition.
    #[strum(serialize = "fast")]
    Fast,
}

/// This simple type defines an arrangement of content on a slide. Each layout type is not tied to an exact
/// positioning of placeholders, but rather provides a higher-level description of the content type and positioning of
/// placeholders. This information can be used by the application to aid in mapping between different layouts. The
/// application can choose which, if any, of these layouts to make available through its user interface.
///
/// Each layout contains zero or more placeholders, each with a specific content type. An "object" placeholder can
/// contain any kind of data. Media placeholders are intended to hold video or audio clips. The enumeration value
/// descriptions include illustrations of sample layouts for each value of the simple type.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum SlideLayoutType {
    /// Blank
    #[strum(serialize = "blank")]
    Blank,
    /// Title and chart
    #[strum(serialize = "chart")]
    Chart,
    /// Title, chart on left and text on right
    #[strum(serialize = "chartAndTx")]
    ChartAndText,
    /// Title, clipart on left, text on right
    #[strum(serialize = "clipArtAndTx")]
    ClipArtAndText,
    /// Title, clip art on left, vertical text on right
    #[strum(serialize = "clipArtAndVertTx")]
    ClipArtAndVerticalText,
    /// Custom layout defined by user
    #[strum(serialize = "cust")]
    Custom,
    /// Title and diagram
    #[strum(serialize = "dgm")]
    Diagram,
    /// Title and four objects
    #[strum(serialize = "fourObj")]
    FourObjects,
    /// Title, media on left, text on right
    #[strum(serialize = "mediaAndTx")]
    MediaAndText,
    /// Title and object
    #[strum(serialize = "obj")]
    Object,
    /// Title, one object on left, two objects on right
    #[strum(serialize = "objAndTwoObj")]
    ObjectAndTwoObject,
    /// Title, object on left, text on right
    #[strum(serialize = "objAndTx")]
    ObjectAndText,
    /// Object only
    #[strum(serialize = "objOnly")]
    ObjectOnly,
    /// Title, object on top, text on bottom
    #[strum(serialize = "objOverTx")]
    ObjectOverText,
    /// Title, object and caption text
    #[strum(serialize = "objTx")]
    ObjectText,
    /// Title, picture, and caption text
    #[strum(serialize = "picTx")]
    PictureText,
    /// Section header title and subtitle text
    #[strum(serialize = "secHead")]
    SectionHeader,
    /// Title and table
    #[strum(serialize = "tbl")]
    Table,
    /// Title layout with centered title and subtitle placeholders
    #[strum(serialize = "title")]
    Title,
    /// Title only
    #[strum(serialize = "titleOnly")]
    TitleOnly,
    /// Title, text on left, text on right
    #[strum(serialize = "twoColTx")]
    TwoColumnText,
    /// Title, object on left, object on right
    #[strum(serialize = "twoObj")]
    TwoObject,
    /// Title, two objects on left, one object on right
    #[strum(serialize = "twoObjAndObj")]
    TwoObjectsAndObject,
    /// Title, two objects on left, text on right
    #[strum(serialize = "twoObjAndTx")]
    TwoObjectsAndText,
    /// Title, two objects on top, text on bottom
    #[strum(serialize = "twoObjOverTx")]
    TwoObjectsOverText,
    /// Title, two objects each with text
    #[strum(serialize = "twoTxTwoObj")]
    TwoTextTwoObjects,
    /// Title and text
    #[strum(serialize = "tx")]
    Text,
    /// Title, text on left and chart on right
    #[strum(serialize = "txAndChart")]
    TextAndChart,
    /// Title, text on left, clip art on right
    #[strum(serialize = "txAndClipArt")]
    TextAndClipArt,
    /// Title, text on left, media on right
    #[strum(serialize = "txAndMedia")]
    TextAndMedia,
    /// Title, text on left, object on right
    #[strum(serialize = "txAndObj")]
    TextAndObject,
    /// Title, text on left, two objects on right
    #[strum(serialize = "txAndTwoObj")]
    TextAndTwoObjects,
    /// Title, text on top, object on bottom
    #[strum(serialize = "txOverObj")]
    TextOverObject,
    /// Vertical title on right, vertical text on left
    #[strum(serialize = "vertTitleAndTx")]
    VerticalTitleAndText,
    /// Vertical title on right, vertical text on top, chart on bottom
    #[strum(serialize = "vertTitleAndTxOverChart")]
    VerticalTitleAndTextOverChart,
    /// Title and vertical text body
    #[strum(serialize = "vertTx")]
    VerticalText,
}

/// This element specifies an instance of a slide master slide. Within a slide master slide are contained all elements
/// that describe the objects and their corresponding formatting for within a presentation slide. Within a slide
/// master slide are two main elements. The common_slide_data element specifies the common slide elements such as shapes and
/// their attached text bodies. Then the text_styles element specifies the formatting for the text within each of these
/// shapes. The other properties within a slide master slide specify other properties for within a presentation slide
/// such as color information, headers and footers, as well as timing and transition information for all corresponding
/// presentation slides.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideMaster {
    /// Specifies whether the corresponding slide layout is deleted when all the slides that follow
    /// that layout are deleted. If this attribute is not specified then a value of false should be
    /// assumed by the generating application. This would mean that the slide would in fact be
    /// deleted if no slides within the presentation were related to it.
    ///
    /// Defaults to false
    pub preserve: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    /// This element specifies the mapping layer that transforms one color scheme definition to another. Each attribute
    /// represents a color name that can be referenced in this master, and the value is the corresponding color in the
    /// theme.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1"
    /// accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
    /// accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    /// ```
    pub color_mapping: Box<ColorMapping>,
    /// This element specifies the existence of the slide layout identification list. This list is contained within the slide
    /// master and is used to determine which layouts are being used within the slide master file. Each layout within the
    /// list of slide layouts has its own identification number and relationship identifier that uniquely identifies it within
    /// both the presentation document and the particular master slide within which it is used.
    ///
    /// The SlideLayoutIdListEntry specifies the relationship information for each slide layout that is used within the slide master.
    /// The slide master has relationship identifiers that it uses internally for determining the slide layouts that should be
    /// used. Then, to resolve what these slide layouts should be the sldLayoutId elements in the sldLayoutIdLst are
    /// utilized.
    pub slide_layout_id_list: Option<SlideLayoutIdList>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
    pub timing: Option<SlideTiming>,
    /// This element specifies the header and footer information for a slide. Headers and footers consist of
    /// placeholders for text that should be consistent across all slides and slide types, such as a date and time, slide
    /// numbering, and custom header and footer text.
    pub header_footer: Option<HeaderFooter>,
    /// This element specifies the text styles within a slide master. Within this element is the styling information for title
    /// text, the body text and other slide text as well. This element is only for use within the Slide Master and thus sets
    /// the text styles for the corresponding presentation slides.
    pub text_styles: Option<SlideMasterTextStyles>,
}

impl SlideMaster {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;

        Self::from_xml_element(&XmlNode::from_str(xml_string.as_str())?)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preserve = xml_node.attributes.get("preserve").map(parse_xml_bool).transpose()?;

        let mut common_slide_data = None;
        let mut color_mapping = None;
        let mut slide_layout_id_list = None;
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;
        let mut text_styles = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMap" => color_mapping = Some(Box::new(ColorMapping::from_xml_element(child_node)?)),
                "sldLayoutIdLst" => slide_layout_id_list = Some(SlideLayoutIdList::from_xml_element(child_node)?),
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                "hf" => header_footer = Some(HeaderFooter::from_xml_element(child_node)?),
                "txStyles" => text_styles = Some(SlideMasterTextStyles::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data =
            common_slide_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cSld"))?;
        let color_mapping = color_mapping.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrMap"))?;

        Ok(Self {
            common_slide_data,
            color_mapping,
            slide_layout_id_list,
            transition,
            timing,
            header_footer,
            text_styles,
            preserve,
        })
    }
}

/// This element specifies an instance of a slide layout. The slide layout contains in essence a template slide design
/// that can be applied to any existing slide. When applied to an existing slide all corresponding content should be
/// mapped to the new slide layout.
#[derive(Debug, Clone, PartialEq)]
pub struct SlideLayout {
    /// Specifies a name to be used in place of the name attribute within the cSld element. This
    /// is used for layout matching in response to layout changes and template applications.
    ///
    /// Defaults to ""
    pub matching_name: Option<String>,
    /// Specifies the slide layout type that is used by this slide.
    ///
    /// Defaults to SlideLayoutType::Custom
    pub slide_layout_type: Option<SlideLayoutType>,
    /// Specifies whether the corresponding slide layout is deleted when all the slides that follow
    /// that layout are deleted. If this attribute is not specified then a value of false should be
    /// assumed by the generating application. This would mean that the slide would in fact be
    /// deleted if no slides within the presentation were related to it.
    ///
    /// Defaults to false
    pub preserve: Option<bool>,
    /// Specifies if the corresponding object has been drawn by the user and should thus not be
    /// deleted. This allows for the flagging of slides that contain user drawn data.
    pub is_user_drawn: Option<bool>,
    /// Specifies if shapes on the master slide should be shown on slides or not.
    ///
    /// Defaults to true
    pub show_master_shapes: Option<bool>,
    /// Specifies whether or not to display animations on placeholders from the master slide.
    ///
    /// Defaults to true
    pub show_master_placeholder_animations: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    pub color_mapping_override: Option<ColorMappingOverride>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
    pub timing: Option<SlideTiming>,
    pub header_footer: Option<HeaderFooter>,
}

impl SlideLayout {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;

        Self::from_xml_element(&XmlNode::from_str(xml_string.as_str())?)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut matching_name = None;
        let mut slide_layout_type = None;
        let mut preserve = None;
        let mut is_user_drawn = None;
        let mut show_master_shapes = None;
        let mut show_master_placeholder_animations = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "matchingName" => matching_name = Some(value.clone()),
                "type" => slide_layout_type = Some(value.parse()?),
                "preserve" => preserve = Some(parse_xml_bool(value)?),
                "userDrawn" => is_user_drawn = Some(parse_xml_bool(value)?),
                "showMasterSp" => show_master_shapes = Some(parse_xml_bool(value)?),
                "showMasterPhAnim" => show_master_placeholder_animations = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_slide_data = None;
        let mut color_mapping_override = None;
        let mut transition = None;
        let mut timing = None;
        let mut header_footer = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMapOvr" => {
                    color_mapping_override = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(ColorMappingOverride::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| {
                                MissingChildNodeError::new(
                                    child_node.name.clone(),
                                    "masterClrMapping|overrideClrMapping",
                                )
                            })?,
                    );
                }
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                "hf" => header_footer = Some(HeaderFooter::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data =
            common_slide_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cSld"))?;

        Ok(Self {
            matching_name,
            slide_layout_type,
            preserve,
            is_user_drawn,
            show_master_shapes,
            show_master_placeholder_animations,
            common_slide_data,
            color_mapping_override,
            transition,
            timing,
            header_footer,
        })
    }
}

/// This element specifies a slide within a slide list. The slide list is used to specify an ordering of slides.
///
/// # Xml example
///
/// ```xml
/// <p:custShowLst>
///   <p:custShow name="Custom Show 1" id="0">
///     <p:sldLst>
///       <p:sld r:id="rId4"/>
///       <p:sld r:id="rId3"/>
///       <p:sld r:id="rId2"/>
///       <p:sld r:id="rId5"/>
///     </p:sldLst>
///   </p:custShow>
/// </p:custShowLst>
/// ```
/// In the above example the order specified to present the slides is slide 4, then 3, 2 and finally 5.
#[derive(Debug, Clone, PartialEq)]
pub struct Slide {
    /// Specifies that the current slide should be shown in slide show. If this attribute is omitted
    /// then a value of true is assumed.
    ///
    /// Defaults to true
    pub show: Option<bool>,
    /// Specifies if shapes on the master slide should be shown on slides or not.
    ///
    /// Defaults to true
    pub show_master_shapes: Option<bool>,
    /// Specifies whether or not to display animations on placeholders from the master slide.
    ///
    /// Defaults to true
    pub show_master_placeholder_animations: Option<bool>,
    pub common_slide_data: Box<CommonSlideData>,
    /// This element provides a mechanism with which to override the color schemes listed within the
    /// SlideMaster::color_mapping element.
    /// If the ColorMappingOverride::UseMaster element is present, the color scheme defined by the master is used.
    /// If the ColorMappingOverride::Override element is present, it defines a new color scheme specific to the
    /// parent notes slide, presentation slide, or slide layout.
    pub color_mapping_override: Option<ColorMappingOverride>,
    /// This element specifies the kind of slide transition that should be used to transition to the current slide from the
    /// previous slide. That is, the transition information is stored on the slide that appears after the transition is
    /// complete.
    pub transition: Option<Box<SlideTransition>>,
    /// This element specifies the timing information for handling all animations and timed events within the
    /// corresponding slide. This information is tracked via time nodes within the timing element. More information on
    /// the specifics of these time nodes and how they are to be defined can be found within the Animation section of
    /// the PresentationML framework.
    pub timing: Option<SlideTiming>,
}

impl Slide {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;

        Self::from_xml_element(&XmlNode::from_str(xml_string.as_str())?)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut show = None;
        let mut show_master_shapes = None;
        let mut show_master_placeholder_animations = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "show" => show = Some(parse_xml_bool(value)?),
                "showMasterSp" => show_master_shapes = Some(parse_xml_bool(value)?),
                "showMasterPhAnim" => show_master_placeholder_animations = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_slide_data = None;
        let mut color_mapping_override = None;
        let mut transition = None;
        let mut timing = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cSld" => common_slide_data = Some(Box::new(CommonSlideData::from_xml_element(child_node)?)),
                "clrMapOvr" => {
                    color_mapping_override = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(ColorMappingOverride::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| {
                                MissingChildNodeError::new(
                                    child_node.name.clone(),
                                    "masterClrMapping|overrideClrMapping",
                                )
                            })?,
                    );
                }
                "transition" => transition = Some(Box::new(SlideTransition::from_xml_element(child_node)?)),
                "timing" => timing = Some(SlideTiming::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_slide_data =
            common_slide_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cSld"))?;

        Ok(Self {
            show,
            show_master_shapes,
            show_master_placeholder_animations,
            common_slide_data,
            color_mapping_override,
            transition,
            timing,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct BackgroundProperties {
    /// Specifies whether the background of the slide is of a shade to title background type. This
    /// kind of gradient fill is on the slide background and changes based on the placement of
    /// the slide title placeholder.
    ///
    /// Defaults to false
    pub shade_to_title: Option<bool>,
    pub fill: FillProperties,
    pub effect: Option<EffectProperties>,
}

impl BackgroundProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let shade_to_title = xml_node
            .attributes
            .get("shadeToTitle")
            .map(parse_xml_bool)
            .transpose()?;

        let mut fill = None;
        let mut effect = None;

        for child_node in &xml_node.child_nodes {
            if FillProperties::is_choice_member(child_node.local_name()) {
                fill = Some(FillProperties::from_xml_element(child_node)?);
            } else if EffectProperties::is_choice_member(child_node.local_name()) {
                effect = Some(EffectProperties::from_xml_element(child_node)?);
            }
        }

        let fill = fill.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;

        Ok(Self {
            shade_to_title,
            fill,
            effect,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum BackgroundGroup {
    /// This element specifies visual effects used to render the slide background. This includes any fill, image, or effects
    /// that are to make up the background of the slide.
    Properties(BackgroundProperties),
    /// This element specifies the slide background is to use a fill style defined in the style matrix.
    /// The idx attribute refers to the index of a background fill style or fill style within the presentation's
    /// style matrix, defined by the fmtScheme element.
    /// A value of 0 or 1000 indicates no background, values 1-999 refer to the index of a fill style
    /// within the fillStyleLst element, and values 1001 and above refer to the index of a background fill style within
    /// the bgFillStyleLst element. The value 1001 corresponds to the first background fill style, 1002 to the second
    /// background fill style, and so on.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:bgRef idx="2">
    ///   <a:schemeClr val="bg2"/>
    /// </p:bgRef>
    /// ```
    ///
    /// The above code indicates a slide background with the style's second fill style using the second background color
    /// of the color scheme.
    ///
    /// ```xml
    /// <p:bgRef idx="1001">
    ///   <a:schemeClr val="bg2"/>
    /// </p:bgRef>
    /// ```
    ///
    /// The above code indicates a slide background with the style's first background fill style using the second
    /// background color of the color scheme.
    Reference(StyleMatrixReference),
}

impl XsdType for BackgroundGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bgPr" => Ok(BackgroundGroup::Properties(BackgroundProperties::from_xml_element(
                xml_node,
            )?)),
            "bgRef" => Ok(BackgroundGroup::Reference(StyleMatrixReference::from_xml_element(
                xml_node,
            )?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Background").into()),
        }
    }
}

impl XsdChoice for BackgroundGroup {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "bgPr" | "bgRef" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Background {
    /// Specifies that the background should be rendered using only black and white coloring.
    /// That is, the coloring information for the background should be converted to either black
    /// or white when rendering the picture.
    ///
    /// # Note
    ///
    /// No gray is to be used in rendering this background, only stark black and stark
    /// white.
    pub black_and_white_mode: Option<BlackWhiteMode>, // white
    pub background: BackgroundGroup,
}

impl Background {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_and_white_mode = xml_node.attributes.get("bwMode").map(|val| val.parse()).transpose()?;

        let background = xml_node
            .child_nodes
            .iter()
            .find_map(BackgroundGroup::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Background"))?;

        Ok(Self {
            background,
            black_and_white_mode,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Placeholder {
    /// Specifies what content type a placeholder is intended to contain.
    pub placeholder_type: Option<PlaceholderType>,
    /// Specifies the orientation of a placeholder.
    pub orientation: Option<Direction>,
    /// Specifies the size of a placeholder.
    pub size: Option<PlaceholderSize>,
    /// Specifies the placeholder index. This is used when applying templates or changing
    /// layouts to match a placeholder on one template/master to another.
    pub index: Option<u32>,
    /// Specifies whether the corresponding placeholder should have a custom prompt or not.
    pub has_custom_prompt: Option<bool>,
}

impl Placeholder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "type" => instance.placeholder_type = Some(value.parse()?),
                    "orient" => instance.orientation = Some(value.parse()?),
                    "sz" => instance.size = Some(value.parse()?),
                    "idx" => instance.index = Some(value.parse()?),
                    "hasCustomPrompt" => instance.has_custom_prompt = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies non-visual properties for objects. These properties include multimedia content associated
/// with an object and properties indicating how the object is to be used or displayed in different contexts.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct ApplicationNonVisualDrawingProps {
    /// Specifies whether the picture belongs to a photo album and should thus be included
    /// when editing a photo album within the generating application.
    pub is_photo: Option<bool>,
    /// Specifies if the corresponding object has been drawn by the user and should thus not be
    /// deleted. This allows for the flagging of slides that contain user drawn data.
    pub is_user_drawn: Option<bool>, // false
    /// This element specifies that the corresponding shape should be represented by the generating application as a
    /// placeholder. When a shape is considered a placeholder by the generating application it can have special
    /// properties to alert the user that they can enter content into the shape. Different placeholder types are allowed
    /// and can be specified by using the placeholder type attribute for this element.
    pub placeholder: Option<Placeholder>,
    pub media: Option<Media>,
    pub customer_data_list: Option<CustomerDataList>,
}

impl ApplicationNonVisualDrawingProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "isPhoto" => instance.is_photo = Some(parse_xml_bool(value)?),
                    "userDrawn" => instance.is_user_drawn = Some(parse_xml_bool(value)?),
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
                            "ph" => instance.placeholder = Some(Placeholder::from_xml_element(child_node)?),
                            "custDataLst" => {
                                instance.customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?)
                            }
                            local_name if Media::is_choice_member(local_name) => {
                                instance.media = Some(Media::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ShapeGroup {
    /// This element specifies the existence of a single shape. A shape can either be a preset or a custom geometry,
    /// defined using the DrawingML framework. In addition to a geometry each shape can have both visual and non-
    /// visual properties attached. Text and corresponding styling information can also be attached to a shape. This
    /// shape is specified along with all other shapes within either the shape tree or group shape elements.
    Shape(Box<Shape>),
    /// This element specifies a group shape that represents many shapes grouped together. This shape is to be treated
    /// just as if it were a regular shape but instead of being described by a single geometry it is made up of all the
    /// shape geometries encompassed within it. Within a group shape each of the shapes that make up the group are
    /// specified just as they normally would. The idea behind grouping elements however is that a single transform can
    /// apply to many shapes at the same time.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:grpSp>
    ///   <p:nvGrpSpPr>
    ///     <p:cNvPr id="10" name="Group 9"/>
    ///     <p:cNvGrpSpPr/>
    ///     <p:nvPr/>
    ///   </p:nvGrpSpPr>
    ///   <p:grpSpPr>
    ///     <a:xfrm>
    ///       <a:off x="838200" y="990600"/>
    ///       <a:ext cx="2426208" cy="978408"/>
    ///       <a:chOff x="838200" y="990600"/>
    ///       <a:chExt cx="2426208" cy="978408"/>
    ///     </a:xfrm>
    ///   </p:grpSpPr>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///   ...
    ///   </p:sp>
    /// </p:grpSp>
    /// ```
    ///
    /// In the above example we see three shapes specified within a single group. These three shapes have their
    /// position and sizes specified just as they normally would within the shape tree.
    /// The generating application should apply the transformation after the bounding box for the group shape has been
    /// calculated.
    GroupShape(Box<GroupShape>),
    /// This element specifies the existence of a graphics frame. This frame contains a graphic that was generated by an
    /// external source and needs a container in which to be displayed on the slide surface.
    GraphicFrame(Box<GraphicalObjectFrame>),
    /// This element specifies a connection shape that is used to connect two sp elements.
    /// Once a connection is specified using a Connector, it is left to the generating application to determine the
    /// exact path the connector takes.
    /// That is the connector routing algorithm is left up to the generating application as the desired path might be
    /// different depending on the specific needs of the application.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:spTree>
    ///   ...
    ///   <p:sp>
    ///     <p:nvSpPr>
    ///       <p:cNvPr id="1" name="Rectangle 1"/>
    ///       <p:cNvSpPr/>
    ///       <p:nvPr/>
    ///     </p:nvSpPr>
    ///   ...
    ///   </p:sp>
    ///   <p:sp>
    ///     <p:nvSpPr>
    ///       <p:cNvPr id="2" name="Rectangle 2"/>
    ///       <p:cNvSpPr/>
    ///       <p:nvPr/>
    ///     </p:nvSpPr>
    ///     ...
    ///   </p:sp>
    ///   <p:cxnSp>
    ///     <p:nvCxnSpPr>
    ///       <p:cNvPr id="3" name="Elbow Connector 3"/>
    ///       <p:cNvCxnSpPr>
    ///         <a:stCxn id="1" idx="3"/>
    ///         <a:endCxn id="2" idx="1"/>
    ///       </p:cNvCxnSpPr>
    ///       <p:nvPr/>
    ///     </p:nvCxnSpPr>
    ///     ...
    ///   </p:cxnSp>
    /// </p:spTree>
    /// ```
    Connector(Box<Connector>),
    /// This element specifies the existence of a picture object within the document.
    ///
    /// # Xml example
    ///
    /// Consider the following PresentationML that specifies the existence of a picture within a document.
    /// This picture can have non-visual properties, a picture fill as well as shape properties attached to it.
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="4" name="lake.JPG" descr="Picture of a Lake" />
    ///     <p:cNvPicPr>
    ///       <a:picLocks noChangeAspect="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr/>
    ///   </p:nvPicPr>
    ///   <p:blipFill>
    ///   ...
    ///   </p:blipFill>
    ///   <p:spPr>
    ///   ...
    ///   </p:spPr>
    /// </p:pic>
    /// ```
    Picture(Box<Picture>),
    /// This element specifies a reference to XML content in a format not defined by ECMA-376.
    ///
    /// The relationship type of the explicit relationship specified by this element shall be
    /// http://purl.oclc.org/ooxml/officeDocument/relationships/customXml and have a TargetMode attribute value of
    /// Internal. If an application cannot process content of the content type specified by the targeted part, then it
    /// should continue to process the file.
    /// If possible, it should also provide some indication that unknown content was not imported.
    ///
    /// # Note
    ///
    /// This part allows the native use of other commonly used interchange formats, such as:
    /// * [MathML](http://www.w3.org/TR/MathML2/)
    /// * [SMIL](http://www.w3.org/TR/REC-smil/)
    /// * [SVG](http://www.w3.org/TR/SVG11/)
    ///
    /// For better interoperability, only standard XML formats should be used.
    ContentPart(RelationshipId),
}

impl XsdType for ShapeGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "sp" => Ok(ShapeGroup::Shape(Box::new(Shape::from_xml_element(xml_node)?))),
            "grpSp" => Ok(ShapeGroup::GroupShape(Box::new(GroupShape::from_xml_element(
                xml_node,
            )?))),
            "graphicFrame" => Ok(ShapeGroup::GraphicFrame(Box::new(
                GraphicalObjectFrame::from_xml_element(xml_node)?,
            ))),
            "cxnSp" => Ok(ShapeGroup::Connector(Box::new(Connector::from_xml_element(xml_node)?))),
            "pic" => Ok(ShapeGroup::Picture(Box::new(Picture::from_xml_element(xml_node)?))),
            "contentPart" => {
                let rel_id = xml_node
                    .attributes
                    .get("r:id")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?
                    .clone();

                Ok(ShapeGroup::ContentPart(rel_id))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_ShapeGroup",
            ))),
        }
    }
}

impl XsdChoice for ShapeGroup {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "sp" | "grpSp" | "graphicFrame" | "cxnSp" | "pic" | "contentPart" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    /// Specifies that the shape fill should be set to that of the slide background surface.
    ///
    /// # Note
    ///
    /// This attribute does not set the fill of the shape to be transparent but instead sets it
    /// to be filled with the portion of the slide background that is directly behind it.
    pub use_bg_fill: Option<bool>,
    /// This element specifies all non-visual properties for a shape. This element is a container for the non-visual
    /// identification properties, shape properties and application properties that are to be associated with a shape.
    /// This allows for additional information that does not affect the appearance of the shape to be stored.
    pub non_visual_props: Box<ShapeNonVisual>,
    /// This element specifies the visual shape properties that can be applied to a shape. These properties include the
    /// shape fill, outline, geometry, effects, and 3D orientation.
    pub shape_props: Box<ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    ///
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<ShapeStyle>>,
    /// This element specifies the existence of text to be contained within the corresponding shape. All visible text and
    /// visible text related properties are contained within this element. There can be multiple paragraphs and within
    /// paragraphs multiple runs of text.
    pub text_body: Option<TextBody>,
}

impl Shape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let use_bg_fill = xml_node.attributes.get("useBgFill").map(parse_xml_bool).transpose()?;

        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut shape_style = None;
        let mut text_body = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvSpPr" => non_visual_props = Some(Box::new(ShapeNonVisual::from_xml_element(child_node)?)),
                "spPr" => shape_props = Some(Box::new(ShapeProperties::from_xml_element(child_node)?)),
                "style" => shape_style = Some(Box::new(ShapeStyle::from_xml_element(child_node)?)),
                "txBody" => text_body = Some(TextBody::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;

        Ok(Self {
            use_bg_fill,
            non_visual_props,
            shape_props,
            shape_style,
            text_body,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ShapeNonVisual {
    pub drawing_props: Box<NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a shape. These properties are to be used by the
    /// generating application to determine how the shape should be dealt with.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:nvSpPr>
    ///     <p:cNvPr id="2" name="Rectangle 1"/>
    ///     <p:cNvSpPr>
    ///       <a:spLocks noGrp="1"/>
    ///     </p:cNvSpPr>
    ///   </p:nvSpPr>
    ///   ...
    /// </p:sp>
    /// ```
    ///
    /// This shape lock is stored within the non-visual drawing properties for this shape.
    pub shape_drawing_props: NonVisualDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut shape_drawing_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(NonVisualDrawingProps::from_xml_element(child_node)?)),
                "cNvSpPr" => shape_drawing_props = Some(NonVisualDrawingShapeProps::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let shape_drawing_props =
            shape_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            shape_drawing_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GroupShape {
    /// This element specifies all non-visual properties for a group shape. This element is a container for the
    /// non-visual identification properties, shape properties and application properties that are to be associated
    /// with a group shape.
    /// This allows for additional information that does not affect the appearance of the group shape to be stored.
    pub non_visual_props: Box<GroupShapeNonVisual>,
    /// This element specifies the properties that are to be common across all of the shapes within the corresponding
    /// group. If there are any conflicting properties within the group shape properties and the individual shape
    /// properties then the individual shape properties should take precedence.
    pub group_shape_props: GroupShapeProperties,
    pub shape_array: Vec<ShapeGroup>,
}

impl GroupShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut group_shape_props = None;
        let mut shape_array = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvGrpSpPr" => non_visual_props = Some(Box::new(GroupShapeNonVisual::from_xml_element(child_node)?)),
                "grpSpPr" => group_shape_props = Some(GroupShapeProperties::from_xml_element(child_node)?),
                local_name if ShapeGroup::is_choice_member(local_name) => {
                    shape_array.push(ShapeGroup::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvGrpSpPr"))?;
        let group_shape_props =
            group_shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "grpSpPr"))?;

        Ok(Self {
            non_visual_props,
            group_shape_props,
            shape_array,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GroupShapeNonVisual {
    pub drawing_props: Box<NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a group shape. These non-visual properties are
    /// properties that the generating application would utilize when rendering the slide surface.
    pub group_drawing_props: NonVisualGroupDrawingShapeProps,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GroupShapeNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut group_drawing_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(NonVisualDrawingProps::from_xml_element(child_node)?)),
                "cNvGrpSpPr" => {
                    group_drawing_props = Some(NonVisualGroupDrawingShapeProps::from_xml_element(child_node)?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let group_drawing_props =
            group_drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGrpSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            group_drawing_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GraphicalObjectFrame {
    /// Specifies how the graphical object should be rendered, using color, black or white,
    /// or grayscale.
    ///
    /// # Note
    ///
    /// This does not mean that the graphical object itself is stored with only black
    /// and white or grayscale information. This attribute instead sets the rendering mode
    /// that the graphical object uses.
    pub black_white_mode: Option<BlackWhiteMode>,
    /// This element specifies all non-visual properties for a graphic frame. This element is a container for the
    /// non-visual identification properties, shape properties and application properties that are to be associated
    /// with a graphic frame.
    /// This allows for additional information that does not affect the appearance of the graphic frame to be stored.
    pub non_visual_props: Box<GraphicalObjectFrameNonVisual>,
    /// This element specifies the transform to be applied to the corresponding graphic frame. This transformation is
    /// applied to the graphic frame just as it would be for a shape or group shape.
    pub transform: Box<Transform2D>,
    pub graphic: GraphicalObject,
}

impl GraphicalObjectFrame {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let black_white_mode = xml_node
            .attributes
            .get("bwMode")
            .map(|value| value.parse())
            .transpose()?;

        let mut non_visual_props = None;
        let mut transform = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvGraphicFramePr" => {
                    non_visual_props = Some(Box::new(GraphicalObjectFrameNonVisual::from_xml_element(child_node)?))
                }
                "xfrm" => transform = Some(Box::new(Transform2D::from_xml_element(child_node)?)),
                "graphic" => graphic = Some(GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvGraphicFramePr"))?;
        let transform = transform.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "xfrm"))?;
        let graphic = graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?;

        Ok(Self {
            black_white_mode,
            non_visual_props,
            transform,
            graphic,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GraphicalObjectFrameNonVisual {
    pub drawing_props: Box<NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties for a graphic frame. These non-visual properties are
    /// properties that the generating application would utilize when rendering the slide surface.
    pub graphic_frame_props: NonVisualGraphicFrameProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl GraphicalObjectFrameNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut graphic_frame_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(NonVisualDrawingProps::from_xml_element(child_node)?)),
                "cNvGraphicFramePr" => {
                    graphic_frame_props = Some(NonVisualGraphicFrameProperties::from_xml_element(child_node)?)
                }
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let graphic_frame_props = graphic_frame_props
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvGraphicFramePr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            graphic_frame_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Connector {
    /// This element specifies all non-visual properties for a connection shape. This element is a container for the non-
    /// visual identification properties, shape properties and application properties that are to be associated with a
    /// connection shape. This allows for additional information that does not affect the appearance of the connection
    /// shape to be stored.
    pub non_visual_props: Box<ConnectorNonVisual>,
    /// This element specifies the visual shape properties that can be applied to a shape. These properties include the
    /// shape fill, outline, geometry, effects, and 3D orientation.
    pub shape_props: Box<ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    ///
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<ShapeStyle>>,
}

impl Connector {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvCxnSpPr" => non_visual_props = Some(Box::new(ConnectorNonVisual::from_xml_element(child_node)?)),
                "spPr" => shape_props = Some(Box::new(ShapeProperties::from_xml_element(child_node)?)),
                "style" => shape_style = Some(Box::new(ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvCxnSpPr"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvCxnSpPr"))?;

        Ok(Self {
            non_visual_props,
            shape_props,
            shape_style,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ConnectorNonVisual {
    pub drawing_props: Box<NonVisualDrawingProps>,
    /// This element specifies the non-visual drawing properties specific to a connector shape. This includes
    /// information specifying the shapes to which the connector shape is connected.
    pub connector_props: NonVisualConnectorProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl ConnectorNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut connector_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(NonVisualDrawingProps::from_xml_element(child_node)?)),
                "cNvCxnSpPr" => connector_props = Some(NonVisualConnectorProperties::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let connector_props =
            connector_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvCxnSpPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            connector_props,
            app_props,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Picture {
    /// This element specifies all non-visual properties for a picture. This element is a container for the non-visual
    /// identification properties, shape properties and application properties that are to be associated with a picture.
    /// This allows for additional information that does not affect the appearance of the picture to be stored.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   ...
    ///   <p:nvPicPr>
    ///   ...
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    pub non_visual_props: Box<PictureNonVisual>,
    pub blip_fill: Box<BlipFillProperties>,
    pub shape_props: Box<ShapeProperties>,
    /// This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
    /// the preset styles defined by the style matrix for the theme.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:style>
    ///   <a:lnRef idx="3">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:lnRef>
    ///   <a:fillRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:fillRef>
    ///   <a:effectRef idx="1">
    ///     <a:schemeClr val="accent3"/>
    ///   </a:effectRef>
    ///   <a:fontRef idx="minor">
    ///     <a:schemeClr val="lt1"/>
    ///   </a:fontRef>
    /// </p:style>
    /// ```
    ///
    /// The parent shape of the above code is to have an outline that uses the third line style defined by the theme, use
    /// the first fill defined by the scheme, and be rendered with the first effect defined by the theme. Text inside the
    /// shape is to use the minor font defined by the theme.
    pub shape_style: Option<Box<ShapeStyle>>,
}

impl Picture {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut non_visual_props = None;
        let mut blip_fill = None;
        let mut shape_props = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "nvPicPr" => non_visual_props = Some(Box::new(PictureNonVisual::from_xml_element(child_node)?)),
                "blipFill" => blip_fill = Some(Box::new(BlipFillProperties::from_xml_element(child_node)?)),
                "spPr" => shape_props = Some(Box::new(ShapeProperties::from_xml_element(child_node)?)),
                "style" => shape_style = Some(Box::new(ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let non_visual_props =
            non_visual_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPicPr"))?;
        let blip_fill = blip_fill.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "blipFill"))?;
        let shape_props = shape_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;

        Ok(Self {
            non_visual_props,
            blip_fill,
            shape_props,
            shape_style,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PictureNonVisual {
    pub drawing_props: Box<NonVisualDrawingProps>,
    /// This element specifies the non-visual properties for the picture canvas. These properties are to be used by the
    /// generating application to determine how certain properties are to be changed for the picture object in question.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   ...
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="4" name="Lilly_by_Lisher.jpg"/>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noChangeAspect="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr/>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    pub picture_props: NonVisualPictureProperties,
    pub app_props: ApplicationNonVisualDrawingProps,
}

impl PictureNonVisual {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut drawing_props = None;
        let mut picture_props = None;
        let mut app_props = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cNvPr" => drawing_props = Some(Box::new(NonVisualDrawingProps::from_xml_element(child_node)?)),
                "cNvPicPr" => picture_props = Some(NonVisualPictureProperties::from_xml_element(child_node)?),
                "nvPr" => app_props = Some(ApplicationNonVisualDrawingProps::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let drawing_props = drawing_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPr"))?;
        let picture_props =
            picture_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cNvPicPr"))?;
        let app_props = app_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "nvPr"))?;

        Ok(Self {
            drawing_props,
            picture_props,
            app_props,
        })
    }
}

/// This element specifies a container for slide information that is relevant to all of the slide types.
/// All slides share a common set of properties that is independent of the slide type; the description of these
/// properties for any particular slide is stored within the slide's common_slide_data container.
/// Slide data specific to the slide type indicated by the parent element is stored elsewhere.
///
/// # Note
///
/// The actual data in CommonSlideData describe only the particular parent slide; it is only the kind of information
/// stored that is common across all slides.
///
/// # Xml example
///
/// ```xml
/// <p:sld>
///   <p:cSld>
///     <p:spTree>
///     ...
///     </p:spTree>
///   </p:cSld>
///   ...
/// </p:sld>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct CommonSlideData {
    /// Specifies the slide name property that is used to further identify this unique configuration
    /// of common slide data. This might be used to aid in distinguishing different slide layouts or
    /// various other slide types.
    pub name: Option<String>,
    /// This element specifies the background appearance information for a slide. The slide background covers the
    /// entire slide and is visible where no objects exist and as the background for transparent objects.
    pub background: Option<Box<Background>>,
    /// This element specifies all shape-based objects, either grouped or not, that can be referenced on a given slide. As
    /// most objects within a slide are shapes, this represents the majority of content within a slide. Text and effects are
    /// attached to shapes that are contained within the shape_tree element.
    ///
    /// Each shape-based object within the shape tree, whether grouped or not, shall represent one unique level of z-
    /// ordering on the slide. The z-order for each shape-based object shall be determined by the lexical ordering of
    /// each shape-based object within the shape tree: the first shape-based object shall have the lowest z-order, while
    /// the last shape-based object shall have the highest z-order.
    ///
    /// The z-ordering of shape-based objects within the shape tree shall also determine the navigation (tab) order of
    /// the shape-based objects: the shape-based object with the lowest z-order (the first shape in lexical order) shall be
    /// first in navigation order, with objects being navigated in ascending z-order.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:sld>
    ///   <p:cSld>
    ///     <p:spTree>
    ///       <p:nvGrpSpPr>
    ///       ...
    ///       </p:nvGrpSpPr>
    ///       <p:grpSpPr>
    ///       ...
    ///       </p:grpSpPr>
    ///       <p:sp>
    ///       ...
    ///       </p:sp>
    ///     </p:spTree>
    ///   </p:cSld>
    ///   ...
    /// </p:sld>
    /// ```
    pub shape_tree: Box<GroupShape>,
    pub customer_data_list: Option<CustomerDataList>,
    /// This element specifies a list of embedded controls for the corresponding slide. Custom embedded controls can
    /// be embedded on slides.
    pub control_list: Option<Vec<Control>>,
}

impl CommonSlideData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attributes.get("name").cloned();
        let mut background = None;
        let mut shape_tree = None;
        let mut customer_data_list = None;
        let mut control_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "bg" => background = Some(Box::new(Background::from_xml_element(child_node)?)),
                "spTree" => shape_tree = Some(Box::new(GroupShape::from_xml_element(child_node)?)),
                "custDataList" => customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?),
                "controls" => {
                    control_list = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .filter(|control_node| control_node.local_name() == "control")
                            .map(Control::from_xml_element)
                            .collect::<Result<Vec<_>>>()?,
                    );
                }
                _ => (),
            }
        }

        let shape_tree = shape_tree.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spTree"))?;

        Ok(Self {
            name,
            background,
            shape_tree,
            customer_data_list,
            control_list,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct SlideMasterTextStyles {
    /// This element specifies the text formatting style for the title text within a master slide. This formatting is used on
    /// all title text within related presentation slides. The text formatting is specified by utilizing the DrawingML
    /// framework just as within a regular presentation slide. Within a title style there can be many different style types
    /// defined as there are different kinds of text stored within a slide title.
    pub title_styles: Option<Box<TextListStyle>>,
    /// This element specifies the text formatting style for all body text within a master slide.
    /// This formatting is used on all body text within presentation slides related to this master.
    /// The text formatting is specified by utilizing the DrawingML framework just as within a regular
    /// presentation slide.
    /// Within the bodyStyle element there can be many different style types defined as there are different kinds of
    /// text stored within the body of a slide.
    pub body_styles: Option<Box<TextListStyle>>,
    /// This element specifies the text formatting style for the all other text within a master slide. This formatting is
    /// used on all text not covered by the title_styles or body_styles elements within related presentation slides. The text
    /// formatting is specified by utilizing the DrawingML framework just as within a regular presentation slide. Within
    /// the otherStyle element there can be many different style types defined as there are different kinds of text
    /// stored within a slide.
    ///
    /// # Note
    ///
    /// The other_styles element is to be used for specifying the text formatting of text within a slide shape but
    /// not within a text box. Text box styling is handled from within the body_styles element.
    pub other_styles: Option<Box<TextListStyle>>,
}

impl SlideMasterTextStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "titleStyle" => instance.title_styles = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "bodyStyle" => instance.body_styles = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "otherStyle" => instance.other_styles = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct OrientationTransition {
    /// This attribute specifies a horizontal or vertical transition.
    ///
    /// Defaults to Direction::Horizontal
    pub direction: Option<Direction>,
}

impl OrientationTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = xml_node.attributes.get("dir").map(|value| value.parse()).transpose()?;

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct EightDirectionTransition {
    /// This attribute specifies if the direction of the transition.
    ///
    /// Defaults to TransitionEightDirectionType::Left
    pub direction: Option<TransitionEightDirectionType>,
}

impl EightDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = xml_node.attributes.get("dir").map(|value| value.parse()).transpose()?;

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct OptionalBlackTransition {
    /// This attribute specifies if the transition starts from a black screen (and then transition the
    /// new slide over black).
    ///
    /// Defaults to false
    pub through_black: Option<bool>,
}

impl OptionalBlackTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let through_black = xml_node.attributes.get("thruBlk").map(parse_xml_bool).transpose()?;

        Ok(Self { through_black })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct SideDirectionTransition {
    /// This attribute specifies the direction of the slide transition.
    ///
    /// Defaults to TransitionSideDirectionType::Left
    pub direction: Option<TransitionSideDirectionType>,
}

impl SideDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = xml_node.attributes.get("dir").map(|value| value.parse()).transpose()?;

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct SplitTransition {
    /// This attribute specifies the orientation of a "split" slide transition.
    ///
    /// Defaults to Direction::Horizontal
    pub orientation: Option<Direction>,
    /// This attribute specifies the direction of a "split" slide transition.
    ///
    /// Defaults to TransitionInOutDirectionType::Out
    pub direction: Option<TransitionInOutDirectionType>,
}

impl SplitTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "orient" => instance.orientation = Some(value.parse()?),
                    "dir" => instance.direction = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct CornerDirectionTransition {
    /// This attribute specifies if the direction of the transition.
    ///
    /// Defaults to TransitionCornerDirectionType::LeftUp
    pub direction: Option<TransitionCornerDirectionType>,
}

impl CornerDirectionTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = xml_node.attributes.get("dir").map(|value| value.parse()).transpose()?;

        Ok(Self { direction })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct WheelTransition {
    /// This attributes specifies the number of spokes ("pie pieces") in the wheel
    ///
    /// Defaults to 4
    pub spokes: Option<u32>,
}

impl WheelTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let spokes = xml_node
            .attributes
            .get("spokes")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self { spokes })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct InOutTransition {
    /// This attribute specifies the direction of an "in/out" slide transition.
    ///
    /// Defaults to TransitionInOutDirectionType::Out
    pub direction: Option<TransitionInOutDirectionType>,
}

impl InOutTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let direction = xml_node.attributes.get("dir").map(|value| value.parse()).transpose()?;

        Ok(Self { direction })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum SlideTransitionGroup {
    /// This element describes the blinds slide transition effect, which uses a set of horizontal or vertical bars and wipes
    /// them either left-to-right or top-to-bottom, respectively, until the new slide is fully shown. The rendering of this
    /// transition depends upon the attributes specified.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:blinds dir="horz"/>
    /// </p:transition>
    /// ```
    Blinds(OrientationTransition),
    /// This element describes the checker slide transition effect, which uses a set of horizontal or vertical
    /// checkerboard squares and wipes them either left-to-right or top-to-bottom, respectively, until the new slide is
    /// fully shown. The rendering of this transition depends upon the attributes specified.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:checker dir="horz"/>
    /// </p:transition>
    /// ```
    Checker(OrientationTransition),
    /// This element describes the circle slide transition effect, which uses a circle pattern centered on the slide that
    /// increases in size until the new slide is fully shown. The rendering of this transition has been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:circle/>
    /// </p:transition>
    /// ```
    Circle,
    /// This element describes the dissolve slide transition effect, which uses a set of randomly placed squares on the
    /// slide that continue to be added to until the new slide is fully shown. The rendering of this transition has been
    /// shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:dissolve/>
    /// </p:transition>
    /// ```
    Dissolve,
    /// This element describes the comb slide transition effect, which uses a set of horizontal or vertical bars and wipes
    /// them from one end of the slide to the other until the new slide is fully shown. The rendering of this transition
    /// depends upon the attributes specified which have been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:comb dir="horz"/>
    /// </p:transition>
    /// ```
    Comb(OrientationTransition),
    /// This element describes the cover slide transition effect, which moves the new slide in from an off-screen
    /// location, continually covering more of the previous slide until the new slide is fully shown. The rendering of this
    /// transition depends upon the attributes specified which have been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:cover dir="d"/>
    /// </p:transition>
    /// ```
    Cover(EightDirectionTransition),
    /// This element describes the cut slide transition effect, which simply replaces the previous slide with the new slide
    /// instantaneously. No animation is used, but an option exists to cut to a black screen before showing the new
    /// slide. The rendering of this transition depends upon the attributes specified which have been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:cut thruBlk="0"/>
    /// </p:transition>
    /// ```
    Cut(OptionalBlackTransition),
    /// This element describes the diamond slide transition effect, which uses a diamond pattern centered on the slide
    /// that increases in size until the new slide is fully shown. The rendering of this transition has been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:diamond/>
    /// </p:transition>
    /// ```
    Diamond,
    /// This element describes the fade slide transition effect, which smoothly fades the previous slide either directly to
    /// the new slide or first to a black screen and then to the new slide. The rendering of this transition depends upon
    /// the attributes specified which have been shown below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:fade thruBlk="0"/>
    /// </p:transition>
    /// ```
    Fade(OptionalBlackTransition),
    /// This element describes the newsflash slide transition effect, which grows and spins the new slide counterclockwise
    /// into place over the previous slide. The rendering of this transition has been shown below.
    ///
    /// # Xml example
    ///
    /// Consider the following case in which the “newsflash” slide transition is applied to a slide, along with a
    /// set of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:newsflash/>
    /// </p:transition>
    /// ```
    Newsflash,
    /// This element describes the plus slide transition effect, which uses a plus pattern centered on the slide that
    /// increases in size until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following case in which the “plus” slide transition is applied to a slide, along with a set of
    /// attributes
    /// ```xml
    /// <p:transition>
    ///   <p:plus/>
    /// </p:transition>
    /// ```
    Plus,
    /// This element describes the pull slide transition effect, which moves the previous slide to an off-screen location,
    /// continually revealing more of the new slide until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “pull” slide transition is applied to a slide, along with a set of
    /// attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:pull dir="d"/>
    /// </p:transition>
    /// ```
    Pull(EightDirectionTransition),
    /// This element describes the push slide transition effect, which moves the new slide in from an off-screen
    /// location, continually pushing the previous slide to an opposite off-screen location until the new slide is fully
    /// shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “push” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:push dir="d"/>
    /// </p:transition>
    /// ```
    Push(SideDirectionTransition),
    /// This element describes the random slide transition effect, which chooses a random transition from the set
    /// available in the rendering application. This transition thus can be different each time it is used.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:random/>
    /// </p:transition>
    /// ```
    Random,
    /// This element describes the randomBar slide transition effect, which uses a set of randomly placed horizontal or
    /// vertical bars on the slide that continue to be added to until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “randomBar” slide transition is applied to a slide, along with
    /// a set of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:randomBar dir="horz"/>
    /// </p:transition>
    /// ```
    RandomBar(OrientationTransition),
    /// This element describes the split slide transition effect, which reveals the new slide directly on top of the
    /// previous one by wiping either horizontal or vertical from the outside in, or from the inside out, until the new
    /// slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “split” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:split orient="horz" dir="in"/>
    /// </p:transition>
    /// ```
    Split(SplitTransition),
    /// This element describes the strips slide transition effect, which uses a set of bars that are arranged in a staggered
    /// fashion and wipes them across the screen until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “strips” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:strips dir="ld"/>
    /// </p:transition>
    /// ```
    Strips(CornerDirectionTransition),
    /// This element describes the wedge slide transition effect, which uses two radial edges that wipe from top to
    /// bottom in opposite directions until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following case in which the “wedge” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wedge/>
    /// </p:transition>
    /// ```
    Wedge,
    /// This element describes the wheel slide transition effect, which uses a set of radial edges and wipes them in the
    /// clockwise direction until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “wheel” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wheel spokes="1"/>
    /// </p:transition>
    /// ```
    Wheel(WheelTransition),
    /// This element describes the wipe slide transition effect, which wipes the new slide over the previous slide from
    /// one edge of the screen to the opposite until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “wipe” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:wipe dir="d"/>
    /// </p:transition>
    /// ```
    Wipe(SideDirectionTransition),
    /// This element describes the zoom slide transition effect, which uses a box pattern centered on the slide that
    /// increases in size until the new slide is fully shown.
    ///
    /// # Xml example
    ///
    /// Consider the following cases in which the “zoom” slide transition is applied to a slide, along with a set
    /// of attributes.
    /// ```xml
    /// <p:transition>
    ///   <p:zoom dir="in"/>
    /// </p:transition>
    /// ```
    Zoom(InOutTransition),
}

impl XsdType for SlideTransitionGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "blinds" => Ok(SlideTransitionGroup::Blinds(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "checker" => Ok(SlideTransitionGroup::Checker(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "circle" => Ok(SlideTransitionGroup::Circle),
            "dissolve" => Ok(SlideTransitionGroup::Dissolve),
            "comb" => Ok(SlideTransitionGroup::Comb(OrientationTransition::from_xml_element(
                xml_node,
            )?)),
            "cover" => Ok(SlideTransitionGroup::Cover(EightDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "cut" => Ok(SlideTransitionGroup::Cut(OptionalBlackTransition::from_xml_element(
                xml_node,
            )?)),
            "diamond" => Ok(SlideTransitionGroup::Diamond),
            "fade" => Ok(SlideTransitionGroup::Fade(OptionalBlackTransition::from_xml_element(
                xml_node,
            )?)),
            "newsflash" => Ok(SlideTransitionGroup::Newsflash),
            "plus" => Ok(SlideTransitionGroup::Plus),
            "pull" => Ok(SlideTransitionGroup::Pull(EightDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "push" => Ok(SlideTransitionGroup::Push(SideDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "random" => Ok(SlideTransitionGroup::Random),
            "randomBar" => Ok(SlideTransitionGroup::RandomBar(
                OrientationTransition::from_xml_element(xml_node)?,
            )),
            "split" => Ok(SlideTransitionGroup::Split(SplitTransition::from_xml_element(
                xml_node,
            )?)),
            "strips" => Ok(SlideTransitionGroup::Strips(
                CornerDirectionTransition::from_xml_element(xml_node)?,
            )),
            "wedge" => Ok(SlideTransitionGroup::Wedge),
            "wheel" => Ok(SlideTransitionGroup::Wheel(WheelTransition::from_xml_element(
                xml_node,
            )?)),
            "wipe" => Ok(SlideTransitionGroup::Wipe(SideDirectionTransition::from_xml_element(
                xml_node,
            )?)),
            "zoom" => Ok(SlideTransitionGroup::Zoom(InOutTransition::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_SlideTransition",
            ))),
        }
    }
}

impl XsdChoice for SlideTransitionGroup {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "blinds" | "checker" | "circle" | "dissolve" | "comb" | "cover" | "cut" | "diamond" | "fade"
            | "newsflash" | "plus" | "pull" | "push" | "random" | "randomBar" | "split" | "strips" | "wedge"
            | "wheel" | "wipe" | "zoom" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TransitionStartSoundAction {
    /// This attribute specifies if the sound loops until the next sound event occurs in slideshow.
    ///
    /// Defaults to false
    pub is_looping: Option<bool>,
    /// This element specifies the audio information to play during a slide transition.
    ///
    /// # Xml example
    ///
    /// Consider a slide transition with an audio effect. The <snd> element should be used as follows:
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2" />
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    pub sound_file: EmbeddedWAVAudioFile,
}

impl TransitionStartSoundAction {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_looping = xml_node.attributes.get("loop").map(parse_xml_bool).transpose()?;

        let sound_file = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "snd")
            .map(EmbeddedWAVAudioFile::from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "snd"))?;

        Ok(Self { is_looping, sound_file })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TransitionSoundAction {
    /// This element describes the sound that starts playing during a slide transition.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2"/>
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    StartSound(TransitionStartSoundAction),
    /// This element stops all previous sounds during a slide transition.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:endSnd/>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    EndSound,
}

impl XsdType for TransitionSoundAction {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "stSnd" => Ok(TransitionSoundAction::StartSound(
                TransitionStartSoundAction::from_xml_element(xml_node)?,
            )),
            "endSnd" => Ok(TransitionSoundAction::EndSound),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_TransitionSoundAction",
            ))),
        }
    }
}

impl XsdChoice for TransitionSoundAction {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "stSnd" | "endSnd" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct SlideTransition {
    /// Specifies the transition speed that is to be used when transitioning from the current slide
    /// to the next.
    ///
    /// Defaults to TransitionSpeed::Fast
    pub speed: Option<TransitionSpeed>,
    /// Specifies whether a mouse click advances the slide or not. If this attribute is not specified
    /// then a value of true is assumed.
    ///
    /// Defaults to true
    pub advance_on_click: Option<bool>,
    /// Specifies the time, in milliseconds, after which the transition should start. This setting can
    /// be used in conjunction with the advance_on_click attribute. If this attribute is not specified then it
    /// is assumed that no auto-advance occurs.
    pub advance_on_time: Option<u32>,
    pub transition_type: Option<SlideTransitionGroup>,
    /// This element describes a sound action for slide transition. This element specifies that the start of the slide
    /// transition is accompanied by the playback of an audio file; the actual audio file used is specified by the snd
    /// element
    ///
    /// # Xml example
    ///
    /// Consider a slide transition with a sound effect. The <sndAc> element should be used as follows:
    /// ```xml
    /// <p:transition>
    ///   <p:sndAc>
    ///     <p:stSnd>
    ///       <p:snd r:embed="rId2"/>
    ///     </p:stSnd>
    ///   </p:sndAc>
    /// </p:transition>
    /// ```
    pub sound_action: Option<TransitionSoundAction>,
}

impl SlideTransition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "spd" => instance.speed = Some(value.parse()?),
                    "advClick" => instance.advance_on_click = Some(value.parse()?),
                    "advTm" => instance.advance_on_time = Some(value.parse()?),
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
                            "sndAc" => {
                                instance.sound_action = Some(TransitionSoundAction::from_xml_element(child_node)?)
                            }
                            local_name if SlideTransitionGroup::is_choice_member(local_name) => {
                                instance.transition_type = Some(SlideTransitionGroup::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct SlideTiming {
    /// This element specifies a list of time node elements used in an animation sequence.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par> … </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    pub time_node_list: Option<Vec<TimeNodeGroup>>,
    /// This element specifies the list of graphic elements to build. This refers to how the different sub-shapes or sub-
    /// components of a object are displayed. The different objects that can have build properties are text, diagrams,
    /// and charts.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="1" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    pub build_list: Option<Vec<Build>>,
}

impl SlideTiming {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "tnLst" => {
                        let vec = child_node
                            .child_nodes
                            .iter()
                            .filter(|tn_node| tn_node.local_name() == "tn")
                            .map(TimeNodeGroup::from_xml_element)
                            .collect::<Result<Vec<_>>>()?;

                        instance.time_node_list = if !vec.is_empty() {
                            Some(vec)
                        } else {
                            return Err(Box::<dyn Error>::from(MissingChildNodeError::new(
                                child_node.name.clone(),
                                "tn",
                            )));
                        }
                    }
                    "bldLst" => {
                        let vec = child_node
                            .child_nodes
                            .iter()
                            .filter(|bld_node| bld_node.local_name() == "bld")
                            .map(Build::from_xml_element)
                            .collect::<Result<Vec<_>>>()?;

                        instance.build_list = if !vec.is_empty() {
                            Some(vec)
                        } else {
                            return Err(Box::<dyn Error>::from(MissingChildNodeError::new(
                                child_node.name.clone(),
                                "bld",
                            )));
                        }
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct HeaderFooter {
    /// Specifies whether the slide number placeholder is enabled. If this attribute is not
    /// specified, a value of true should be assumed by the generating application.
    pub slide_number_enabled: Option<bool>, // true
    /// Specifies whether the Header placeholder is enabled for this master. If this attribute is
    /// not specified, a value of true should be assumed by the generating application.
    pub header_enabled: Option<bool>,
    /// Specifies whether the Footer placeholder is enabled for this master. If this attribute is not
    /// specified, a value of true should be assumed by the generating application.
    pub footer_enabled: Option<bool>,
    /// Specifies whether the Date/Time placeholder is enabled for this master. If this attribute is
    /// not specified, a value of true should be assumed by the generating application.
    pub date_time_enabled: Option<bool>,
}

impl HeaderFooter {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "sldNum" => instance.slide_number_enabled = Some(parse_xml_bool(value)?),
                    "hdr" => instance.header_enabled = Some(parse_xml_bool(value)?),
                    "ftr" => instance.footer_enabled = Some(parse_xml_bool(value)?),
                    "dt" => instance.date_time_enabled = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Control {
    pub picture: Option<Box<Picture>>,
    pub ole_attributes: Box<OleAttributes>,
}

impl Control {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            instance.ole_attributes.try_attribute_parse(attr, value)?;
        }

        instance.picture = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pic")
            .map(Picture::from_xml_element)
            .transpose()?
            .map(Box::new);

        Ok(instance)
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct OleAttributes {
    pub shape_id: Option<ShapeId>,
    /// Specifies the identifying name class used by scripting languages. This name is also used to
    /// construct the clipboard name.
    pub name: Option<String>,
    /// Specifies whether the Embedded object shows as an icon or using its native representation.
    ///
    /// Defaults to false
    pub show_as_icon: Option<bool>,
    /// Specifies the relationship id that is used to identify this Embedded object from within a slide.
    pub id: Option<RelationshipId>,
    /// Specifies the width of the embedded control.
    pub image_width: Option<PositiveCoordinate32>,
    /// Specifies the height of the embedded control.
    pub image_height: Option<PositiveCoordinate32>,
}

impl OleAttributes {
    pub fn try_attribute_parse<T: AsRef<str>>(&mut self, attr: T, value: T) -> Result<()> {
        match attr.as_ref() {
            "spid" => self.shape_id = Some(value.as_ref().parse()?),
            "name" => self.name = Some(value.as_ref().to_string()),
            "showAsIcon" => self.show_as_icon = Some(parse_xml_bool(value)?),
            "r:id" => self.id = Some(value.as_ref().to_string()),
            "imgW" => self.image_width = Some(value.as_ref().parse()?),
            "imgH" => self.image_height = Some(value.as_ref().parse()?),
            _ => (),
        }

        Ok(())
    }
}
