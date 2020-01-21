use super::util::XmlNodeExt;
use crate::{
    shared::drawingml::{
        audiovideo::EmbeddedWAVAudioFile,
        colors::Color,
        core::{AnimationElementChoice, AnimationGraphicalObjectBuildProperties},
        simpletypes::{
            Angle, DrawingElementId, FixedPercentage, Percentage, PositiveFixedPercentage, PositivePercentage, ShapeId,
        },
    },
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError, ParseEnumError},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::{error::Error, str::FromStr};

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

/// This simple type defines the position of an object in an ordered list.
pub type Index = u32;
/// This simple type represents a node or event on the timeline by its identifier.
pub type TLTimeNodeId = u32;
pub type TLSubShapeId = ShapeId;

/// This simple type defines an animation target element that is represented by a subelement of a chart.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLChartSubelementType {
    #[strum(serialize = "gridLegend")]
    GridLegend,
    #[strum(serialize = "series")]
    Series,
    #[strum(serialize = "category")]
    Category,
    #[strum(serialize = "ptInSeries")]
    PointInSeries,
    #[strum(serialize = "ptInCategory")]
    PointInCategory,
}

/// This simple type describes how to build a paragraph.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLParaBuildType {
    /// Specifies to animate all paragraphs at once.
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
    /// Specifies to animate paragraphs grouped by bullet level.
    #[strum(serialize = "p")]
    Paragraph,
    /// Specifies the build has custom user settings.
    #[strum(serialize = "cust")]
    Custom,
    /// Specifies to animate the entire body of text as one block.
    #[strum(serialize = "whole")]
    Whole,
}

/// This simple type specifies the different diagram build types.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLDiagramBuildType {
    #[strum(serialize = "whole")]
    Whole,
    #[strum(serialize = "depthByNode")]
    DepthByNode,
    #[strum(serialize = "depthByBranch")]
    DepthByBranch,
    #[strum(serialize = "breadthByNode")]
    BreadthByNode,
    #[strum(serialize = "breadthByLvl")]
    BreadthByLevel,
    #[strum(serialize = "cw")]
    Clockwise,
    #[strum(serialize = "cwIn")]
    ClockwiseIn,
    #[strum(serialize = "cwOut")]
    ClockwiseOut,
    #[strum(serialize = "ccw")]
    CounterClockwise,
    #[strum(serialize = "ccwIn")]
    CounterClockwiseIn,
    #[strum(serialize = "ccwOut")]
    CounterClockwiseOut,
    #[strum(serialize = "inByRing")]
    InByRing,
    #[strum(serialize = "outByRing")]
    OutByRing,
    #[strum(serialize = "up")]
    Up,
    #[strum(serialize = "down")]
    Down,
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
    #[strum(serialize = "cust")]
    Custom,
}

/// This simple type describes how to build an embedded Chart.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLOleChartBuildType {
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
    #[strum(serialize = "series")]
    Series,
    #[strum(serialize = "category")]
    Category,
    #[strum(serialize = "seriesEl")]
    SeriesElement,
    #[strum(serialize = "categoryEl")]
    CategoryElement,
}

/// This simple type specifies the child time node that triggers a time condition. References a child TimeNode or all
/// child nodes. Order is based on the child's end time.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTriggerRuntimeNode {
    #[strum(serialize = "first")]
    First,
    #[strum(serialize = "last")]
    Last,
    #[strum(serialize = "all")]
    All,
}

/// This simple type specifies a particular event that causes the time condition to be true.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTriggerEvent {
    /// Fire trigger at the beginning
    #[strum(serialize = "onBegin")]
    OnBegin,
    /// Fire trigger at the end
    #[strum(serialize = "onEnd")]
    OnEnd,
    /// Fire trigger at the beginning
    #[strum(serialize = "begin")]
    Begin,
    /// Fire trigger at the end
    #[strum(serialize = "end")]
    End,
    /// Fire trigger on a mouse click
    #[strum(serialize = "onClick")]
    OnClick,
    /// Fire trigger on double-mouse click
    #[strum(serialize = "onDblClick")]
    OnDoubleClick,
    /// Fire trigger on mouse over
    #[strum(serialize = "onMouseOver")]
    OnMouseOver,
    /// Fire trigger on mouse out
    #[strum(serialize = "onMouseOut")]
    OnMouseOut,
    /// Fire trigger on next node
    #[strum(serialize = "onNext")]
    OnNext,
    /// Fire trigger on previous node
    #[strum(serialize = "onPrev")]
    OnPrev,
    /// Fire trigger on stop audio
    #[strum(serialize = "onStopAudio")]
    OnStopAudio,
}

/// This simple type specifies how the animation is applied over subelements of the target element.
#[repr(C)]
#[derive(Debug, Copy, Clone, PartialEq, EnumString)]
pub enum IterateType {
    /// Iterate by element.
    #[strum(serialize = "el")]
    Element,
    /// Iterate by Letter.
    #[strum(serialize = "wd")]
    Word,
    /// Iterate by Word.
    #[strum(serialize = "lt")]
    Letter,
}

/// This simple type specifies the class of effect in which this effect belongs.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodePresetClassType {
    #[strum(serialize = "entr")]
    Entrance,
    #[strum(serialize = "exit")]
    Exit,
    #[strum(serialize = "emph")]
    Emphasis,
    #[strum(serialize = "path")]
    Path,
    #[strum(serialize = "verb")]
    Verb,
    #[strum(serialize = "mediacall")]
    Mediacall,
}

/// This simple type determines whether an effect can play more than once.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodeRestartType {
    /// Always restart node
    #[strum(serialize = "always")]
    Always,
    /// Restart when node is not active
    #[strum(serialize = "whenNotActive")]
    WhenNotActive,
    /// Never restart node
    #[strum(serialize = "never")]
    Never,
}

/// This simple type specifies what modifications the effect leaves on the target element's properties when the
/// effect ends.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodeFillType {
    #[strum(serialize = "remove")]
    Remove,
    #[strum(serialize = "freeze")]
    Freeze,
    #[strum(serialize = "hold")]
    Hold,
    #[strum(serialize = "transition")]
    Transition,
}

/// This simple type specifies how the time node synchronizes to its group.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodeSyncType {
    #[strum(serialize = "canSlip")]
    CanSlip,
    #[strum(serialize = "locked")]
    Locked,
}

/// This simple type specifies how the time node plays back relative to its master time node.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodeMasterRelation {
    #[strum(serialize = "sameClick")]
    SameClick,
    #[strum(serialize = "lastClick")]
    LastClick,
    #[strum(serialize = "nextClick")]
    NextClick,
}

/// This simple type specifies time node types.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLTimeNodeType {
    #[strum(serialize = "clickEffect")]
    ClickEffect,
    #[strum(serialize = "withEffect")]
    WithEffect,
    #[strum(serialize = "afterEffect")]
    AfterEffect,
    #[strum(serialize = "mainSequence")]
    MainSequence,
    #[strum(serialize = "interactiveSeq")]
    InteractiveSequence,
    #[strum(serialize = "clickPar")]
    ClickParagraph,
    #[strum(serialize = "withGroup")]
    WithGroup,
    #[strum(serialize = "afterGroup")]
    AfterGroup,
    #[strum(serialize = "tmRoot")]
    TimingRoot,
}

/// This simple type specifies what to do when going forward in a sequence. When the value is Seek, it seeks the
/// current child element to its natural end time before advancing to the next element.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLNextActionType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "seek")]
    Seek,
}

/// This simple type specifies what to do when going backwards in a sequence. When the value is SkipTimed, the
/// sequence continues to go backwards until it reaches a sequence element that was defined to being only on a
/// "next" event.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLPreviousActionType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "skipTimed")]
    SkipTimed,
}

/// This simple type specifies how the animation flows from point to point.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateBehaviorCalcMode {
    #[strum(serialize = "discrete")]
    Discrete,
    #[strum(serialize = "fmla")]
    Formula,
    #[strum(serialize = "lin")]
    Linear,
}

/// This simple type specifies the type of property value.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateBehaviorValueType {
    #[strum(serialize = "clr")]
    Color,
    #[strum(serialize = "num")]
    Number,
    #[strum(serialize = "str")]
    String,
}

/// This simple type specifies how to apply the animation values to the original value for the property.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLBehaviorAdditiveType {
    #[strum(serialize = "base")]
    Base,
    #[strum(serialize = "sum")]
    Sum,
    #[strum(serialize = "repl")]
    Replace,
    #[strum(serialize = "mult")]
    Multiply,
    #[strum(serialize = "none")]
    None,
}

/// This simple type makes a repeating animation build with each iteration when set to "always."
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLBehaviorAccumulateType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "always")]
    Always,
}

/// This simple type specifies how the behavior animates the target element.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLBehaviorTransformType {
    #[strum(serialize = "pt")]
    Point,
    #[strum(serialize = "img")]
    Image,
}

/// This simple type specifies how a behavior should override values of the attribute being animated on the target
/// element. The ChildStyle clears the attributes on the children contained inside the target element.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLBehaviorOverrideType {
    #[strum(serialize = "normal")]
    Normal,
    #[strum(serialize = "childStyle")]
    ChildStyle,
}

/// This simple type specifies the color space of the animation.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateColorSpace {
    #[strum(serialize = "rgb")]
    Rgb,
    #[strum(serialize = "hsl")]
    Hsl,
}

/// This simple type specifies the direction in which to interpolate the animation (clockwise or counterclockwise).
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateColorDirection {
    #[strum(serialize = "cw")]
    Clockwise,
    #[strum(serialize = "ccw")]
    CounterClockwise,
}

/// This simple type specifies whether the effect is a transition in, transition out, or neither.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateEffectTransition {
    #[strum(serialize = "in")]
    In,
    #[strum(serialize = "out")]
    Out,
    #[strum(serialize = "none")]
    None,
}

/// This simple type specifies what the origin of the motion path is relative to.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateMotionBehaviorOrigin {
    #[strum(serialize = "parent")]
    Parent,
    #[strum(serialize = "layout")]
    Layout,
}

/// This simple type specifies how the motion path moves when the target element is moved.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLAnimateMotionPathEditMode {
    #[strum(serialize = "relative")]
    Relative,
    #[strum(serialize = "fixed")]
    Fixed,
}

/// This simple type specifies a command type.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum TLCommandType {
    #[strum(serialize = "evt")]
    Event,
    #[strum(serialize = "call")]
    Call,
    #[strum(serialize = "verb")]
    Verb,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct IndexRange {
    /// This attribute defines the start of the index range.
    pub start: Index,
    /// This attribute defines the end of the index range.
    pub end: Index,
}

impl IndexRange {
    pub fn new(start: Index, end: Index) -> Self {
        Self { start, end }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut start = None;
        let mut end = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "st" => start = Some(value.parse()?),
                "end" => end = Some(value.parse()?),
                _ => (),
            }
        }

        let start = start.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "st"))?;
        let end = end.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "end"))?;

        Ok(Self { start, end })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TimeNodeGroup {
    /// This element describes the Parallel time node which can be activated along with other parallel time node
    /// containers.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par>
    ///       <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    ///         <p:childTnLst>
    ///           <p:seq concurrent="1" nextAc="seek">
    ///             …
    ///           </p:seq>
    ///         </p:childTnLst>
    ///       </p:cTn>
    ///     </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    Parallel(Box<TLCommonTimeNodeData>),
    /// This element describes the Sequence time node and it can only be activated when the one before it finishes.
    ///
    /// # Xml example
    ///
    /// For example, suppose we have a simple animation with a blind entrance.
    /// ```xml
    /// <p:timing>
    ///   <p:tnLst>
    ///     <p:par>
    ///       <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
    ///         <p:childTnLst>
    ///           <p:seq concurrent="1" nextAc="seek">
    ///           …
    ///           </p:seq>
    ///         </p:childTnLst>
    ///       </p:cTn>
    ///     </p:par>
    ///   </p:tnLst>
    /// </p:timing>
    /// ```
    Sequence(Box<TLTimeNodeSequence>),
    /// This element describes the Exclusive time node. This time node is used to pause all other timelines when it is
    /// activated.
    Exclusive(Box<TLCommonTimeNodeData>),
    /// This element is a generic animation element that requires little or no semantic understanding of the attribute
    /// being animated. It can animate text within a shape or even the shape itself.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize text within a shape by changing the size of its font by 150%. The
    /// <anim> element should be used as follows:
    ///
    /// ```xml
    /// <p:anim to="1.5" calcmode="lin" valueType="num">
    ///   <p:cBhvr override="childStyle">
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1">
    ///         <p:txEl>
    ///           <p:charRg st="1" end="4"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>style.fontSize</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:anim>
    /// ```
    Animate(Box<TLAnimateBehavior>),
    /// This animation element is responsible for animating the color of an object.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize a shape by changing its fill color to scheme color accent2. The
    /// <animClr> element should be used as follows:
    /// ```xml
    /// <p:animClr clrSpc="rgb">
    ///   <p:cBhvr>
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>fillcolor</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:to>
    ///     <a:schemeClr val="accent2"/>
    ///   </p:to>
    /// </p:animClr>
    /// ```
    AnimateColor(Box<TLAnimateColorBehavior>),
    /// This animation behavior provides the ability to do image transform/filter effects on elements. Some visual
    /// effects are dynamic in nature and have a progress that animates from 0 to 1 over a period of time to do visual
    /// transitions between hidden and visible states. Other filters are static and apply a effects like a blur or drop-
    /// shadow which aren't inherently time-based.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize a shape by creating an entrance animation using a "blinds" motion.
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    AnimateEffect(Box<TLAnimateEffectBehavior>),
    /// Animate motion provides an abstracted way to move positioned elements. It provides the ability to specify
    /// from/to/by motion as well as to use more detailed path descriptions for motion over polylines or bezier curves.
    ///
    /// # Xml example
    ///
    /// Consider animating a shape from its original position to the right.. The <animMotion> element should
    /// be used as follows:
    ///
    /// ```xml
    /// <p:animMotion origin="layout" path="M 0 0 L 0.25 0 E" pathEditMode="relative">
    ///   <p:cBhvr>
    ///     <p:cTn id="1" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="1"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animMotion>
    /// ```
    AnimateMotion(Box<TLAnimateMotionBehavior>),
    /// This animation element is responsible for animating the rotation of an object. Rotation values set in the "by",
    /// "to, and "from" attributes are specified in degrees measured to a 60,000th, i.e 1 degree is 60,000. Rotation
    /// values can be larger than 360°.
    ///
    /// The sign of the rotation angle specifies the direction for rotation. A negative rotation specifies that the rotation
    /// should appear in the host to go counter-clockwise".
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize a shape by rotating it 360 degrees clockwise. The <animRot> element
    /// should be used as follows:
    ///
    /// ```xml
    /// <p:animRot by="21600000">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="5"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>r</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animRot>
    /// ```
    AnimateRotation(Box<TLAnimateRotationBehavior>),
    /// This animation element is responsible for animating the scale of an object. When animating the scale, the
    /// element shall scale around the reference point of the element and the positioning system used should be
    /// consistent with the one used for motion paths. When animating the width and height of an element, all of the
    /// width/height animation values are calculated first then the scale animations are applied on top of that. So for
    /// example, an animation from 0 to 100 of the width with a concurrent scale from 100% to 200% would result in
    /// the element appearing to scale from 0 to 200.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize a shape by scaling it larger by 150%. The <animScale> element should
    /// be used as follows:
    ///
    /// ```xml
    /// <p:childTnLst>
    ///   <p:animScale>
    ///     <p:cBhvr>
    ///       <p:cTn id="6" dur="2000" fill="hold"/>
    ///       <p:tgtEl>
    ///         <p:spTgt spid="5"/>
    ///       </p:tgtEl>
    ///     </p:cBhvr>
    ///     <p:by x="150000" y="150000"/>
    ///   </p:animScale>
    /// </p:childTnLst>
    /// ```
    AnimateScale(Box<TLAnimateScaleBehavior>),
    /// This element describes the several non-durational commands that can be executed within a timeline. This can
    /// be used to send events, call functions on elements, and send verbs to embedded objects. For example “Object
    /// Action” effects for Embedded objects and Media commands for sounds/movies such as "PlayFrom(0.0)" and
    /// "togglePause".
    Command(Box<TLCommandBehavior>),
    /// This element allows the setting of a particular property value to a fixed value while the behavior is active and
    /// restores the value when the behavior is reset or turned off.
    ///
    /// # Xml example
    ///
    /// For example, suppose we want to set certain properties during an animation effect. The <set>
    /// element should be used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:set>
    ///     <p:cBhvr>
    ///       <p:cTn id="6" dur="1" fill="hold"> … </p:cTn>
    ///       <p:tgtEl>
    ///         <p:spTgt spid="4"/>
    ///       </p:tgtEl>
    ///       <p:attrNameLst>
    ///         <p:attrName>style.visibility</p:attrName>
    ///       </p:attrNameLst>
    ///     </p:cBhvr>
    ///     <p:to>
    ///       <p:strVal val="visible"/>
    ///     </p:to>
    ///   </p:set>
    ///   <p:animEffect transition="in" filter="blinds(horizontal)">
    ///     …
    ///   </p:animEffect>
    /// </p:childTnLst>
    /// ```
    Set(Box<TLSetBehavior>),
    /// This element is used to include audio during an animation. This element specifies that this node within the
    /// animation tree triggers the playback of an audio file; the actual audio file used is specified by the sndTgt
    /// element (§19.5.70).
    ///
    /// # Xml example
    ///
    /// Consider adding applause sound to an animation sequence. The audio element is used as follows:
    ///
    /// ```xml
    /// <p:cTn ...>
    ///   <p:stCondLst>...</p:stCondLst>
    ///   <p:childTnLst>...</p:childTnLst>
    ///   <p:subTnLst>
    ///     <p:audio>
    ///       <p:cMediaNode vol="50%">...
    ///         <p:tgtEl>
    ///           <p:sndTgt r:embed="rId2" />
    ///         </p:tgtEl>
    ///       </p:cMediaNode>
    ///     </p:audio>
    ///   </p:subTnLst>
    /// </p:cTn>
    /// ```
    ///
    /// The audio element specifies the location of the audio playback within the animation; its child sndTgt element
    /// specifies that the audio to be played is the target of the relationship with ID rId2.
    Audio(Box<TLMediaNodeAudio>),
    /// This element specifies video information in an animation sequence. This element specifies that this node within
    /// the animation tree triggers the playback of a video file; the actual video file used is specified by the videoFile
    /// element
    ///
    /// # Xml example
    ///
    /// Consider a slide with an animated video content. The <video> element is used as follows:
    /// ```xml
    /// <p:cSld>
    ///   <p:spTree>
    ///     <p:pic>
    ///       <p:nvPicPr>
    ///       <p:cNvPr id="4"/>
    ///       …
    ///       <p:nvPr>
    ///         <a:videoFile r:link="rId1" contentType="video/ogg"/>
    ///       </p:nvPr>
    ///     </p:nvPicPr>
    ///     …
    ///     </p:pic>
    ///   </p:spTree>
    /// </p:cSld>
    /// …
    /// <p:childTnLst>
    ///   <p:seq concurrent="1" nextAc="seek">
    ///     …
    ///   </p:seq>
    ///   <p:video>
    ///     <p:cMediaNode>
    ///       …
    ///       <p:tgtEl>
    ///         <p:spTgt spid="4"/>
    ///       </p:tgtEl>
    ///     </p:cMediaNode>
    ///   </p:video>
    /// </p:childTnLst>
    /// ```
    ///
    /// The video element specifies the location of the video playback within the animation sequence; its child spTgt
    /// element specifies that the shape which contains the video to be played has a shape ID of 4. If we look at the
    /// shape with that ID value, its child videoFile element references an external video file of content type video/ogg
    /// located at the target of the relationship with ID rId1
    Video(Box<TLMediaNodeVideo>),
}

impl XsdType for TimeNodeGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "par" => Ok(TimeNodeGroup::Parallel(Box::new(
                TLCommonTimeNodeData::from_xml_element(xml_node)?,
            ))),
            "seq" => Ok(TimeNodeGroup::Sequence(Box::new(TLTimeNodeSequence::from_xml_element(
                xml_node,
            )?))),
            "excl" => Ok(TimeNodeGroup::Exclusive(Box::new(
                TLCommonTimeNodeData::from_xml_element(xml_node)?,
            ))),
            "anim" => Ok(TimeNodeGroup::Animate(Box::new(TLAnimateBehavior::from_xml_element(
                xml_node,
            )?))),
            "animClr" => Ok(TimeNodeGroup::AnimateColor(Box::new(
                TLAnimateColorBehavior::from_xml_element(xml_node)?,
            ))),
            "animEffect" => Ok(TimeNodeGroup::AnimateEffect(Box::new(
                TLAnimateEffectBehavior::from_xml_element(xml_node)?,
            ))),
            "animMotion" => Ok(TimeNodeGroup::AnimateMotion(Box::new(
                TLAnimateMotionBehavior::from_xml_element(xml_node)?,
            ))),
            "animRot" => Ok(TimeNodeGroup::AnimateRotation(Box::new(
                TLAnimateRotationBehavior::from_xml_element(xml_node)?,
            ))),
            "animScale" => Ok(TimeNodeGroup::AnimateScale(Box::new(
                TLAnimateScaleBehavior::from_xml_element(xml_node)?,
            ))),
            "cmd" => Ok(TimeNodeGroup::Command(Box::new(TLCommandBehavior::from_xml_element(
                xml_node,
            )?))),
            "set" => Ok(TimeNodeGroup::Set(Box::new(TLSetBehavior::from_xml_element(xml_node)?))),
            "audio" => Ok(TimeNodeGroup::Audio(Box::new(TLMediaNodeAudio::from_xml_element(
                xml_node,
            )?))),
            "video" => Ok(TimeNodeGroup::Video(Box::new(TLMediaNodeVideo::from_xml_element(
                xml_node,
            )?))),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TimeNodeGroup",
            ))),
        }
    }
}

impl XsdChoice for TimeNodeGroup {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "par" | "seq" | "excl" | "anim" | "animClr" | "animEffect" | "animMotion" | "animRot" | "animScale"
            | "cmd" | "set" | "audio" | "video" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TLTimeNodeList(pub Vec<TimeNodeGroup>);

impl TLTimeNodeList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let vec = xml_node
            .child_nodes
            .iter()
            .filter_map(TimeNodeGroup::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(vec))
    }
}

/// This element describes the common behaviors of animations.
///
/// # Xml example
///
/// ```xml
/// <p:anim to="1.5" calcmode="lin" valueType="num">
///   <p:cBhvr override="childStyle">
///     <p:cTn id="6" dur="2000" fill="hold"/>
///     <p:tgtEl>
///       <p:spTgt spid="3">
///         <p:txEl>
///           <p:charRg st="4294967295" end="4294967295"/>
///         </p:txEl>
///       </p:spTgt>
///     </p:tgtEl>
///     <p:attrNameLst>
///       <p:attrName>style.fontSize</p:attrName>
///     </p:attrNameLst>
///   </p:cBhvr>
/// </p:anim>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct TLCommonBehaviorData {
    pub additive: Option<TLBehaviorAdditiveType>,
    pub accumulate: Option<TLBehaviorAccumulateType>,
    pub transform_type: Option<TLBehaviorTransformType>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub by: Option<String>,
    pub runtime_context: Option<String>,
    pub override_type: Option<TLBehaviorOverrideType>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    /// This element specifies the target children elements which have the animation effects applied to.
    ///
    /// # Xml example
    ///
    /// Consider a shape with ID 3 with a fade effect animation applied to it. The <tgtEl> element should be
    /// used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="fade">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="2000"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    pub target_element: TLTimeTargetElement,
    /// This element is used to describe a list of attributes in which to apply an animation to.
    ///
    /// The elements of this list is used to contain an attribute value for an Attribute Name List. This value defines the specific
    /// attribute that an animation should be applied to, such as fill, style, and shadow, etc. A specific property is
    /// defined by using a "property.sub-property" format which is often extended to multiple sub properties as seen in
    /// the allowed values below.
    ///
    /// Allowed property values:
    /// * style.opacity
    /// * style.rotation
    /// * style.visibility
    /// * style.color
    /// * style.fontSize
    /// * style.fontWeight
    /// * style.fontStyle
    /// * style.fontFamily
    /// * style.textEffectEmboss
    /// * style.textShadow
    /// * style.textTransform
    /// * style.textDecorationUnderline
    /// * style.textEffectOutline
    /// * style.textDecorationLineThrough
    /// * style.sRotation
    /// * imageData.cropTop
    /// * imageData.cropBottom
    /// * imageData.cropLeft
    /// * imageData.cropRight
    /// * imageData.gain
    /// * imageData.blacklevel
    /// * imageData.gamma
    /// * imageData.grayscale
    /// * imageData.chromakey
    /// * fill.on
    /// * fill.type
    /// * fill.color
    /// * fill.opacity
    /// * fill.color2
    /// * fill.method
    /// * fill.opacity2
    /// * fill.angle
    /// * fill.focus
    /// * fill.focusposition.x
    /// * fill.focusposition.y
    /// * fill.focussize.x
    /// * fill.focussize.y
    /// * stroke.on
    /// * stroke.color
    /// * stroke.weight
    /// * stroke.opacity
    /// * stroke.linestyle
    /// * stroke.dashstyle
    /// * stroke.filltype
    /// * stroke.src
    /// * stroke.color2
    /// * stroke.imagesize.x
    /// * stroke.imagesize.y
    /// * stroke.startArrow
    /// * stroke.endArrow
    /// * stroke.startArrowWidth
    /// * stroke.startArrowLength
    /// * stroke.endArrowWidth
    /// * stroke.endArrowLength
    /// * shadow.on
    /// * shadow.type
    /// * shadow.color
    /// * shadow.color2
    /// * shadow.opacity
    /// * shadow.offset.x
    /// * shadow.offset.y
    /// * shadow.offset2.x
    /// * shadow.offset2.y
    /// * shadow.origin.x
    /// * shadow.origin.y
    /// * shadow.matrix.xtox
    /// * shadow.matrix.ytox
    /// * shadow.matrix.xtoy
    /// * shadow.matrix.ytoy
    /// * shadow.matrix.perspectiveX
    /// * shadow.matrix.perspectiveY
    /// * skew.on
    /// * skew.offset.x
    /// * skew.offset.y
    /// * skew.origin.x
    /// * skew.origin.y
    /// * skew.matrix.xtox
    /// * skew.matrix.ytox
    /// * skew.matrix.xtoy
    /// * skew.matrix.ytoy
    /// * skew.matrix.perspectiveX
    /// * skew.matrix.perspectiveY
    /// * extrusion.on
    /// * extrusion.type
    /// * extrusion.render
    /// * extrusion.viewpointorigin.x
    /// * extrusion.viewpointorigin.y
    /// * extrusion.viewpoint.x
    /// * extrusion.viewpoint.y
    /// * extrusion.viewpoint.z
    /// * extrusion.plane
    /// * extrusion.skewangle
    /// * extrusion.skewamt
    /// * extrusion.backdepth,
    /// * extrusion.foredepth
    /// * extrusion.orientation.x
    /// * extrusion.orientation.y
    /// * extrusion.orientation.zand
    /// * extrusion.orientationangle
    /// * extrusion.color,
    /// * extrusion.rotationangle.x
    /// * extrusion.rotationangle.y
    /// * extrusion.lockrotationcenter
    /// * extrusion.autorotationcenter
    /// * extrusion.rotationcenter.x
    /// * extrusion.rotationcenter.y
    /// * extrusion.rotationcenter.z
    /// * extrusion.colormode.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize the txt font size within the body of a shape. The attribute would be
    /// 'style.fontSize' and this can be done by doing the following:
    ///
    /// ```xml
    /// <p:anim to="1.5" calcmode="lin" valueType="num">
    ///   <p:cBhvr override="childStyle">
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:charRg st="4294967295" end="4294967295"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>style.fontSize</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:anim>
    /// ```
    pub attr_name_list: Option<Vec<String>>,
}

impl TLCommonBehaviorData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut additive = None;
        let mut accumulate = None;
        let mut transform_type = None;
        let mut from = None;
        let mut to = None;
        let mut by = None;
        let mut runtime_context = None;
        let mut override_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "additive" => additive = Some(value.parse()?),
                "accumulate" => accumulate = Some(value.parse()?),
                "xfrmType" => transform_type = Some(value.parse()?),
                "from" => from = Some(value.clone()),
                "to" => to = Some(value.clone()),
                "by" => by = Some(value.clone()),
                "rctx" => runtime_context = Some(value.clone()),
                "override" => override_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut target_element = None;
        let mut attr_name_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "tgtEl" => {
                    target_element = Some(
                        child_node
                            .child_nodes
                            .first()
                            .ok_or_else(|| {
                                MissingChildNodeError::new(child_node.name.clone(), "sldTgt|sndTgt|spTgt|inkTgt").into()
                            })
                            .and_then(TLTimeTargetElement::from_xml_element)?,
                    );
                }
                "attrNameLst" => {
                    let vec = child_node
                        .child_nodes
                        .iter()
                        .filter(|attr_name_node| attr_name_node.local_name() == "attrName")
                        .map(|attr_name_node| attr_name_node.text.as_ref().cloned().unwrap_or_default())
                        .collect::<Vec<_>>();

                    if vec.is_empty() {
                        return Err(Box::new(LimitViolationError::new(
                            child_node.name.clone(),
                            "attrName",
                            1,
                            MaxOccurs::Unbounded,
                            0,
                        )));
                    }

                    attr_name_list = Some(vec);
                }
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;
        let target_element =
            target_element.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tgtEl"))?;

        Ok(Self {
            additive,
            accumulate,
            transform_type,
            from,
            to,
            by,
            runtime_context,
            override_type,
            common_time_node_data,
            target_element,
            attr_name_list,
        })
    }
}

/// This element is used to describe behavior of media elements, such as sound or movies, in an animation.
///
/// # Xml example
///
/// ```xml
/// <p:audio>
///   <p:cMediaNode mute="1">
///     <p:cTn display="0" masterRel="sameClick">
///       <p:stCondLst> … </p:stCondLst>
///       <p:endCondLst> … </p:endCondLst>
///     </p:cTn>
///     <p:tgtEl> … </p:tgtEl>
///   </p:cMediaNode>
/// </p:audio>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct TLCommonMediaNodeData {
    /// This attribute describes the volume of the media element.
    ///
    /// Defaults to 50000
    pub volume: Option<PositiveFixedPercentage>,
    /// This attribute describes whether the media should be mute.
    ///
    /// Defaults to false
    pub mute: Option<bool>,
    /// This attribute describes the numbers of slides across which the media should play.
    ///
    /// Defaults to 1
    pub number_of_slides: Option<u32>,
    /// This attribute describes whether the media should be displayed when it is stopped.
    ///
    /// Defaults to true
    pub show_when_stopped: Option<bool>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    pub target_element: TLTimeTargetElement,
}

impl TLCommonMediaNodeData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut volume = None;
        let mut mute = None;
        let mut number_of_slides = None;
        let mut show_when_stopped = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "vol" => volume = Some(value.parse()?),
                "mute" => mute = Some(parse_xml_bool(value)?),
                "numSld" => number_of_slides = Some(value.parse()?),
                "showWhenStopped" => show_when_stopped = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut target_element = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "tgtEl" => {
                    target_element = Some(
                        child_node
                            .child_nodes
                            .first()
                            .ok_or_else(|| {
                                MissingChildNodeError::new(child_node.name.clone(), "TLTimeTargetElement").into()
                            })
                            .and_then(TLTimeTargetElement::from_xml_element)?,
                    )
                }
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;
        let target_element =
            target_element.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tgtEl"))?;

        Ok(Self {
            volume,
            mute,
            number_of_slides,
            show_when_stopped,
            common_time_node_data,
            target_element,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLBuildParagraph {
    pub build_common: TLBuildCommonAttributes,
    /// This attribute describe the build types.
    ///
    /// Defaults to TLParaBuildType::Whole
    pub build_type: Option<TLParaBuildType>,
    /// This attribute describes the build level for the paragraph. It is only supported in
    /// paragraph type builds i.e the build attribute shall also be set to "byParagraph" for this
    /// attribute to apply.
    ///
    /// Defaults to 1
    pub build_level: Option<u32>,
    /// This attribute indicates whether to animate the background of the shape associated with
    /// the text.
    ///
    /// Defaults to false
    pub animate_bg: Option<bool>,
    /// This attribute indicates whether to automatically update the "animateBg" setting to true
    /// when the shape associated with the text has a fill or line.
    ///
    /// Defaults to true
    pub auto_update_anim_bg: Option<bool>,
    /// This attribute is only supported in paragraph type builds. This specifies the direction of
    /// the build relative to the order of the elements in the container. When this is set to "true",
    /// the animations for the paragraphs are persisted in reverse order to the order of the
    /// paragraphs themselves such that the last paragraph animates first.
    ///
    /// Defaults to false
    pub reverse: Option<bool>,
    /// This attribute specifies time after which to automatically advance the build to the next
    /// step.
    ///
    /// Defaults to TLTime::Indefinite
    pub auto_advance_time: Option<TLTime>,
    pub template_list: Option<TLTemplateList>,
}

impl TLBuildParagraph {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;
        let mut build_level = None;
        let mut animate_bg = None;
        let mut auto_update_anim_bg = None;
        let mut reverse = None;
        let mut auto_advance_time = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "build" => build_type = Some(value.parse()?),
                "bldLvl" => build_level = Some(value.parse()?),
                "animBg" => animate_bg = Some(parse_xml_bool(value)?),
                "autoUpdateAnimBg" => auto_update_anim_bg = Some(parse_xml_bool(value)?),
                "rev" => reverse = Some(parse_xml_bool(value)?),
                "advAuto" => auto_advance_time = Some(value.parse()?),
                _ => (),
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        let template_list = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tmplLst")
            .map(TLTemplateList::from_xml_element)
            .transpose()?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_type,
            build_level,
            animate_bg,
            auto_update_anim_bg,
            reverse,
            auto_advance_time,
            template_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLPoint {
    /// This attribute describes the X coordinate.
    pub x: Percentage,
    /// This attribute describes the Y coordinate.
    pub y: Percentage,
}

impl TLPoint {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut x = None;
        let mut y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "x" => x = Some(value.parse()?),
                "y" => y = Some(value.parse()?),
                _ => (),
            }
        }

        let x = x.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "x"))?;
        let y = y.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "y"))?;

        Ok(Self { x, y })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLTime {
    TimePoint(u32),
    Indefinite,
}

impl FromStr for TLTime {
    type Err = ParseEnumError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "indefinite" => Ok(TLTime::Indefinite),
            _ => Ok(TLTime::TimePoint(s.parse().map_err(|_| Self::Err::new("TLTime"))?)),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TLTemplate {
    /// This attribute describes the paragraph indent level to which this template effect applies.
    ///
    /// Defaults to 0
    pub level: Option<u32>,
    pub time_node_list: TLTimeNodeList,
}

impl TLTemplate {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let level = xml_node.attributes.get("lvl").map(|value| value.parse()).transpose()?;

        let time_node_list = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tnLst")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tnLst").into())
            .and_then(TLTimeNodeList::from_xml_element)?;

        Ok(Self { level, time_node_list })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TLTemplateList(pub Vec<TLTemplate>);

impl TLTemplateList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let vec = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "tmpl")
            .map(TLTemplate::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        match vec.len() {
            0..=9 => Ok(Self(vec)),
            len => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "tmpl",
                0,
                MaxOccurs::Value(9),
                len as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLBuildCommonAttributes {
    /// This attribute specifies the shape to which the build applies.
    pub shape_id: DrawingElementId,
    /// This attribute ties effects persisted in the animation to the build information. The
    /// attribute is used by the editor when changes to the build information are made.
    /// GroupIDs are unique for a given shape. They are not guaranteed to be unique IDs across
    /// all shapes on a slide.
    pub group_id: u32,
    /// This attribute describes the view option indicating if the build should be displayed
    /// expanded.
    ///
    /// Defaults to false
    pub ui_expand: Option<bool>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLBuildDiagram {
    pub build_common: TLBuildCommonAttributes,
    /// This attribute describes how the diagram is built. The animation animates the sub-
    /// elements in the container in the particular order defined by this attribute.
    ///
    /// Defaults to TLDiagramBuildType::Whole
    pub build_type: Option<TLDiagramBuildType>,
}

impl TLBuildDiagram {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "bld" => build_type = Some(value.parse()?),
                _ => (),
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_type,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLOleBuildChart {
    pub build_common: TLBuildCommonAttributes,
    /// This attribute describes how the diagram is built. The animation animates the sub-
    /// elements in the container in the particular order defined by this attribute.
    ///
    /// Defaults to TLOleChartBuildType::AllAtOnce
    pub build_type: Option<TLOleChartBuildType>,
    /// This attribute describes whether to animate the background of the shape.
    ///
    /// Defaults to true
    pub animate_bg: Option<bool>,
}

impl TLOleBuildChart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;
        let mut build_type = None;
        let mut animate_bg = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                "bld" => build_type = Some(value.parse()?),
                "animBg" => animate_bg = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_type,
            animate_bg,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLGraphicalObjectBuild {
    pub build_common: TLBuildCommonAttributes,
    pub build_choice: TLGraphicalObjectBuildChoice,
}

impl TLGraphicalObjectBuild {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_id = None;
        let mut group_id = None;
        let mut ui_expand = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "spid" => shape_id = Some(value.parse()?),
                "grpId" => group_id = Some(value.parse()?),
                "uiExpand" => ui_expand = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let build_choice = xml_node
            .child_nodes
            .iter()
            .find_map(TLGraphicalObjectBuildChoice::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "TLGraphicalObjectBuildChoice"))?;

        let shape_id = shape_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?;
        let group_id = group_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "grpId"))?;

        Ok(Self {
            build_common: TLBuildCommonAttributes {
                shape_id,
                group_id,
                ui_expand,
            },
            build_choice,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLGraphicalObjectBuildChoice {
    /// This element specifies in the build list to build the entire graphical object as one entity.
    ///
    /// # Xml example
    ///
    /// Consider having a graph appear as on entity as opposed to by category. The <bldAsOne> element
    /// should be used as follows:
    ///
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldAsOne/>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildAsOne,
    /// This element specifies the animation properties of a graphical object's sub-elements.
    ///
    /// # Xml example
    ///
    /// Consider applying animation to a graphical element consisting of a diagram. The <bldSub> element
    /// should be used as follows:
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldGraphic spid="5" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldDgm bld="one"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    BuildSubElements(AnimationGraphicalObjectBuildProperties),
}

impl XsdType for TLGraphicalObjectBuildChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldAsOne" => Ok(TLGraphicalObjectBuildChoice::BuildAsOne),
            "bldSub" => Ok(TLGraphicalObjectBuildChoice::BuildSubElements(
                AnimationGraphicalObjectBuildProperties::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLGraphicalObjectBuildChoice",
            ))),
        }
    }
}

impl XsdChoice for TLGraphicalObjectBuildChoice {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "bldAsOne" | "bldSub" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLTimeNodeSequence {
    /// This attribute specifies if concurrency is enabled or disabled. By default this attribute has
    /// a value of "disabled". When the value is set to "enabled", the previous element is left
    /// enabled when advancing to the next element in a sequence instead of being ended. This
    /// is only relevant for advancing via the next condition element being triggered. The only
    /// other way to advance to the next element would be to have the current element end,
    /// which implies it is no longer concurrent.
    pub concurrent: Option<bool>,
    /// This attribute specifies what to do when going backwards in a sequence. By default it is
    /// set to TLPreviousActionType::None and nothing special is done. When the value is TLPreviousActionType::SkipTimed,
    /// the sequence continues to go backwards until it reaches a sequence element that was defined to begin
    /// only on the next condition element.
    pub prev_action_type: Option<TLPreviousActionType>,
    /// This attribute specifies what to do when going forward in sequence. By default this
    /// attribute has a value of TLNextActionType::None. When this is set to seek it seeks the element to a natural
    /// end time (not necessarily the actual end time).
    ///
    /// The natural end position is defined as the latest non-infinite end time of the children. If a
    /// child loops forever, the end of its first loop is used as its "end time" for the purposes of
    /// this calculation.
    ///
    /// Some container elements can have infinite durations due to an infinite-duration child
    /// element. The engine needs to recurse down through all infinite duration containers to
    /// calculate their natural duration in case a child might have non-infinite duration within it
    /// that needs to be taken into account.
    pub next_action_type: Option<TLNextActionType>,
    pub common_time_node_data: Box<TLCommonTimeNodeData>,
    /// This element describes a list of conditions that shall be met in order to go backwards in an animation sequence.
    ///
    /// # Xml example
    ///
    /// Consider trying to emphasize text within a shape by changing the size of its font.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
    ///   </p:cTn>
    ///   <p:prevCondLst>
    ///     <p:cond evt="onPrev" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:prevCondLst>
    ///   <p:nextCondLst>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
    pub prev_condition_list: Option<TLTimeConditionList>,
    /// This element describes a list of conditions that shall be met to advance to the next animation sequence.
    ///
    /// # Xml example
    ///
    /// Consider a shape with a text emphasis changing the size of its font.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq"> … </p:cTn>
    ///   <p:prevCondLst> … </p:prevCondLst>
    ///   <p:nextCondLst>
    ///     <p:cond evt="onNext" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
    pub next_condition_list: Option<TLTimeConditionList>,
}

impl TLTimeNodeSequence {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut concurrent = None;
        let mut prev_action_type = None;
        let mut next_action_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "concurrent" => concurrent = Some(parse_xml_bool(value)?),
                "prevAc" => prev_action_type = Some(value.parse()?),
                "nextAc" => next_action_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_time_node_data = None;
        let mut prev_condition_list = None;
        let mut next_condition_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cTn" => common_time_node_data = Some(Box::new(TLCommonTimeNodeData::from_xml_element(child_node)?)),
                "prevCondLst" => prev_condition_list = Some(TLTimeConditionList::from_xml_element(child_node)?),
                "nextCondLst" => next_condition_list = Some(TLTimeConditionList::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_time_node_data =
            common_time_node_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cTn"))?;

        Ok(Self {
            concurrent,
            prev_action_type,
            next_action_type,
            common_time_node_data,
            prev_condition_list,
            next_condition_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateBehavior {
    /// This attribute specifies a relative offset value for the animation with respect to its
    /// position before the start of the animation.
    pub by: Option<String>,
    /// This attribute specifies the starting value of the animation.
    pub from: Option<String>,
    /// This attribute specifies the ending value for the animation as a percentage.
    pub to: Option<String>,
    /// This attribute specifies the interpolation mode for the animation.
    pub calc_mode: Option<TLAnimateBehaviorCalcMode>,
    /// This attribute specifies the type of property value.
    pub value_type: Option<TLAnimateBehaviorValueType>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element specifies a list of time animated value elements.
    ///
    /// ```xml
    /// <p:anim calcmode="lin" valueType="num">
    ///   <p:cBhvr additive="base"> … </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%">
    ///       <p:val>
    ///         <p:strVal val="1+#ppt_h/2"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100000">
    ///       <p:val>
    ///         <p:strVal val="#ppt_y"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    /// ```
    pub time_animate_value_list: Option<TLTimeAnimateValueList>,
}

impl TLAnimateBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut by = None;
        let mut from = None;
        let mut to = None;
        let mut calc_mode = None;
        let mut value_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "by" => by = Some(value.clone()),
                "from" => from = Some(value.clone()),
                "to" => to = Some(value.clone()),
                "calcmode" => calc_mode = Some(value.parse()?),
                "valueType" => value_type = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut time_animate_value_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "tavLst" => time_animate_value_list = Some(TLTimeAnimateValueList::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            by,
            from,
            to,
            calc_mode,
            value_type,
            common_behavior_data,
            time_animate_value_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateColorBehavior {
    /// This attribute specifies the color space in which to interpolate the animation. Values for
    /// example can be HSL & RGB.
    ///
    /// The values for from/to/by/etc. can still be specified in any supported color format
    /// without affecting the color space within which the animation happens.
    ///
    /// The RGB color space is best used for doing animations between two different colors since
    /// it doesn't require going through any other hues between the two colors specified. The
    /// HSL space is useful for animating through a rainbow of colors or for modifying just the
    /// saturation by 30% for example.
    pub color_space: Option<TLAnimateColorSpace>,
    /// This attribute specifies which direction to cycle the hue around the color wheel. Values
    /// are clockwise or counter clockwise. Default is clockwise.
    pub direction: Option<TLAnimateColorDirection>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the color animation.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animClr clrSpc="hsl">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:hsl h="0" s="0" l="0"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    pub by: Option<TLByAnimateColorTransform>,
    /// This element is used to specify the starting color of the target element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animClr clrSpc="rgb" dir="cw">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl> … </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>fillcolor</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:from>
    ///     <a:schemeClr val="accent3"/>
    ///   </p:from>
    ///   <p:to>
    ///     <a:schemeClr val="accent2"/>
    ///   </p:to>
    /// </p:animClr>
    /// ```
    pub from: Option<Color>,
    /// This element specifies the resulting color for the animation color change.
    ///
    /// # Xml example
    ///
    /// Consider emphasize a shape by changing its fill color from blue to red. The <to> element should be
    /// used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:animClr clrSpc="rgb">
    ///     <p:cBhvr> … </p:cBhvr>
    ///     <p:to>
    ///       <a:schemeClr val="accent2"/>
    ///     </p:to>
    ///   </p:animClr>
    /// </p:childTnLst>
    /// ```
    pub to: Option<Color>,
}

impl TLAnimateColorBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut color_space = None;
        let mut direction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "clrSpc" => color_space = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => {
                    by = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(TLByAnimateColorTransform::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| {
                                MissingChildNodeError::new(child_node.name.clone(), "TLByAnimateColorTransform")
                            })?,
                    )
                }
                "from" => {
                    from = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(Color::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?,
                    )
                }
                "to" => {
                    to = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(Color::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?,
                    )
                }
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            color_space,
            direction,
            common_behavior_data,
            by,
            from,
            to,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateEffectBehavior {
    /// This attribute specifies whether to transition the element in or out or treat it as a static
    /// filter. The values are "None", "In" and "Out", and the default value is "In".
    ///
    /// When a value of "In" is specified, the element is not visible at the start of the animation
    /// and is completely visible be the end of the duration. When "Out" is specified, the element
    /// is visible at the start and not visible at the end of the effect. This visibility is in addition to
    /// the effect of setting CSS visibility or display attributes.
    pub transition: Option<TLAnimateEffectTransition>,
    /// This attribute specifies the animation types and subtypes to be used for the effect.
    /// Multiple animations are allowed to be listed so that in the event that a superseding
    /// animation (leftmost) cannot be rendered, a fallback animation is available. That is, the
    /// rendering application parses the list from left to right until a supported animation is
    /// found.
    ///
    /// The syntax used for the filter attribute value is as follows: "type(subtype);type(subtype)".
    /// Subtype can be a string value such as "fromLeft" or a numerical value depending on the
    /// type specified.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal);blinds(vertical)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="5"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ///
    /// There are two animation filters shown in this example. The first is the blinds (horizontal),
    /// which the rendering application is to use as the primary animation effect. If, however,
    /// the rendering application does not support this animation, the blinds (vertical) animation
    /// is used. In this example there are only two animation filters listed, a primary and a
    /// fallback, but it is possible to list multiple fallback filters using the syntax defined above.
    pub filter: Option<String>,
    /// This attribute specifies a list of properties that coincide with the effect specified.
    /// Although there are many animation types allowed, this attribute allows the setting of
    /// specific property settings in order to describe an even wider variety of animation types.
    ///
    /// The syntax used for the prLst attribute value is as follows: “name:value;name:value”.
    /// When multiple animation types are listed in the filter attribute, the rendering application
    /// attempts to apply each property value even though some might not apply to it.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animEffect filter="image" prLst="opacity: 0.5">
    ///   <p:cBhvr rctx="IE">
    ///     <p:cTn id="7" dur="indefinite"/>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ///
    /// The animation filter specified is an image filter type that has a specific property called
    /// opacity set to a value of 0.5.
    pub property_list: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element defines the progression of an animation. The default for the way animation progress happens
    /// through an animEffect is a linear ramp from 0 to 1, starting at the effect’s begin time & ending at the effect’s
    /// end time. When you specify a value for the progress attribute, you are overriding this default behaviour. The
    /// value between 0 and 1 represents a percentage through the effect, where 0 is 0% and 1 is 100%.
    ///
    /// Each animEffect is in fact an object-based transition. These transitions can be specified as “In” (where the object
    /// is not visible at 0% and becomes completely visible at 100%) or “Out” (where the object is visible at 0% and
    /// becomes completely invisible at 100%). You would set the progress attribute if you want to use the animEffect
    /// as a “static” effect, where the transition properties do not actually change over time. As an alternative to using
    /// the progress attribute, you can use the tmFilter (time filter), which is a base attribute of any effect/timenode, to
    /// specify the way that progress through an effect should be performed dynamically.
    pub progress: Option<TLAnimVariant>,
}

impl TLAnimateEffectBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut transition = None;
        let mut filter = None;
        let mut property_list = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "transition" => transition = Some(value.parse()?),
                "filter" => filter = Some(value.clone()),
                "prLst" => property_list = Some(value.clone()),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut progress = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "progress" => {
                    progress = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(TLAnimVariant::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "CT_TLAnimVariant"))?,
                    )
                }
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            transition,
            filter,
            property_list,
            common_behavior_data,
            progress,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateMotionBehavior {
    /// Specifies what the origin of the motion path is relative to such as the layout of the slide,
    /// or the parent.
    pub origin: Option<TLAnimateMotionBehaviorOrigin>,
    /// Specifies the path primitive followed by coordinates for the animation motion. The
    /// allowed values that are understood within a path are as follows:
    ///
    /// M = move to, L = line to, C = curve to, Z=close loop, E=end
    /// UPPERCASE = absolute coords, lowercase = relative coords
    /// Thus total allowed set = {M,L,C,Z,E,m,l,c,z,e)
    ///
    /// # Example
    ///
    /// The following string is a sample path.
    /// path: “M 0 0 L 1 1 c 1 2 3 4 4 4 Z”
    pub path: Option<String>,
    /// This attribute specifies how the motion path moves when the target element is moved.
    pub path_edit_mode: Option<TLAnimateMotionPathEditMode>,
    /// The attribute describes the relative angle of the motion path.
    pub rotate_angle: Option<Angle>,
    /// This attribute describes the point type of the points in the path attribute. The allowed
    /// values that are understood for the ptsTypes attribute are as follows:
    ///
    /// A = Auto, F = Corner, T = Straight, S = Smooth
    /// UPPERCASE = Straight Line follows point, lowercase = curve follows point.
    /// Thus, the total allowed set = {A,F,T,S,a,f,t,s}
    pub points_types: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the animation.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:by x="150.000%" y="150.000%"/>
    /// </p:animScale>
    /// ```
    pub by: Option<TLPoint>,
    pub from: Option<TLPoint>,
    /// This element specifies the target location for an animation motion or animation scale effect
    ///
    /// # Xml example
    ///
    /// Consider an animation with a "light speed" entrance effect.
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="200" decel="10.5%" autoRev="1" fill="hold">
    ///       <p:stCondLst>
    ///         <p:cond delay="600"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
    pub to: Option<TLPoint>,
    /// This element describes the center of the rotation used to rotate a motion path by X angle.
    ///
    /// # Xml example
    ///
    /// For example, suppose we have a simple animation with a checkerbox text entrance.
    /// ```xml
    /// <p:animMotion origin="layout" path="M 0 0 L 0.25 0.33333 E" pathEditMode="relative" rAng="0" ptsTypes="">
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:rCtr x="56.7%" y="83.4%"/>
    /// </p:animMotion>
    /// ```
    pub rotation_center: Option<TLPoint>,
}

impl TLAnimateMotionBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut origin = None;
        let mut path = None;
        let mut path_edit_mode = None;
        let mut rotate_angle = None;
        let mut points_types = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "origin" => origin = Some(value.parse()?),
                "path" => path = Some(value.clone()),
                "pathEditMode" => path_edit_mode = Some(value.parse()?),
                "rAng" => rotate_angle = Some(value.parse()?),
                "ptsTypes" => points_types = Some(value.clone()),
                _ => (),
            }
        }

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;
        let mut rotation_center = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => by = Some(TLPoint::from_xml_element(child_node)?),
                "from" => from = Some(TLPoint::from_xml_element(child_node)?),
                "to" => to = Some(TLPoint::from_xml_element(child_node)?),
                "rCtr" => rotation_center = Some(TLPoint::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            origin,
            path,
            path_edit_mode,
            rotate_angle,
            points_types,
            common_behavior_data,
            by,
            from,
            to,
            rotation_center,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateRotationBehavior {
    /// This attribute describes the relative offset value for the animation.
    pub by: Option<Angle>,
    /// This attribute describes the starting value for the animation.
    pub from: Option<Angle>,
    /// This attribute describes the ending value for the animation.
    pub to: Option<Angle>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
}

impl TLAnimateRotationBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "by" => by = Some(value.parse()?),
                "from" => from = Some(value.parse()?),
                "to" => to = Some(value.parse()?),
                _ => (),
            }
        }

        let common_behavior_data = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cBhvr")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "cBhvr")))
            .and_then(TLCommonBehaviorData::from_xml_element)?
            .into();

        Ok(Self {
            by,
            from,
            to,
            common_behavior_data,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLAnimateScaleBehavior {
    /// This attribute specifies whether to zoom the contents of an object when doing a scaling
    /// animation.
    pub zoom_contents: Option<bool>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// This element describes the relative offset value for the animation.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:by x="150.000%" y="150.000%"/>
    /// </p:animScale>
    /// ```
    pub by: Option<TLPoint>,
    /// This element specifies an x/y co-ordinate to start the animation from.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn> … </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
    pub from: Option<TLPoint>,
    /// This element specifies the target location for an animation motion or animation scale effect
    ///
    /// # Xml example
    ///
    /// Consider an animation with a "light speed" entrance effect.
    /// ```xml
    /// <p:animScale>
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="200" decel="10.5%" autoRev="1" fill="hold">
    ///       <p:stCondLst>
    ///         <p:cond delay="600"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    ///   <p:from x="100%" y="100%"/>
    ///   <p:to x="80%" y="100%"/>
    /// </p:animScale>
    /// ```
    pub to: Option<TLPoint>,
}

impl TLAnimateScaleBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let zoom_contents = xml_node
            .attributes
            .get("zoomContents")
            .map(parse_xml_bool)
            .transpose()?;

        let mut common_behavior_data = None;
        let mut by = None;
        let mut from = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "by" => by = Some(TLPoint::from_xml_element(child_node)?),
                "from" => from = Some(TLPoint::from_xml_element(child_node)?),
                "to" => to = Some(TLPoint::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            zoom_contents,
            common_behavior_data,
            by,
            from,
            to,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLCommandBehavior {
    /// This attribute specifies the kind of command that is issued by the rendering application to
    /// the appropriate target application or object.
    ///
    /// There are three possible values, call, evt, and verb. A call command type is used to
    /// specify the class of commands that can then be issued.
    ///
    /// * Call commands: This command type is used to call methods on the object specified (play(), pause(), etc.)
    ///
    /// * Event commands: This command type is used to set an event for the object at this point in the timeline
    ///                   (onstopaudio, etc.)
    ///
    /// * Verb Commands: This command type is used to set verbs for the object to occur at this point in the timeline
    ///                  (0, 1, etc.)
    pub command_type: Option<TLCommandType>,
    /// This attribute defines the actual command to be issued. Depending on the command
    /// specified, the actual command can be made to invoke a wide range of actions on the
    /// linked or embedded object
    ///
    /// Reserved Values (when command_type == TLCommandType::Call):
    /// * play: play corresponding media
    /// * playFrom(s): play corresponding media starting from s, where s is the number of
    ///                seconds from the beginning of the clip
    /// * pause: pause corresponding media
    /// * resume: resume play of corresponding media
    /// * stop: stop play of corresponding media
    /// * togglePause: play corresponding media if media is already paused, pause
    ///                corresponding media if media is already playing. If the corresponding
    ///                media is not active, this command restarts the media and plays from
    ///                its beginning.
    ///
    /// Reserved Values (when command_type == TLCommandType::Event):
    /// * onstopaudio: stop play of all audio
    ///
    /// Reserved Values (when command_type == TLCommandType::Verb):
    /// * 0: Open the object for editing
    /// * 1: Open the object for viewing
    ///
    /// The value of the cmd attribute shall be the string representation of an integer that
    /// represents the embedded object verb number. This verb number determines the action
    /// that the rendering application should take corresponding to this object when this point in
    /// the animation is reached.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:cmd type="evt" cmd="onstopaudio">
    ///   <p:cBhvr>
    ///     <p:cTn display="0" masterRel="sameClick">
    ///       <p:stCondLst>
    ///         <p:cond evt="begin" delay="0">
    ///           <p:tn val="5"/>
    ///         </p:cond>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:sldTgt/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:cmd>
    /// ```
    ///
    /// In the above example, the event of onstopaudio stops all audio from playing once this
    /// particular animation is reached in the timeline.
    pub command: Option<String>,
    pub common_behavior_data: Box<TLCommonBehaviorData>,
}

impl TLCommandBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut command_type = None;
        let mut command = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => command_type = Some(value.parse()?),
                "cmd" => command = Some(value.clone()),
                _ => (),
            }
        }

        let common_behavior_data = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cBhvr")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "cBhvr")))
            .and_then(TLCommonBehaviorData::from_xml_element)?
            .into();

        Ok(Self {
            command_type,
            command,
            common_behavior_data,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLSetBehavior {
    pub common_behavior_data: Box<TLCommonBehaviorData>,
    /// The element specifies the certain attribute of a time node after an animation effect.
    ///
    /// # Xml example
    ///
    /// Consider an animation effect that leaves a string value visible afterwards. The <to> element should
    /// be used as follows:
    /// ```xml
    /// <p:childTnLst>
    ///   <p:set>
    ///     <p:cBhvr> … </p:cBhvr>
    ///     <p:to>
    ///       <p:strVal val="visible"/>
    ///     </p:to>
    ///   </p:set>
    ///   <p:anim calcmode="lin" valueType="num"> … </p:anim> …
    /// </p:childTnLst>
    /// ```
    pub to: Option<TLAnimVariant>,
}

impl TLSetBehavior {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut common_behavior_data = None;
        let mut to = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "cBhvr" => common_behavior_data = Some(Box::new(TLCommonBehaviorData::from_xml_element(child_node)?)),
                "to" => {
                    to = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .find_map(TLAnimVariant::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "CT_TLAnimVariant"))?,
                    )
                }
                _ => (),
            }
        }

        let common_behavior_data =
            common_behavior_data.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cBhvr"))?;

        Ok(Self {
            common_behavior_data,
            to,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLMediaNodeAudio {
    /// This attribute indicates whether the audio is a narration for the slide.
    ///
    /// Defaults to false
    pub is_narration: Option<bool>,
    pub common_media_node_data: Box<TLCommonMediaNodeData>,
}

impl TLMediaNodeAudio {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_narration = xml_node.attributes.get("isNarration").map(parse_xml_bool).transpose()?;

        let common_media_node_data = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cMediaNode")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "cMediaNode")))
            .and_then(TLCommonMediaNodeData::from_xml_element)?
            .into();

        Ok(Self {
            is_narration,
            common_media_node_data,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLMediaNodeVideo {
    /// This attribute specifies if the video is displayed in full-screen.
    ///
    /// Defaults to false
    pub fullscreen: Option<bool>,
    pub common_media_node_data: Box<TLCommonMediaNodeData>,
}

impl TLMediaNodeVideo {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fullscreen = xml_node.attributes.get("fullScrn").map(parse_xml_bool).transpose()?;

        let common_media_node_data = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cMediaNode")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "cMediaNode")))
            .and_then(TLCommonMediaNodeData::from_xml_element)?
            .into();

        Ok(Self {
            fullscreen,
            common_media_node_data,
        })
    }
}

/// This element defines a "keypoint" in animation interpolation.
///
/// # Xml example
///
/// Consider a shape with a "fly-in" animation. The <tav> element should be used as follows:
/// ```xml
/// <p:anim calcmode="lin" valueType="num">
///   <p:cBhvr additive="base"> … </p:cBhvr>
///   <p:tavLst>
///     <p:tav tm="0%">
///       <p:val>
///         <p:strVal val="1+#ppt_h/2"/>
///       </p:val>
///     </p:tav>
///     <p:tav tm="100%">
///       <p:val>
///         <p:strVal val="#ppt_y"/>
///       </p:val>
///     </p:tav>
///   </p:tavLst>
/// </p:anim>
/// ```
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TLTimeAnimateValue {
    /// This attribute specifies the time at which the attribute being animated takes on the value.
    ///
    /// Defaults to TLTimeAnimateValueTime::Indefinite
    pub time: Option<TLTimeAnimateValueTime>,
    /// This attribute allows for the specification of a formula to be used for describing a
    /// complex motion for an animated object. The formula manipulates the motion of the
    /// object by modifying a property of the object over a specified period of time. Each formula
    /// has zero or more inputs specified by the ($) symbol, zero or more variables specified by
    /// the (#) symbol pre-pended to the variable name and a target variable which is specified
    /// by the previously specified attrName element. The formula can contain one or more of
    /// any of the constants, operators or functions listed below. In addition to this, the formula
    /// can also contain floating point numbers and parentheses.
    ///
    /// Mathematical operations have the following order of precedence, listed from lowest to
    /// highest. Operators listed on the same line have equal precedence.
    ///
    /// * “+”, “-“
    /// * “*”, “/”, “%”
    /// * “^”
    /// * Unary minus, Unary plus (e.g. -2, meaning 3*-2 is the same as 3*(-2))
    /// * Variables, Constants (including numbers) and Functions (as listed previously)
    ///
    /// # Language Description
    ///
    /// Digit       = '0' | '1' | ‘2’ | ‘3’ | ‘4’ | ‘5’ | ‘6’ | ‘7’ | ‘8’ | '9' ;
    ///
    /// number      = digit , { digit } ;
    ///
    /// exponent    = [ '-' ] , ( 'e' | 'E' ) , number ;
    ///
    /// value       = number , [ '.' number ] , [ exponent ] ;
    ///
    /// variable    = '$' | 'ppt_x' | 'ppt_y' | 'ppt_w' | 'ppt_h' ;
    ///
    /// constant    = value | 'pi' | 'e' ;
    ///
    /// ident       = 'abs' | ‘acos’ | ‘asin’ | ‘atan’ | ‘ceil’
    ///               | ‘cos’ | ‘cosh’ | ‘deg’ | ‘exp’ | ‘floor’ | ‘ln’
    ///               | ‘max’ | ‘min’ | ‘rad’ | ‘rand’ | ‘sin’ | ‘sinh’
    ///               | ‘sqrt’ | ‘tan’ | 'tanh' ;
    ///
    /// function    = ident , '(' , formula [ ',' , formula ] , ')' ;
    ///
    /// formula     = term , { [ '+' | '-' ] , term } ;
    ///
    /// term        = power , { [ '*' | '/' | '%' ] , power } ;
    ///
    /// power       = unary [ '^' , unary ] ;
    ///
    /// unary       = [ '+' | '-' ] , factor ;
    ///
    /// factor      = variable | constant | function | parens ;
    ///
    /// parens      = '(' , formula , ')' ;
    ///
    /// ## Note
    ///
    /// Formulas can only support a calcMode (Calculation Mode) of linear or discrete. If
    /// another calcMode is specified or no calcMode is specified then a calcMode of linear is
    /// assumed.
    ///
    /// Any additional characters in the formula string that are not contained within the
    /// set described are considered invalid.
    ///
    /// # Variables
    ///
    /// |Name       |Description                                        |
    /// |-----------|---------------------------------------------------|
    /// |$          |Formula input                                      |
    /// |ppt_x      |Pre-animation x position of the object on the slide|
    /// |ppt_y      |Pre-animation y position of the object on the slide|
    /// |ppt_w      |Pre-animation width of the object                  |
    /// |ppt_h      |Pre-animation height of the object                 |
    ///
    /// # Constants
    ///
    /// |Name       |Description                                        |
    /// |-----------|---------------------------------------------------|
    /// |pi         |Mathematical constant pi                           |
    /// |e          |Mathematical constant e                            |
    ///
    /// # Operators
    ///
    /// |Name       |Description        |Usage                                  |
    /// |-----------|-------------------|---------------------------------------|
    /// |+          |Addition           |“x+y”, adds x to the value y           |
    /// |-          |Subtraction        |“x-y”, subtracts y from the value x    |
    /// |*          |Multiplication     |“x*y”, multiplies x by the value y     |
    /// |/          |Division           |“x/y”, divides x by the value y        |
    /// |%          |Modulus            |“x%y”, the remainder of x/y            |
    /// |^          |Power              |“x^y”, x raised to the power y         |
    ///
    /// # Functions
    ///
    /// |Name       |Description                |Usage                                                              |
    /// |-----------|---------------------------|-------------------------------------------------------------------|
    /// |abs        |Absolute value             |“abs(x)”, absolute value of x                                      |
    /// |acos       |Arc Cosine                 |“acos(x)”, arc cosine of the value x                               |
    /// |asin       |Arc Sine                   |“asin(x)”, arc sine of the value x                                 |
    /// |atan       |Arc Tangent                |“atan(x)”, arc tangent of the value x                              |
    /// |ceil       |Ceil value                 |“ceil(x)”, value of x rounded up                                   |
    /// |cos        |Cosine                     |“cos(x)”, cosine of the value of x                                 |
    /// |cosh       |Hyperbolic Cosine          |“cosh(x)", hyperbolic cosine of the value x                        |
    /// |deg        |Radiant to Degree convert  |“deg(x)”, the degree value of radiant value x                      |
    /// |exp        |Exponent                   |“exp(x)”, value of constant e raised to the power of x             |
    /// |floor      |Floor value                |“floor(x)”, value of x rounded down                                |
    /// |ln         |Natural logarithm          |“ln(x)”, natural logarithm of x                                    |
    /// |max        |Maximum of two values      |“max(x,y)”, returns x if (x > y) or returns y if (y > x)           |
    /// |min        |Minimum of two values      |“min(x,y)", returns x if (x < y) or returns y if (y < x)           |
    /// |rad        |Degree to Radiant convert  |“rad(x)”, the radiant value of degree value x                      |
    /// |rand       |Random value               |“rand(x)”, returns a random floating point value between 0 and x   |
    /// |sin        |Sine                       |“sin(x)”, sine of the value x                                      |
    /// |sinh       |Hyperbolic Sine            |"sinh(x)”, hyperbolic sine of the value x                          |
    /// |sqrt       |Square root                |“sqrt(x)”, square root of the value x                              |
    /// |tan        |Tangent                    |“tan(x)”, tangent of the value x                                   |
    /// |tanh       |Hyperbolic Tangent         |“tanh(x)", hyperbolic tangent of the value x                       |
    ///
    /// # Xml example
    ///
    /// <p:animcalcmode="lin" valueType="num">
    ///   <p:cBhvr>
    ///     <p:cTn id="9" dur="664" tmFilter="0.0,0.0; 0.25,0.07;0.50,0.2; 0.75,0.467; 1.0,1.0">
    ///       <p:stCondLst>
    ///         <p:cond delay="0"/>
    ///       </p:stCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgtspid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%" fmla="#ppt_y-sin(pi*$)/3">
    ///       <p:val>
    ///         <p:fltValval="0.5"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100%">
    ///       <p:val>
    ///         <p:fltValval="1"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    ///
    /// The animation example above modifies the ppt_y variable of the object by subtracting
    /// sin(pi*$)/3 from the non-animated value of ppt_y. The start value is 0.5 and the end
    /// value is 1 specified in each of the val elements. The total time for this animation is
    /// specified within the dur attribute and the filtered time graph is specified by the tmFilter
    /// attribute. The end result is that the object moves from a point above its non-animated
    /// position back to its non-animated position. With the specification of the tmFilter it has a
    /// modified time graph such that it also appears to accelerate as it reaches its final position.
    ///
    /// ## Note
    ///
    /// For this example, the non-animated value of ppt_y is the value of this variable if
    /// the object were to be statically rendered on the slide without animation properties.
    pub formula: Option<String>,
    /// The element specifies a value for a time animate.
    ///
    /// # Xml example
    ///
    /// Consider a shape with a fade in animation effect. The <val> element should be used as follows:
    /// ```xml
    /// <p:anim calcmode="lin" valueType="num">
    ///   <p:cBhvr additive="base"> … </p:cBhvr>
    ///   <p:tavLst>
    ///     <p:tav tm="0%">
    ///       <p:val>
    ///         <p:strVal val="0-#ppt_w/2"/>
    ///       </p:val>
    ///     </p:tav>
    ///     <p:tav tm="100%">
    ///       <p:val>
    ///         <p:strVal val="#ppt_x"/>
    ///       </p:val>
    ///     </p:tav>
    ///   </p:tavLst>
    /// </p:anim>
    /// ```
    pub value: Option<TLAnimVariant>,
}

impl TLTimeAnimateValue {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "tm" => instance.time = Some(value.parse()?),
                    "fmla" => instance.formula = Some(value.clone()),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|instance| {
                let value = xml_node
                    .child_nodes
                    .iter()
                    .find(|child_node| child_node.local_name() == "val")
                    .map(|child_node| {
                        child_node
                            .child_nodes
                            .iter()
                            .find(|val_node| TLAnimVariant::is_choice_member(val_node.local_name()))
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "TLAnimVariant").into())
                            .and_then(TLAnimVariant::from_xml_element)
                    })
                    .transpose()?;

                Ok(Self { value, ..instance })
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TLTimeAnimateValueList(pub Vec<TLTimeAnimateValue>);

impl TLTimeAnimateValueList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let vec = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "tav")
            .map(TLTimeAnimateValue::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(vec))
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLTimeAnimateValueTime {
    Percentage(PositiveFixedPercentage),
    Indefinite,
}

impl FromStr for TLTimeAnimateValueTime {
    type Err = ParseEnumError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "indefinite" => Ok(TLTimeAnimateValueTime::Indefinite),
            _ => Ok(TLTimeAnimateValueTime::Percentage(
                s.parse().map_err(|_| Self::Err::new("TLTimeAnimateValueTime"))?,
            )),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLAnimVariant {
    /// This element specifies a boolean value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Bool(bool),
    /// This element specifies an integer value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Int(i32),
    /// This element specifies a floating point value to be used for evaluation by a parent element. The exact meaning
    /// of the value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    Float(f32),
    /// This element specifies a string value to be used for evaluation by a parent element. The exact meaning of the
    /// value contained within this element is not defined here but is dependent on the usage of this element in
    /// conjunction with one of the listed parent elements.
    String(String),
    /// This element describes the color variant. This is used to specify a color that is to be used for animating the color
    /// property of an object.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:set>
    ///   <p:cBhvr override="childStyle">
    ///     …
    ///   </p:cBhvr>
    ///   <p:to>
    ///     <p:clrVal>
    ///       <a:schemeClr val="accent2"/>
    ///     </p:clrVal>
    ///   </p:to>
    /// </p:set>
    /// ```
    Color(Color),
}

impl XsdType for TLAnimVariant {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "boolVal" => Ok(TLAnimVariant::Bool(parse_xml_bool(xml_node.get_val_attribute()?)?)),
            "intVal" => Ok(TLAnimVariant::Int(xml_node.get_val_attribute()?.parse()?)),
            "fltVal" => Ok(TLAnimVariant::Float(xml_node.get_val_attribute()?.parse()?)),
            "strVal" => Ok(TLAnimVariant::String(xml_node.get_val_attribute()?.clone())),
            "clrVal" => {
                let color = xml_node
                    .child_nodes
                    .iter()
                    .find_map(Color::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;
                Ok(TLAnimVariant::Color(color))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TLAnimVariant").into()),
        }
    }
}

impl XsdChoice for TLAnimVariant {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "boolVal" | "intVal" | "fltVal" | "strVal" | "clrVal" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLTimeConditionTriggerGroup {
    TargetElement(TLTimeTargetElement),
    /// This element describes the time node trigger choice.
    ///
    /// # Xml example
    ///
    /// Consider a time node with an event condition. The <tn> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5">
    ///     <p:stCondLst>
    ///       <p:cond delay="0"/>
    ///     </p:stCondLst>
    ///     <p:endCondLst>
    ///       <p:cond evt="begin" delay="0">
    ///         <p:tn val="5"/>
    ///       </p:cond>
    ///     </p:endCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    TimeNode(TLTimeNodeId),
    /// This element specifies the child time node that triggers a time condition. References a child time node or all
    /// child nodes. Order is based on the child's end time.
    ///
    /// # Xml example
    ///
    /// Consider an animation which ends the synchronization of all parallel time nodes when all the child
    /// nodes have ended their animation. The <rtn> element should be used as follows:
    /// ```xml
    /// <p:cTn>
    ///   <p:stCondLst> … </p:stCondLst>
    ///   <p:endSync evt="end" delay="0">
    ///     <p:rtn val="all"/>
    ///   </p:endSync>
    ///   <p:childTnLst> … </p:childTnLst>
    /// </p:cTn>
    /// ```
    RuntimeNode(TLTriggerRuntimeNode),
}

impl XsdType for TLTimeConditionTriggerGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tgtEl" => {
                let target_element = xml_node
                    .child_nodes
                    .iter()
                    .find_map(TLTimeTargetElement::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "sldTgt|sndTgt|spTgt|inkTgt"))?;

                Ok(TLTimeConditionTriggerGroup::TargetElement(target_element))
            }
            "tn" => {
                let val = xml_node.get_val_attribute()?.parse()?;
                Ok(TLTimeConditionTriggerGroup::TimeNode(val))
            }
            "rtn" => {
                let val = xml_node.get_val_attribute()?.parse()?;
                Ok(TLTimeConditionTriggerGroup::RuntimeNode(val))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_TLTimeConditionTriggerGroup",
            ))),
        }
    }
}

impl XsdChoice for TLTimeConditionTriggerGroup {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "tgtEl" | "tn" | "rtn" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLTimeTargetElement {
    /// This element specifies the slide as the target element.
    ///
    /// # Xml example
    ///
    /// For example, suppose we have a simple animation with a blind entrance.
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn id="2" dur="indefinite" nodeType="mainSeq"> … </p:cTn>
    ///   <p:prevCondLst> … </p:prevCondLst>
    ///   <p:nextCondLst>
    ///     <p:cond evt="onNext" delay="0">
    ///       <p:tgtEl>
    ///         <p:sldTgt/>
    ///       </p:tgtEl>
    ///     </p:cond>
    ///   </p:nextCondLst>
    /// </p:seq>
    /// ```
    SlideTarget,
    /// This element describes the sound information for a target object.
    ///
    /// # Xml example
    ///
    /// Consider a shape with a sound effect animation. The <sndTgt> element should be used as follows:
    /// ```xml
    /// <p:subTnLst>
    ///   <p:audio>
    ///     <p:cMediaNode>
    ///       <p:cTn display="0" masterRel="sameClick"> … </p:cTn>
    ///       <p:tgtEl>
    ///         <p:sndTgt r:embed="rId2" r:link="rId3"/>
    ///       </p:tgtEl>
    ///     </p:cMediaNode>
    ///   </p:audio>
    /// </p:subTnLst>
    /// ```
    SoundTarget(EmbeddedWAVAudioFile),
    /// The element specifies the shape in which to apply a certain animation to.Err
    ///
    /// # Xml example
    ///
    /// Consider a shape whose id is 3 in which we want to apply a fade animation effect. The <spTgt> should
    /// be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="fade">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="2000"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ShapeTarget(TLShapeTargetElement),
    /// This element specifies an animation target element that is represented by a sub-shape in a legacy graphical
    /// object.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:inkTgt spid="_x0000_s2057"/>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    InkTarget(TLSubShapeId),
}

impl XsdType for TLTimeTargetElement {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "sldTgt" => Ok(TLTimeTargetElement::SlideTarget),
            "sndTgt" => Ok(TLTimeTargetElement::SoundTarget(
                EmbeddedWAVAudioFile::from_xml_element(xml_node)?,
            )),
            "spTgt" => Ok(TLTimeTargetElement::ShapeTarget(
                TLShapeTargetElement::from_xml_element(xml_node)?,
            )),
            "inkTgt" => {
                let spid = xml_node
                    .attributes
                    .get("spid")
                    .ok_or_else(|| Box::<dyn Error>::from(MissingAttributeError::new(xml_node.name.clone(), "spid")))
                    .and_then(|value| value.parse().map_err(Into::into))?;

                Ok(TLTimeTargetElement::InkTarget(spid))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_TLTimeTargetElement",
            ))),
        }
    }
}

impl XsdChoice for TLTimeTargetElement {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "sldTgt" | "sndTgt" | "spTgt" | "inkTgt" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLShapeTargetElement {
    /// This attribute specifies the shape identifier.
    pub shape_id: DrawingElementId,
    pub target: Option<TLShapeTargetElementGroup>,
}

impl TLShapeTargetElement {
    pub fn new(shape_id: DrawingElementId, target: Option<TLShapeTargetElementGroup>) -> Self {
        Self { shape_id, target }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let shape_id = xml_node
            .attributes
            .get("spid")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?
            .parse()?;

        let target = xml_node
            .child_nodes
            .iter()
            .find_map(TLShapeTargetElementGroup::try_from_xml_element)
            .transpose()?;

        Ok(Self { shape_id, target })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLShapeTargetElementGroup {
    /// This element is used to specify animating the background of an object.
    ///
    /// # Xml example
    ///
    /// Consider adding animation to the background of Shape Id 3. The <bg> tag can be used as follows:
    ///
    /// ```xml
    /// <p:tgtEl>
    ///   <p:spTgt spid="3">
    ///     <p:bg/>
    ///   </p:spTgt>
    /// </p:tgtEl>
    /// ```
    Background,
    /// This element specifies the subshape of a legacy graphical object to animate.
    ///
    /// # Xml example
    ///
    /// Consider adding animation to a legacy diagram. The <subSp> element should be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="7" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="2053">
    ///         <p:subSp spid="_x0000_s70664"/>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    SubShape(TLSubShapeId),
    /// This element specifies the subelement of an embedded chart to animate.
    ///
    /// # Xml example
    ///
    /// Consider an embedded Chart with a entrance animation effect applied to each of the graph's
    /// categories. The <oldChartEl> element should be used as follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="blinds(horizontal)">
    ///   <p:cBhvr>
    ///     <p:cTn id="12" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="19460">
    ///         <p:oleChartEl type="category" lvl="1"/>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    OleChartElement(TLOleChartTargetElement),
    /// This element specifies a text element to animate.
    ///
    /// # Xml example
    ///
    /// Consider a shape containing text to be animated. The <txEl> should be used as follows:
    /// ```xml
    /// <p:tgtEl>
    ///   <p:spTgt spid="5">
    ///     <p:txEl>
    ///       <p:pRg st="1" end="1"/>
    ///     </p:txEl>
    ///   </p:spTgt>
    /// </p:tgtEl>
    /// ```
    TextElement(Option<TLTextTargetElement>),
    /// This element specifies a graphical element which to animate
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:set>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="1" fill="hold"> … </p:cTn>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4">
    ///         <p:graphicEl>
    ///           <a:dgm id="{87C2C707-C3F4-4E81-A967-A8B8AE13E575}"/>
    ///         </p:graphicEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst> … </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:to> … </p:to>
    /// </p:set>
    /// ```
    GraphicElement(AnimationElementChoice),
}

impl XsdType for TLShapeTargetElementGroup {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bg" => Ok(TLShapeTargetElementGroup::Background),
            "subSp" => {
                let spid = xml_node
                    .attributes
                    .get("spid")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "spid"))?
                    .parse()?;

                Ok(TLShapeTargetElementGroup::SubShape(spid))
            }
            "oleChartEl" => Ok(TLShapeTargetElementGroup::OleChartElement(
                TLOleChartTargetElement::from_xml_element(xml_node)?,
            )),
            "txEl" => {
                let text_element = xml_node
                    .child_nodes
                    .iter()
                    .find_map(TLTextTargetElement::try_from_xml_element)
                    .transpose()?;

                Ok(TLShapeTargetElementGroup::TextElement(text_element))
            }
            "graphicEl" => {
                let animation_element = xml_node
                    .child_nodes
                    .iter()
                    .find_map(AnimationElementChoice::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_AnimationElementChoice"))?;

                Ok(TLShapeTargetElementGroup::GraphicElement(animation_element))
            }
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLShapeTargetElementGroup",
            ))),
        }
    }
}

impl XsdChoice for TLShapeTargetElementGroup {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "bg" | "subSp" | "oleChartEl" | "txEl" | "graphicEl" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLOleChartTargetElement {
    /// This attribute specifies how to chart should be built during its animation.
    pub element_type: TLChartSubelementType,
    /// This attribute describes the element levels to animate.
    ///
    /// Defaults to 0
    pub level: Option<u32>,
}

impl TLOleChartTargetElement {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut element_type = None;
        let mut level = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => element_type = Some(value.parse()?),
                "lvl" => level = Some(value.parse()?),
                _ => (),
            }
        }

        let element_type = element_type.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self { element_type, level })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLTextTargetElement {
    /// This element specifies animation on a character range defined by a start and end character position.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animMotion>
    ///   <p:cBhvr>
    ///     <p:cTn id="6" dur="2000" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:charRg st="0" end="9"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>ppt_x</p:attrName>
    ///       <p:attrName>ppt_y</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    /// </p:animMotion>
    /// ```
    CharRange(IndexRange),
    /// This element specifies a text range to animate based on starting and ending paragraph number.
    ///
    /// # Xml example
    ///
    /// Consider an animation entrance of the first 3 text paragraphs. The <pRg> element should be used as
    /// follows:
    /// ```xml
    /// <p:animEffect transition="in" filter="checkerboard(across)">
    ///   <p:cBhvr>
    ///     <p:cTn id="12" dur="500"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="3">
    ///         <p:txEl>
    ///           <p:pRg st="0" end="2"/>
    ///         </p:txEl>
    ///       </p:spTgt>
    ///     </p:tgtEl>
    ///   </p:cBhvr>
    /// </p:animEffect>
    /// ```
    ParagraphRange(IndexRange),
}

impl XsdType for TLTextTargetElement {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "charRg" => Ok(TLTextTargetElement::CharRange(IndexRange::from_xml_element(xml_node)?)),
            "pRg" => Ok(TLTextTargetElement::ParagraphRange(IndexRange::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLTextTargetElement",
            ))),
        }
    }
}

impl XsdChoice for TLTextTargetElement {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "charRg" | "pRg" => true,
            _ => false,
        }
    }
}

/// This element specifies conditions on time nodes in a timeline. It is used within a list of start condition or list of
/// end condition elements.
///
/// # Xml example
///
/// ```xml
/// <p:cTn>
///   <p:stCondLst>
///     <p:cond delay="2000"/>
///   </p:stCondLst>
///   <p:childTnLst>
///     <p:set> … </p:set>
///     <p:animEffect transition="in" filter="blinds(horizontal)">
///       <p:cBhvr>
///         <p:cTn id="7" dur="1000"/>
///         <p:tgtEl>
///           <p:spTgt spid="4"/>
///         </p:tgtEl>
///       </p:cBhvr>
///     </p:animEffect>
///   </p:childTnLst>
/// </p:cTn>
/// ```
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TLTimeCondition {
    /// This attribute describes the event that triggers an animation.
    pub trigger_event: Option<TLTriggerEvent>,
    /// This attribute describes the delay after an animation is triggered.
    pub delay: Option<TLTime>,
    pub trigger: Option<TLTimeConditionTriggerGroup>,
}

impl TLTimeCondition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "evt" => instance.trigger_event = Some(value.parse()?),
                    "delay" => instance.delay = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|instance| {
                let trigger = xml_node
                    .child_nodes
                    .iter()
                    .find_map(TLTimeConditionTriggerGroup::try_from_xml_element)
                    .transpose()?;

                Ok(Self { trigger, ..instance })
            })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TLTimeConditionList(pub Vec<TLTimeCondition>);

impl TLTimeConditionList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let list = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "cond")
            .map(TLTimeCondition::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        if list.is_empty() {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "cond",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        } else {
            Ok(Self(list))
        }
    }
}

/// This element describes the properties that are common for time nodes.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TLCommonTimeNodeData {
    /// This attribute specifies the identifier for the timenode.
    pub id: Option<TLTimeNodeId>,
    /// This attribute describes the preset identifier for the time node.
    pub preset_id: Option<i32>,
    /// This attribute descries the class of effect in which it belongs.
    pub preset_class: Option<TLTimeNodePresetClassType>,
    /// This attribute is a bitflag that specifies a direction or some other attribute of the effect.
    /// For example it can be set to specify a “From Bottom” for the Fly In effect, or “Bold” for
    /// the Change Font Style effect.
    pub preset_subtype: Option<i32>,
    /// This attribute describes the duration of the time node, expressed as unit time.
    pub duration: Option<TLTime>,
    /// This attribute describes the number of times the element should repeat, in units of
    /// thousandths.
    ///
    /// Defaults to 100_0
    pub repeat_count: Option<TLTime>,
    /// This attribute describes the amount of time over which the element should repeat. If
    /// absent, the attribute is taken to be the same as the specified duration.
    pub repeat_duration: Option<TLTime>,
    /// This attribute specifies the percentage by which to speed up (or slow down) the timing. If
    /// negative, the timing is reversed.
    ///
    /// Defaults to 100_000
    ///
    /// # Example
    ///
    /// If speed is 200% (200_000) and the specified duration is 10 seconds, the actual duration is 5 seconds.
    pub speed: Option<Percentage>,
    /// This attribute describes the percentage of specified duration over which the element's
    /// time takes to accelerate from 0 up to the "run rate."
    ///
    /// Defaults to 0
    pub acceleration: Option<PositiveFixedPercentage>,
    /// This attribute describes the percentage of specified duration over which the element's
    /// time takes to decelerate from the "run rate" down to 0.
    ///
    /// Defaults to 0
    pub deceleration: Option<PositiveFixedPercentage>,
    /// This attribute describes whether to automatically play the animation in reverse after
    /// playing it in the forward direction.
    ///
    /// Defaults to false
    pub auto_reverse: Option<bool>,
    /// This attribute specifies if a node is to restart when it completes its action
    pub restart_type: Option<TLTimeNodeRestartType>,
    /// This attribute describes the fill type for the time node.
    pub fill_type: Option<TLTimeNodeFillType>,
    /// This attribute specifies how the time node synchronizes to its group.
    pub sync_behavior: Option<TLTimeNodeSyncType>,
    /// This attribute specifies the time filter for the time node.
    pub time_filter: Option<String>,
    /// This attribute describes the event filter for this time node.
    pub event_filter: Option<String>,
    /// This attribute describes whether the state of the time node is visible or hidden.
    pub display: Option<bool>,
    /// This attribute specifies how the time node plays back relative to its master time node.
    pub master_relationship: Option<TLTimeNodeMasterRelation>,
    /// This attribute describes the build level of the animation.
    pub build_level: Option<i32>,
    /// This attribute describes the Group ID of the time node.
    pub group_id: Option<u32>,
    /// This attribute specifies whether there is an after effect applied to the time node.
    pub after_effect: Option<bool>,
    /// This attribute specifies the type of time node.
    pub node_type: Option<TLTimeNodeType>,
    /// This attribute describes whether this node is a placeholder.
    pub node_placeholder: Option<bool>,
    /// This element contains a list conditions that shall be met for a time node to be activated.
    ///
    /// # Xml example
    ///
    /// example, suppose we have a shape with an entrance appearance after 5 seconds. The
    /// <stCondLst> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" nodeType="clickEffect">
    ///     <p:stCondLst>
    ///       <p:cond delay="5000"/>
    ///     </p:stCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    pub start_condition_list: Option<TLTimeConditionList>,
    /// This element describes a list of the end conditions that shall be met in order to stop the time node.
    ///
    /// # Xml example
    ///
    /// Consider a shape a shape with an audio attached to the animation. The <endCondList> element
    /// should be used as follows to specifies when the sound is done:
    /// ```xml
    /// <p:audio>
    ///   <p:cMediaNode>
    ///     <p:cTn display="0" masterRel="sameClick">
    ///       <p:stCondLst> … </p:stCondLst>
    ///       <p:endCondLst>
    ///         <p:cond evt="onStopAudio" delay="0">
    ///           <p:tgtEl>
    ///             <p:sldTgt/>
    ///           </p:tgtEl>
    ///         </p:cond>
    ///       </p:endCondLst>
    ///     </p:cTn>
    ///     <p:tgtEl> … </p:tgtEl>
    ///   </p:cMediaNode>
    /// </p:audio>
    /// ```
    pub end_condition_list: Option<TLTimeConditionList>,
    /// This element is used to synchronizes the stopping of parallel elements in the timing tree. It is used on interactive
    /// timeline sequences to specify that the interactive sequence’s duration ends when all of the child timenodes
    /// have ended. It is also used to make interactive sequences restart-able (so that the entire interactive sequence
    /// can be repeated if the trigger object is clicked on repeatedly).
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:seq concurrent="1" nextAc="seek">
    ///   <p:cTn>
    ///     <p:stCondLst/>
    ///     <p:endSync evt="end" delay="0">
    ///       <p:rtn val="all"/>
    ///     </p:endSync>
    ///     <p:childTnLst/>
    ///   </p:cTn>
    ///   <p:nextCondLst/>
    /// </p:seq>
    /// ```
    pub end_sync: Option<TLTimeCondition>,
    /// This element specifies how the animation should be successively applied to sub elements of the target element
    /// for a repeated effect. It can be applied to contained timing and animation structures over the letters, words, or
    /// shapes within a target element.
    ///
    /// # Xml example
    ///
    /// Consider a text animation where the words appear letter by letter. The <iterate> element should be
    /// used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="1" >
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmPct val="10000"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    pub iterate: Option<TLIterateData>,
    /// This element describes the list of time nodes that have a fixed location in the timing tree based on their parent
    /// time node. The children's start time is defined relative to their parent time node’s start.
    pub child_time_node_list: Option<TLTimeNodeList>,
    /// This element describes time nodes that have a start time which is not based on the containing timenode. It is
    /// instead based on their master relationship (masterRel). At runtime, they are inserted dynamically into the
    /// timing tree as child timenodes for playback, based on the logic defined by the master relationship. These
    /// elements are used for animations such as "dim after" and "play sound effects"
    ///
    /// # Xml example
    ///
    /// Consider an animation with a "Fly In" effect on paragraphs so that each paragraph flies in on a
    /// separate click. Then the "Dim After" effect for paragraph 1 doe not happen until paragraph 2 flies in. The
    /// <subTnLst> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" grpId="0" nodeType="clickEffect">
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:childTnLst> … </p:childTnLst>
    ///     <p:subTnLst>
    ///       <p:set>
    ///         <p:cBhvr override="childStyle">
    ///           <p:cTn fill="hold" masterRel="nextClick" afterEffect="1"/>
    ///           <p:tgtEl> … </p:tgtEl>
    ///           <p:attrNameLst> … </p:attrNameLst>
    ///         </p:cBhvr>
    ///         <p:to> … </p:to>
    ///         </p:set>
    ///     </p:subTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    pub sub_time_node_list: Option<TLTimeNodeList>,
}

impl TLCommonTimeNodeData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "id" => instance.id = Some(value.parse()?),
                    "presetID" => instance.preset_id = Some(value.parse()?),
                    "presetClass" => instance.preset_class = Some(value.parse()?),
                    "presetSubtype" => instance.preset_subtype = Some(value.parse()?),
                    "dur" => instance.duration = Some(value.parse()?),
                    "repeatCount" => instance.repeat_count = Some(value.parse()?),
                    "repeatDur" => instance.repeat_duration = Some(value.parse()?),
                    "spd" => instance.speed = Some(value.parse()?),
                    "accel" => instance.acceleration = Some(value.parse()?),
                    "decel" => instance.deceleration = Some(value.parse()?),
                    "autoRev" => instance.auto_reverse = Some(parse_xml_bool(value)?),
                    "restart" => instance.restart_type = Some(value.parse()?),
                    "fill" => instance.fill_type = Some(value.parse()?),
                    "syncBehavior" => instance.sync_behavior = Some(value.parse()?),
                    "tmFilter" => instance.time_filter = Some(value.clone()),
                    "evtFilter" => instance.event_filter = Some(value.clone()),
                    "display" => instance.display = Some(parse_xml_bool(value)?),
                    "masterRel" => instance.master_relationship = Some(value.parse()?),
                    "bldLvl" => instance.build_level = Some(value.parse()?),
                    "grpId" => instance.group_id = Some(value.parse()?),
                    "afterEffect" => instance.after_effect = Some(parse_xml_bool(value)?),
                    "nodeType" => instance.node_type = Some(value.parse()?),
                    "nodePh" => instance.node_placeholder = Some(parse_xml_bool(value)?),
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
                            "stCondLst" => {
                                instance.start_condition_list = Some(TLTimeConditionList::from_xml_element(child_node)?)
                            }
                            "endCondLst" => {
                                instance.end_condition_list = Some(TLTimeConditionList::from_xml_element(child_node)?)
                            }
                            "endSync" => instance.end_sync = Some(TLTimeCondition::from_xml_element(child_node)?),
                            "iterate" => instance.iterate = Some(TLIterateData::from_xml_element(child_node)?),
                            "childTnLst" => {
                                instance.child_time_node_list = Some(TLTimeNodeList::from_xml_element(child_node)?)
                            }
                            "subTnLst" => {
                                instance.sub_time_node_list = Some(TLTimeNodeList::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLIterateDataChoice {
    /// This element describes the duration of the iteration interval in absolute time.
    ///
    /// # Xml example
    ///
    /// Consider a text animation where the words appear letter by letter every 10 seconds. The <tmAbs>
    /// element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5" >
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmAbs val="10000"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    Absolute(TLTime),
    /// This element describes the duration of the iteration interval in a percentage of time.
    ///
    /// # Xml example
    ///
    /// Consider a text animation where the words appear letter by letter every 10th of the animation
    /// duration. The <tmPct> element should be used as follows:
    /// ```xml
    /// <p:par>
    ///   <p:cTn id="5">
    ///     <p:stCondLst> … </p:stCondLst>
    ///     <p:iterate type="lt">
    ///       <p:tmPct val="10%"/>
    ///     </p:iterate>
    ///     <p:childTnLst> … </p:childTnLst>
    ///   </p:cTn>
    /// </p:par>
    /// ```
    Percent(PositivePercentage),
}

impl XsdType for TLIterateDataChoice {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tmAbs" => Ok(TLIterateDataChoice::Absolute(xml_node.get_val_attribute()?.parse()?)),
            "tmPct" => Ok(TLIterateDataChoice::Percent(xml_node.get_val_attribute()?.parse()?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLIterateDataChoice",
            ))),
        }
    }
}

impl XsdChoice for TLIterateDataChoice {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "tmAbs" | "tmPct" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLIterateData {
    /// This attribute specifies the iteration behavior and applies it to each letter, word or shape
    /// within a container element.
    ///
    /// Values are by word, by letter, or by element. If there is no text or block elements such as
    /// shapes within the container or a single word, letter, or shape (depending on iterate type)
    /// then no iteration happens and the behavior is applied to the element itself instead.
    ///
    /// Defaults to IterateType::Element
    pub iterate_type: Option<IterateType>,
    /// This attribute specifies whether to go backwards in the timeline to the previous node.
    ///
    /// Defaults to false
    pub backwards: Option<bool>,
    pub interval: TLIterateDataChoice,
}

impl TLIterateData {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut iterate_type = None;
        let mut backwards = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => iterate_type = Some(value.parse()?),
                "backwards" => backwards = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let interval = xml_node
            .child_nodes
            .first()
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "TLIterateDataChoice").into())
            .and_then(TLIterateDataChoice::from_xml_element)?;

        Ok(Self {
            iterate_type,
            backwards,
            interval,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TLByAnimateColorTransform {
    /// The element specifies an incremental RGB value to add to the color property
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animClr clrSpc="rgb">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:rgb r="10" g="20" b="30"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    Rgb(TLByRgbColorTransform),
    /// This element specifies an incremental HSL (Hue, Saturation, Lightness) value to add to a color animation.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:animClr clrSpc="hsl">
    ///   <p:cBhvr>
    ///     <p:cTn id="8" dur="500" fill="hold"/>
    ///     <p:tgtEl>
    ///       <p:spTgt spid="4"/>
    ///     </p:tgtEl>
    ///     <p:attrNameLst>
    ///       <p:attrName>stroke.color</p:attrName>
    ///     </p:attrNameLst>
    ///   </p:cBhvr>
    ///   <p:by>
    ///     <p:hsl h="0" s="0" l="0"/>
    ///   </p:by>
    /// </p:animClr>
    /// ```
    Hsl(TLByHslColorTransform),
}

impl XsdType for TLByAnimateColorTransform {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "rgb" => Ok(TLByAnimateColorTransform::Rgb(TLByRgbColorTransform::from_xml_element(
                xml_node,
            )?)),
            "hsl" => Ok(TLByAnimateColorTransform::Hsl(TLByHslColorTransform::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "TLByAnimateColorTransform",
            ))),
        }
    }
}

impl XsdChoice for TLByAnimateColorTransform {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "rgb" | "hsl" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLByRgbColorTransform {
    /// This attribute specifies a red component luminance as a percentage. Values are in the range [-100%, 100%].
    pub r: FixedPercentage,
    /// This attribute specifies a green component luminance as a percentage. Values are in the range [-100%, 100%].
    pub g: FixedPercentage,
    /// This attribute specifies a blue component luminance as a percentage. Values are in the range [-100%, 100%].
    pub b: FixedPercentage,
}

impl TLByRgbColorTransform {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut r = None;
        let mut g = None;
        let mut b = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r" => r = Some(value.parse()?),
                "g" => g = Some(value.parse()?),
                "b" => b = Some(value.parse()?),
                _ => (),
            }
        }

        let r = r.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let g = g.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "g"))?;
        let b = b.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        Ok(Self { r, g, b })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TLByHslColorTransform {
    /// Specifies hue as an angle. The values range from [0, 360] degrees
    pub h: Angle,
    /// Specifies a lightness as a percentage. The values are in the range [-100%, 100%].
    pub s: FixedPercentage,
    /// Specifies a saturation as a percentage. The values are in the range [-100%, 100%].
    pub l: FixedPercentage,
}

impl TLByHslColorTransform {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut h = None;
        let mut s = None;
        let mut l = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "h" => h = Some(value.parse()?),
                "s" => s = Some(value.parse()?),
                "l" => l = Some(value.parse()?),
                _ => (),
            }
        }

        let h = h.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "h"))?;
        let s = s.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "s"))?;
        let l = l.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?;

        Ok(Self { h, s, l })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Build {
    /// This element specifies how to build paragraph level properties.
    ///
    /// # Xml example
    ///
    /// Consider having animation applied only to 1st level paragraphs. The <bldP> element should be used
    /// as follows:
    ///
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldP spid="3" grpId="0" build="p"/>
    /// </p:bldLst>
    /// ```
    Paragraph(Box<TLBuildParagraph>),
    /// This element specifies how to build the animation for a diagram.
    ///
    /// # Xml example
    ///
    /// Consider the following example where a chart is specified to be animated by category rather than as
    /// one entity. Thus, the bldChart element should be used as follows:
    ///
    /// ```xml
    /// <p:bdldLst>
    ///   <p:bldGraphic spid="4" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    Diagram(Box<TLBuildDiagram>),
    /// This element describes animation an a embedded Chart.
    ///
    /// # Xml example
    ///
    /// Consider displaying animation on a embedded graphical chart. The <bldOleChart>element should be
    /// use as follows:
    ///
    /// ```xml
    /// <p:bldLst>
    ///   <p:bldOleChart spid="1025" grpId="0"/>
    /// </p:bldLst>
    /// ```
    OleChart(Box<TLOleBuildChart>),
    /// This element specifies how to build a graphical element.
    ///
    /// # Xml example
    ///
    /// Consider having a chart graphical element appear as a whole as opposed to by a category. The
    /// <bldGraphic> element should be used as follows:
    ///
    /// ```xml
    /// <p:bldLdst>
    ///   <p:bldGraphic spid="3" grpId="0">
    ///     <p:bldSub>
    ///       <a:bldChart bld="category"/>
    ///     </p:bldSub>
    ///   </p:bldGraphic>
    /// </p:bldLst>
    /// ```
    Graphic(Box<TLGraphicalObjectBuild>),
}

impl XsdType for Build {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bldP" => Ok(Build::Paragraph(Box::new(TLBuildParagraph::from_xml_element(
                xml_node,
            )?))),
            "bldDgm" => Ok(Build::Diagram(Box::new(TLBuildDiagram::from_xml_element(xml_node)?))),
            "bldOleChart" => Ok(Build::OleChart(Box::new(TLOleBuildChart::from_xml_element(xml_node)?))),
            "bldGraphic" => Ok(Build::Graphic(Box::new(TLGraphicalObjectBuild::from_xml_element(
                xml_node,
            )?))),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "CT_BuildList",
            ))),
        }
    }
}

impl XsdChoice for Build {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "bldP" | "bldDgm" | "bldOleChart" | "bldGraphic" => true,
            _ => false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    impl IndexRange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} st="0" end="5"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { start: 0, end: 5 }
        }
    }

    #[test]
    pub fn test_index_range_from_xml() {
        let xml = IndexRange::test_xml("index");
        assert_eq!(
            IndexRange::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            IndexRange::test_instance(),
        );
    }

    impl TLCommonBehaviorData {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} additive="sum" accumulate="none" xfrmType="pt" from="Example" to="Example"
                by="Example" rctx="Example" override="childStyle">
                {}
                <tgtEl>
                    <sldTgt />
                </tgtEl>
                <attrNameLst>
                    <attrName>style.fontSize</p:attrName>
                </attrNameLst>
            </{node_name}>"#,
                TLCommonTimeNodeData::test_xml_non_recursive("cTn"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                additive: Some(TLBehaviorAdditiveType::Sum),
                accumulate: Some(TLBehaviorAccumulateType::None),
                transform_type: Some(TLBehaviorTransformType::Point),
                from: Some(String::from("Example")),
                to: Some(String::from("Example")),
                by: Some(String::from("Example")),
                runtime_context: Some(String::from("Example")),
                override_type: Some(TLBehaviorOverrideType::ChildStyle),
                common_time_node_data: Box::new(TLCommonTimeNodeData::test_instance_non_recursive()),
                target_element: TLTimeTargetElement::SlideTarget,
                attr_name_list: Some(vec![String::from("style.fontSize")]),
            }
        }
    }

    #[test]
    pub fn test_tl_common_behavior_data_from_xml() {
        let xml = TLCommonBehaviorData::test_xml("tlCommonBehaviorData");
        assert_eq!(
            TLCommonBehaviorData::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLCommonBehaviorData::test_instance()
        );
    }

    impl TLCommonMediaNodeData {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} vol="50000" mute="false" numSld="1" showWhenStopped="true">
                {}
                <tgtEl>
                    <sldTgt />
                </tgtEl>
            </{node_name}>"#,
                TLCommonTimeNodeData::test_xml_non_recursive("cTn"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                volume: Some(50_000.0),
                mute: Some(false),
                number_of_slides: Some(1),
                show_when_stopped: Some(true),
                common_time_node_data: Box::new(TLCommonTimeNodeData::test_instance_non_recursive()),
                target_element: TLTimeTargetElement::SlideTarget,
            }
        }
    }

    #[test]
    pub fn test_tl_common_media_node_data_from_xml() {
        let xml = TLCommonMediaNodeData::test_xml("tlCommonMediaNodeData");
        assert_eq!(
            TLCommonMediaNodeData::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLCommonMediaNodeData::test_instance(),
        );
    }

    impl TLPoint {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} x="100000" y="100000"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                x: 100_000.0,
                y: 100_000.0,
            }
        }
    }

    #[test]
    pub fn test_tl_point_from_xml() {
        let xml = TLPoint::test_xml("tlPoint");
        assert_eq!(
            TLPoint::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLPoint::test_instance(),
        );
    }

    impl TLTimeCondition {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} evt="onClick" delay="1000">
                <tgtEl>
                    <sldTgt/>
                </tgtEl>
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                trigger_event: Some(TLTriggerEvent::OnClick),
                delay: Some(TLTime::TimePoint(1000)),
                trigger: Some(TLTimeConditionTriggerGroup::TargetElement(
                    TLTimeTargetElement::SlideTarget,
                )),
            }
        }
    }

    #[test]
    pub fn test_tl_time_condition_from_xml() {
        let xml = TLTimeCondition::test_xml("tlTimeCondition");
        assert_eq!(
            TLTimeCondition::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTimeCondition::test_instance(),
        );
    }

    impl TLTimeConditionList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                TLTimeCondition::test_xml("cond"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![TLTimeCondition::test_instance()])
        }
    }

    #[test]
    pub fn test_tl_time_condition_list_from_xml() {
        let xml = TLTimeConditionList::test_xml("tlTimeConditionList");
        assert_eq!(
            TLTimeConditionList::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTimeConditionList::test_instance(),
        );
    }

    impl TLTimeNodeList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TLCommonTimeNodeData::test_xml_non_recursive("par"),
                TLTimeNodeSequence::test_xml("seq"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![
                TimeNodeGroup::Parallel(Box::new(TLCommonTimeNodeData::test_instance_non_recursive())),
                TimeNodeGroup::Sequence(Box::new(TLTimeNodeSequence::test_instance())),
            ])
        }
    }

    #[test]
    pub fn test_tl_time_node_list_from_xml() {
        let xml = TLTimeNodeList::test_xml("tlTimeNodeList");
        assert_eq!(
            TLTimeNodeList::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTimeNodeList::test_instance(),
        );
    }

    impl TLCommonTimeNodeData {
        const TEST_ATTRIBUTES: &'static str = r#"id="1" presetID="1" presetClass="entr" presetSubtype="0" dur="1000"
            repeatCount="1000" repeatDur="1000" spd="100000" accel="0" decel="0" autoRev="false" restart="never"
            fill="hold" syncBehavior="locked" tmFilter="Example filter" evtFilter="Example filter" display="true"
            masterRel="sameClick" bldLvl="1" grpId="1" afterEffect="true" nodeType="mainSequence" nodePh="false""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                TLTimeConditionList::test_xml("stCondLst"),
                TLTimeConditionList::test_xml("endCondLst"),
                TLTimeCondition::test_xml("endSync"),
                TLIterateData::test_xml("iterate"),
                TLTimeNodeList::test_xml("childTnLst"),
                TLTimeNodeList::test_xml("subTnLst"),
                node_name = node_name,
            )
        }

        pub fn test_xml_non_recursive(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}></{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start_condition_list: Some(TLTimeConditionList::test_instance()),
                end_condition_list: Some(TLTimeConditionList::test_instance()),
                end_sync: Some(TLTimeCondition::test_instance()),
                iterate: Some(TLIterateData::test_instance()),
                child_time_node_list: Some(TLTimeNodeList::test_instance()),
                sub_time_node_list: Some(TLTimeNodeList::test_instance()),
                ..Self::test_instance_non_recursive()
            }
        }

        pub fn test_instance_non_recursive() -> Self {
            Self {
                id: Some(1),
                preset_id: Some(1),
                preset_class: Some(TLTimeNodePresetClassType::Entrance),
                preset_subtype: Some(0),
                duration: Some(TLTime::TimePoint(1000)),
                repeat_count: Some(TLTime::TimePoint(1000)),
                repeat_duration: Some(TLTime::TimePoint(1000)),
                speed: Some(100_000.0),
                acceleration: Some(0.0),
                deceleration: Some(0.0),
                auto_reverse: Some(false),
                restart_type: Some(TLTimeNodeRestartType::Never),
                fill_type: Some(TLTimeNodeFillType::Hold),
                sync_behavior: Some(TLTimeNodeSyncType::Locked),
                time_filter: Some(String::from("Example filter")),
                event_filter: Some(String::from("Example filter")),
                display: Some(true),
                master_relationship: Some(TLTimeNodeMasterRelation::SameClick),
                build_level: Some(1),
                group_id: Some(1),
                after_effect: Some(true),
                node_type: Some(TLTimeNodeType::MainSequence),
                node_placeholder: Some(false),
                ..Default::default()
            }
        }
    }

    #[test]
    pub fn test_tl_common_time_node_data_from_xml() {
        let xml = TLCommonTimeNodeData::test_xml("tlCommonTimeNodeData");
        assert_eq!(
            TLCommonTimeNodeData::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLCommonTimeNodeData::test_instance(),
        );
    }

    impl TLTimeNodeSequence {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} concurrent="false" prevAc="none" nextAc="none">
                {}
                {}
                {}
            </{node_name}>"#,
                TLCommonTimeNodeData::test_xml_non_recursive("cTn"),
                TLTimeConditionList::test_xml("prevCondLst"),
                TLTimeConditionList::test_xml("nextCondLst"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                concurrent: Some(false),
                prev_action_type: Some(TLPreviousActionType::None),
                next_action_type: Some(TLNextActionType::None),
                common_time_node_data: Box::new(TLCommonTimeNodeData::test_instance_non_recursive()),
                prev_condition_list: Some(TLTimeConditionList::test_instance()),
                next_condition_list: Some(TLTimeConditionList::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tl_time_node_sequence_from_xml() {
        let xml = TLTimeNodeSequence::test_xml("tlTimeNodeSequence");
        assert_eq!(
            TLTimeNodeSequence::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTimeNodeSequence::test_instance(),
        );
    }

    impl TLIterateData {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} type="el" backwards="false">
                <tmAbs val="10000" />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                iterate_type: Some(IterateType::Element),
                backwards: Some(false),
                interval: TLIterateDataChoice::Absolute(TLTime::TimePoint(10_000)),
            }
        }
    }

    #[test]
    pub fn test_tl_iterate_data_from_xml() {
        let xml = TLIterateData::test_xml("tlIterateData");
        assert_eq!(
            TLIterateData::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLIterateData::test_instance(),
        );
    }

    impl TLTemplate {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} lvl="0">{}</{node_name}>"#,
                TLTimeNodeList::test_xml("tnLst"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                level: Some(0),
                time_node_list: TLTimeNodeList::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tl_template_from_xml() {
        let xml = TLTemplate::test_xml("tlTemplate");
        assert_eq!(
            TLTemplate::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTemplate::test_instance(),
        );
    }

    impl TLTemplateList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                TLTemplate::test_xml("tmpl"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self(vec![TLTemplate::test_instance()])
        }
    }

    #[test]
    pub fn test_tl_template_list_from_xml() {
        let xml = TLTemplateList::test_xml("tlTemplateList");
        assert_eq!(
            TLTemplateList::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLTemplateList::test_instance(),
        );
    }

    impl TLBuildCommonAttributes {
        const TEST_ATTRIBUTES: &'static str = r#"spid="0" grpId="0" uiExpand="true""#;

        pub fn test_instance() -> Self {
            Self {
                shape_id: 0,
                group_id: 0,
                ui_expand: Some(true),
            }
        }
    }

    impl TLBuildParagraph {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} {} build="whole" bldLvl="0" animBg="true" autoUpdateAnimBg="true" rev="false" advAuto="1000">
                {}
            </{node_name}>"#,
                TLBuildCommonAttributes::TEST_ATTRIBUTES,
                TLTemplateList::test_xml("tmplLst"),
                node_name=node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                build_common: TLBuildCommonAttributes::test_instance(),
                build_type: Some(TLParaBuildType::Whole),
                build_level: Some(0),
                animate_bg: Some(true),
                auto_update_anim_bg: Some(true),
                reverse: Some(false),
                auto_advance_time: Some(TLTime::TimePoint(1000)),
                template_list: Some(TLTemplateList::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tl_build_paragraph_from_xml() {
        let xml = TLBuildParagraph::test_xml("tlBuildParagraph");
        assert_eq!(
            TLBuildParagraph::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLBuildParagraph::test_instance(),
        );
    }

    impl TLBuildDiagram {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} bld="whole"></{node_name}>"#,
                TLBuildCommonAttributes::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                build_common: TLBuildCommonAttributes::test_instance(),
                build_type: Some(TLDiagramBuildType::Whole),
            }
        }
    }

    #[test]
    pub fn test_tl_build_diagram_from_xml() {
        let xml = TLBuildDiagram::test_xml("tlBuildDiagram");
        assert_eq!(
            TLBuildDiagram::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLBuildDiagram::test_instance(),
        );
    }

    impl TLOleBuildChart {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} bld="allAtOnce" animBg="true"></{node_name}>"#,
                TLBuildCommonAttributes::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                build_common: TLBuildCommonAttributes::test_instance(),
                build_type: Some(TLOleChartBuildType::AllAtOnce),
                animate_bg: Some(true),
            }
        }
    }

    #[test]
    pub fn test_tl_ole_build_chart_from_xml() {
        let xml = TLOleBuildChart::test_xml("tlOleBuildChart");
        assert_eq!(
            TLOleBuildChart::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLOleBuildChart::test_instance(),
        );
    }

    impl TLGraphicalObjectBuild {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                <bldAsOne />
            </{node_name}>"#,
                TLBuildCommonAttributes::TEST_ATTRIBUTES,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                build_common: TLBuildCommonAttributes::test_instance(),
                build_choice: TLGraphicalObjectBuildChoice::BuildAsOne,
            }
        }
    }

    #[test]
    pub fn test_tl_graphical_object_build_from_xml() {
        let xml = TLGraphicalObjectBuild::test_xml("tlGraphicalObjectBuild");
        assert_eq!(
            TLGraphicalObjectBuild::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
            TLGraphicalObjectBuild::test_instance(),
        );
    }

    // impl TLAnimateBehavior {
    //     pub fn test_xml(node_name: &'static str) -> String {
    //         format!(r#"<{node_name} by="Example" from="Example" to="Example" calcmode="fmla" valueType="str">
    //             {}
    //             {}
    //         </{node_name}>"#,
    //             TLCommonBehaviorData::test_xml("cBhvr"),
    //             TLTimeAnimateValueList::test_xml("tavLst"),
    //             node_name=node_name,
    //         )
    //     }

    //     pub fn test_instance() -> Self {
    //         Self {
    //             by: Some(String::from("Example")),
    //             from: Some(String::from("Example")),
    //             to: Some(String::from("Example")),
    //             calc_mode: Some(TLAnimateBehaviorCalcMode::Formula),
    //             value_type: Some(TLAnimateBehaviorValueType::String),
    //             common_behavior_data: Box::new(TLCommonBehaviorData::test_instance()),
    //             time_animate_value_list: Some(TLTimeAnimateValueList::test_instance()),
    //         }
    //     }
    // }

    // #[test]
    // pub fn test_tl_animate_behavior_from_xml() {
    //     let xml = TLAnimateBehavior::test_xml("tlAnimateBehavior");
    //     assert_eq!(
    //         TLAnimateBehavior::from_xml_element(&XmlNode::from_str(&xml).unwrap()).unwrap(),
    //         TLAnimateBehavior::test_instance(),
    //     );
    // }
}
