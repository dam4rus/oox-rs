use crate::{
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use super::simpletypes::{
    AdjAngle, AdjCoordinate, GeomGuideFormula, GeomGuideName, PathFillMode, PositiveCoordinate, ShapeType,
    TextShapeType,
};
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, PartialEq)]
pub struct GeomRect {
    /// Specifies the x coordinate of the left edge for a shape text rectangle. The units for this
    /// edge is specified in EMUs as the positioning here is based on the shape coordinate
    /// system. The width and height for this coordinate system are specified within the ext
    /// transform element.
    pub left: AdjCoordinate,

    /// Specifies the y coordinate of the top edge for a shape text rectangle. The units for this
    /// edge is specified in EMUs as the positioning here is based on the shape coordinate
    /// system. The width and height for this coordinate system are specified within the ext
    /// transform element.
    pub top: AdjCoordinate,

    /// Specifies the x coordinate of the right edge for a shape text rectangle. The units for this
    /// edge is specified in EMUs as the positioning here is based on the shape coordinate
    /// system. The width and height for this coordinate system are specified within the ext
    /// transform element.
    pub right: AdjCoordinate,

    /// Specifies the y coordinate of the bottom edge for a shape text rectangle. The units for
    /// this edge is specified in EMUs as the positioning here is based on the shape coordinate
    /// system. The width and height for this coordinate system are specified within the ext
    /// transform element.
    pub bottom: AdjCoordinate,
}

impl GeomRect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut left = None;
        let mut top = None;
        let mut right = None;
        let mut bottom = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "l" => left = Some(value.parse()?),
                "t" => top = Some(value.parse()?),
                "r" => right = Some(value.parse()?),
                "b" => bottom = Some(value.parse()?),
                _ => (),
            }
        }

        let left = left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?;
        let top = top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "t"))?;
        let right = right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let bottom = bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        Ok(Self {
            left,
            top,
            right,
            bottom,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PolarAdjustHandle {
    /// Specifies the name of the guide that is updated with the adjustment radius from this
    /// adjust handle.
    pub guide_reference_radial: Option<GeomGuideName>,

    /// Specifies the name of the guide that is updated with the adjustment angle from this
    /// adjust handle.
    pub guide_reference_angle: Option<GeomGuideName>,

    /// Specifies the minimum radial position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move radially. That
    /// is the maxR and minR are equal.
    pub min_radial: Option<AdjCoordinate>,

    /// Specifies the maximum radial position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move radially. That
    /// is the maxR and minR are equal.
    pub max_radial: Option<AdjCoordinate>,

    /// Specifies the minimum angle position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move angularly.
    /// That is the maxAng and minAng are equal.
    pub min_angle: Option<AdjAngle>,

    /// Specifies the maximum angle position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move angularly.
    /// That is the maxAng and minAng are equal.
    pub max_angle: Option<AdjAngle>,
    pub position: AdjPoint2D,
}

impl PolarAdjustHandle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut guide_reference_radial = None;
        let mut guide_reference_angle = None;
        let mut min_radial = None;
        let mut max_radial = None;
        let mut min_angle = None;
        let mut max_angle = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "gdRefR" => guide_reference_radial = Some(value.clone()),
                "gdRefAng" => guide_reference_angle = Some(value.clone()),
                "minR" => min_radial = Some(value.parse()?),
                "maxR" => max_radial = Some(value.parse()?),
                "minAng" => min_angle = Some(value.parse()?),
                "maxAng" => max_angle = Some(value.parse()?),
                _ => (),
            }
        }

        let position = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pos")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "pos")))
            .and_then(AdjPoint2D::from_xml_element)?;

        Ok(Self {
            guide_reference_radial,
            guide_reference_angle,
            min_radial,
            max_radial,
            min_angle,
            max_angle,
            position,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct XYAdjustHandle {
    /// Specifies the name of the guide that is updated with the adjustment x position from this
    /// adjust handle.
    pub guide_reference_x: Option<GeomGuideName>,

    /// Specifies the name of the guide that is updated with the adjustment y position from this
    /// adjust handle.
    pub guide_reference_y: Option<GeomGuideName>,

    /// Specifies the minimum horizontal position that is allowed for this adjustment handle. If
    /// this attribute is omitted, then it is assumed that this adjust handle cannot move in the x
    /// direction. That is the maxX and minX are equal.
    pub min_x: Option<AdjCoordinate>,

    /// Specifies the maximum horizontal position that is allowed for this adjustment handle. If
    /// this attribute is omitted, then it is assumed that this adjust handle cannot move in the x
    /// direction. That is the maxX and minX are equal.
    pub max_x: Option<AdjCoordinate>,

    /// Specifies the minimum vertical position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move in the y
    /// direction. That is the maxY and minY are equal.
    pub min_y: Option<AdjCoordinate>,

    /// Specifies the maximum vertical position that is allowed for this adjustment handle. If this
    /// attribute is omitted, then it is assumed that this adjust handle cannot move in the y
    /// direction. That is the maxY and minY are equal.
    pub max_y: Option<AdjCoordinate>,

    /// Specifies a position coordinate within the shape bounding box. It should be noted that this coordinate is placed
    /// within the shape bounding box using the transform coordinate system which is also called the shape coordinate
    /// system, as it encompasses the entire shape. The width and height for this coordinate system are specified within
    /// the ext transform element.
    ///
    /// # Note
    ///
    /// When specifying a point coordinate in path coordinate space it should be noted that the top left of the
    /// coordinate space is x=0, y=0 and the coordinate points for x grow to the right and for y grow down.
    ///
    /// # Xml example
    ///
    /// To highlight the differences in the coordinate systems consider the drawing of the following triangle.
    /// Notice that the dimensions of the triangle are specified using the shape coordinate system with EMUs as the
    /// units via the ext transform element. Thus we see this shape is 1705233 EMUs wide by 679622 EMUs tall.
    /// However when looking at how the path for this shape is drawn we see that the x and y values fall between 0 and
    /// 2. This is because the path coordinate system has the arbitrary dimensions of 2 for the width and 2 for the
    /// height. Thus we see that a y coordinate of 2 within the path coordinate system specifies a y coordinate of
    /// 679622 within the shape coordinate system for this particular case.
    ///
    /// ```xml
    /// <a:xfrm>
    ///   <a:off x="3200400" y="1600200"/>
    ///   <a:ext cx="1705233" cy="679622"/>
    /// </a:xfrm>
    /// <a:custGeom>
    ///   <a:avLst/>
    ///   <a:gdLst/>
    ///   <a:ahLst/>
    ///   <a:cxnLst/>
    ///   <a:rect l="0" t="0" r="0" b="0"/>
    ///   <a:pathLst>
    ///     <a:path w="2" h="2">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="2"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="2" y="2"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    pub position: AdjPoint2D,
}

impl XYAdjustHandle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut guide_reference_x = None;
        let mut guide_reference_y = None;
        let mut min_x = None;
        let mut max_x = None;
        let mut min_y = None;
        let mut max_y = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "gdRefX" => guide_reference_x = Some(value.clone()),
                "gdRefY" => guide_reference_y = Some(value.clone()),
                "minX" => min_x = Some(value.parse()?),
                "maxX" => max_x = Some(value.parse()?),
                "minY" => min_y = Some(value.parse()?),
                "maxY" => max_y = Some(value.parse()?),
                _ => (),
            }
        }

        let position = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pos")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "pos")))
            .and_then(AdjPoint2D::from_xml_element)?;

        Ok(Self {
            guide_reference_x,
            guide_reference_y,
            min_x,
            max_x,
            min_y,
            max_y,
            position,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum AdjustHandle {
    /// This element specifies an XY-based adjust handle for a custom shape. The position of this adjust handle is
    /// specified by the corresponding pos child element. The allowed adjustment of this adjust handle are specified via
    /// it's min and max type attributes. Based on the adjustment of this adjust handle certain corresponding guides are
    /// updated to contain these values.
    XY(Box<XYAdjustHandle>),

    /// This element specifies a polar adjust handle for a custom shape. The position of this adjust handle is specified by
    /// the corresponding pos child element. The allowed adjustment of this adjust handle are specified via it's min and
    /// max attributes. Based on the adjustment of this adjust handle certain corresponding guides are updated to
    /// contain these values.
    Polar(Box<PolarAdjustHandle>),
}

impl XsdType for AdjustHandle {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "ahXY" => Ok(AdjustHandle::XY(Box::new(XYAdjustHandle::from_xml_element(xml_node)?))),
            "ahPolar" => Ok(AdjustHandle::Polar(Box::new(PolarAdjustHandle::from_xml_element(
                xml_node,
            )?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "AdjustHandle").into()),
        }
    }
}

impl XsdChoice for AdjustHandle {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "ahXY" | "ahPolar" => true,
            _ => false,
        }
    }
}

/// This element specifies an x-y coordinate within the path coordinate space. This coordinate space is determined
/// by the width and height attributes defined within the path element. A point is utilized by one of it's parent
/// elements to specify the next point of interest in custom geometry shape. Depending on the parent element used
/// the point can either have a line drawn to it or the cursor can simply be moved to this new location.
///
/// Specifies a position coordinate within the shape bounding box. It should be noted that this coordinate is placed
/// within the shape bounding box using the transform coordinate system which is also called the shape coordinate
/// system, as it encompasses the entire shape. The width and height for this coordinate system are specified within
/// the ext transform element.
///
/// # Note
///
/// When specifying a point coordinate in path coordinate space it should be noted that the top left of the
/// coordinate space is x=0, y=0 and the coordinate points for x grow to the right and for y grow down.
///
/// # Xml example
///
/// To highlight the differences in the coordinate systems consider the drawing of the following triangle.
/// Notice that the dimensions of the triangle are specified using the shape coordinate system with EMUs as the
/// units via the ext transform element. Thus we see this shape is 1705233 EMUs wide by 679622 EMUs tall.
/// However when looking at how the path for this shape is drawn we see that the x and y values fall between 0 and
/// 2. This is because the path coordinate system has the arbitrary dimensions of 2 for the width and 2 for the
/// height. Thus we see that a y coordinate of 2 within the path coordinate system specifies a y coordinate of
/// 679622 within the shape coordinate system for this particular case.
///
/// ```xml
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst/>
///   <a:ahLst/>
///   <a:cxnLst/>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="2" h="2">
///       <a:moveTo>
///         <a:pt x="0" y="2"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2" y="2"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct AdjPoint2D {
    /// Specifies the x coordinate for this position coordinate. The units for this coordinate space
    /// are defined by the width of the path coordinate system. This coordinate system is
    /// overlayed on top of the shape coordinate system thus occupying the entire shape
    /// bounding box. Because the units for within this coordinate space are determined by the
    /// path width and height an exact measurement unit cannot be specified here.
    pub x: AdjCoordinate,

    /// Specifies the y coordinate for this position coordinate. The units for this coordinate space
    /// are defined by the height of the path coordinate system. This coordinate system is
    /// overlayed on top of the shape coordinate system thus occupying the entire shape
    /// bounding box. Because the units for within this coordinate space are determined by the
    /// path width and height an exact measurement unit cannot be specified here.
    pub y: AdjCoordinate,
}

impl AdjPoint2D {
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
pub struct Path2DArcTo {
    /// This attribute specifies the width radius of the supposed circle being used to draw the
    /// arc. This gives the circle a total width of (2 * wR). This total width could also be called it's
    /// horizontal diameter as it is the diameter for the x axis only.
    pub width_radius: AdjCoordinate,

    /// This attribute specifies the height radius of the supposed circle being used to draw the
    /// arc. This gives the circle a total height of (2 * hR). This total height could also be called
    /// it's vertical diameter as it is the diameter for the y axis only.
    pub height_radius: AdjCoordinate,

    /// Specifies the start angle for an arc. This angle specifies what angle along the supposed
    /// circle path is used as the start position for drawing the arc. This start angle is locked to
    /// the last known pen position in the shape path. Thus guaranteeing a continuos shape
    /// path.
    pub start_angle: AdjAngle,

    /// Specifies the swing angle for an arc. This angle specifies how far angle-wise along the
    /// supposed cicle path the arc is extended. The extension from the start angle is always in
    /// the clockwise direction around the supposed circle.
    pub swing_angle: AdjAngle,
}

impl Path2DArcTo {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut width_radius = None;
        let mut height_radius = None;
        let mut start_angle = None;
        let mut swing_angle = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "wR" => width_radius = Some(value.parse()?),
                "hR" => height_radius = Some(value.parse()?),
                "stAng" => start_angle = Some(value.parse()?),
                "swAng" => swing_angle = Some(value.parse()?),
                _ => (),
            }
        }

        let width_radius = width_radius.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wR"))?;
        let height_radius = height_radius.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hR"))?;
        let start_angle = start_angle.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "stAng"))?;
        let swing_angle = swing_angle.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "swAng"))?;

        Ok(Self {
            width_radius,
            height_radius,
            start_angle,
            swing_angle,
        })
    }
}

/// This element specifies a creation path consisting of a series of moves, lines and curves that when combined
/// forms a geometric shape. This element is only utilized if a custom geometry is specified.
///
/// # Note
///
/// Since multiple paths are allowed the rules for drawing are that the path specified later in the pathLst is
/// drawn on top of all previous paths.
///
/// # Xml example
///
/// ```xml
/// <a:custGeom>
///   <a:pathLst>
///     <a:path w="2824222" h="590309">
///       <a:moveTo>
///         <a:pt x="0" y="428263"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="1620455" y="590309"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="2824222" y="173620"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1562582" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
///
/// In the above example there is specified a four sided geometric shape that has all straight sides. While we only
/// see three lines being drawn via the lnTo element there are actually four sides because the last point of
/// (x=1562585, y=0) is connected to the first point in the creation path via a lnTo element
#[derive(Default, Debug, Clone, PartialEq)]
pub struct Path2D {
    /// Specifies the width, or maximum x coordinate that should be used for within the path
    /// coordinate system. This value determines the horizontal placement of all points within
    /// the corresponding path as they are all calculated using this width attribute as the max x
    /// coordinate.
    ///
    /// Defaults to 0
    pub width: Option<PositiveCoordinate>,

    /// Specifies the height, or maximum y coordinate that should be used for within the path
    /// coordinate system. This value determines the vertical placement of all points within the
    /// corresponding path as they are all calculated using this height attribute as the max y
    /// coordinate.
    ///
    /// Defaults to 0
    pub height: Option<PositiveCoordinate>,

    /// Specifies how the corresponding path should be filled. If this attribute is omitted, a value
    /// of "norm" is assumed.
    ///
    /// Defaults to PathFillMode::Norm
    pub fill_mode: Option<PathFillMode>,

    /// Specifies if the corresponding path should have a path stroke shown. This is a boolean
    /// value that affect the outline of the path. If this attribute is omitted, a value of true is
    /// assumed.
    ///
    /// Defaults to true
    pub stroke: Option<bool>,

    /// Specifies that the use of 3D extrusions are possible on this path. This allows the
    /// generating application to know whether 3D extrusion can be applied in any form. If this
    /// attribute is omitted then a value of 0, or false is assumed.
    ///
    /// Defaults to true
    pub extrusion_ok: Option<bool>,
    pub commands: Vec<Path2DCommand>,
}

impl Path2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "w" => instance.width = Some(value.parse()?),
                    "h" => instance.height = Some(value.parse()?),
                    "fill" => instance.fill_mode = Some(value.parse()?),
                    "stroke" => instance.stroke = Some(parse_xml_bool(value)?),
                    "extrusionOk" => instance.extrusion_ok = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|mut instance| {
                instance.commands = xml_node
                    .child_nodes
                    .iter()
                    .filter_map(Path2DCommand::try_from_xml_element)
                    .collect::<Result<Vec<_>>>()?;

                Ok(instance)
            })
    }
}

/// This element specifies the precense of a shape guide that is used to govern the geometry of the specified shape.
/// A shape guide consists of a formula and a name that the result of the formula is assigned to. Recognized
/// formulas are listed with the fmla attribute documentation for this element.
///
/// # Note
///
/// The order in which guides are specified determines the order in which their values are calculated. For
/// instance it is not possible to specify a guide that uses another guides result when that guide has not yet been
/// calculated.
///
/// # Example
///
/// Consider the case where the user would like to specify a triangle with it's bottom edge defined not by
/// static points but by using a varying parameter, namely an guide. Consider the diagrams and DrawingML shown
/// below. This first triangle has been drawn with a bottom edge that is equal to the 2/3 the value of the shape
/// height. Thus we see in the figure below that the triangle appears to occupy 2/3 of the vertical space within the
/// shape bounding box.
///
/// ```xml
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst>
///     <a:gd name="myGuide" fmla="*/ h 2 3"/>
///   </a:gdLst>
///   <a:ahLst/>
///   <a:cxnLst/>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="1705233" h="679622">
///       <a:moveTo>
///         <a:pt x="0" y="myGuide"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="1705233" y="myGuide"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="852616" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
///
/// If however we change the guide to half that, namely 1/3. Then we see the entire bottom edge of the triangle
/// move to now only occupy 1/3 of the toal space within the shape bounding box. This is because both of the
/// bottom points in this triangle depend on this guide for their coordinate positions.
///
/// ```xml
/// <a:gdLst>
///   <a:gd name="myGuide" fmla="*/ h 1 3"/>
/// </a:gdLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct GeomGuide {
    /// Specifies the name that is used to reference to this guide. This name can be used just as a
    /// variable would within an equation. That is this name can be substituted for literal values
    /// within other guides or the specification of the shape path.
    pub name: GeomGuideName,

    /// Specifies the formula that is used to calculate the value for a guide. Each formula has a
    /// certain number of arguments and a specific set of operations to perform on these
    /// arguments in order to generate a value for a guide. There are a total of 17 different
    /// formulas available. These are shown below with the usage for each defined.
    ///
    /// * **('\*/') - Multiply Divide Formula**
    ///
    ///     Arguments: 3 (fmla="*/ x y z")
    ///
    ///     Usage: "*/ x y z" = ((x * y) / z) = value of this guide
    /// * **('+-') - Add Subtract Formula**
    ///
    ///     Arguments: 3 (fmla="+- x y z")
    ///
    ///     Usage: "+- x y z" = ((x + y) - z) = value of this guide
    ///
    /// * **('+/') - Add Divide Formula**
    ///
    ///     Arguments: 3 (fmla="+/ x y z")
    ///
    ///     Usage: "+/ x y z" = ((x + y) / z) = value of this guide
    ///
    /// * **('?:') - If Else Formula**
    ///
    ///     Arguments: 3 (fmla="?: x y z")
    ///
    ///     Usage: "?: x y z" = if (x > 0), then y = value of this guide,  
    ///     else z = value of this guide
    ///
    /// * **('abs') - Absolute Value Formula**
    ///
    ///     Arguments: 1 (fmla="abs x")
    ///
    ///     Usage: "abs x" = if (x < 0), then (-1) * x = value of this guide  
    ///     else x = value of this guide
    ///
    /// * **('at2') - ArcTan Formula**
    ///
    ///     Arguments: 2 (fmla="at2 x y")
    ///
    ///     Usage: "at2 x y" = arctan(y / x) = value of this guide
    ///
    /// * **('cat2') - Cosine ArcTan Formula**
    ///
    ///     Arguments: 3 (fmla="cat2 x y z")
    ///
    ///     Usage: "cat2 x y z" = (x*(cos(arctan(z / y))) = value of this guide
    ///
    /// * **('cos') - Cosine Formula**
    ///
    ///     Arguments: 2 (fmla="cos x y")
    ///
    ///     Usage: "cos x y" = (x * cos( y )) = value of this guide
    ///
    /// * **('max') - Maximum Value Formula**
    ///
    ///     Arguments: 2 (fmla="max x y")
    ///
    ///     Usage: "max x y" = if (x > y), then x = value of this guide  
    ///     else y = value of this guide
    ///
    /// * **('min') - Minimum Value Formula**
    ///
    ///     Arguments: 2 (fmla="min x y")
    ///
    ///     Usage: "min x y" = if (x < y), then x = value of this guide  
    ///     else y = value of this guide
    ///
    /// * **('mod') - Modulo Formula**
    ///
    ///     Arguments: 3 (fmla="mod x y z")
    ///
    ///     Usage: "mod x y z" = sqrt(x^2 + b^2 + c^2) = value of this guide
    ///
    /// * **('pin') - Pin To Formula**
    ///
    ///     Arguments: 3 (fmla="pin x y z")
    ///
    ///     Usage: "pin x y z" = if (y < x), then x = value of this guide  
    ///     else if (y > z), then z = value of this guide  
    ///     else y = value of this guide
    ///
    /// * **('sat2') - Sine ArcTan Formula**
    ///
    ///     Arguments: 3 (fmla="sat2 x y z")
    ///
    ///     Usage: "sat2 x y z" = (x*sin(arctan(z / y))) = value of this guide
    ///
    /// * **('sin') - Sine Formula**
    ///
    ///     Arguments: 2 (fmla="sin x y")
    ///
    ///     Usage: "sin x y" = (x * sin( y )) = value of this guide
    ///
    /// * **('sqrt') - Square Root Formula**
    ///
    ///     Arguments: 1 (fmla="sqrt x")
    ///
    ///     Usage: "sqrt x" = sqrt(x) = value of this guide
    ///
    /// * **('tan') - Tangent Formula**
    ///
    ///     Arguments: 2 (fmla="tan x y")
    ///
    ///     Usage: "tan x y" = (x * tan( y )) = value of this guide
    ///
    /// * **('val') - Literal Value Formula**
    ///
    ///     Arguments: 1 (fmla="val x")
    ///
    ///     Usage: "val x" = x = value of this guide
    ///
    /// # Note
    ///
    /// Guides that have a literal value formula specified via fmla="val x" above should
    /// only be used within the avLst as an adjust value for the shape. This however is not
    /// strictly enforced.
    pub formula: GeomGuideFormula,
}

impl GeomGuide {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut name = None;
        let mut formula = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                "fmla" => formula = Some(value.clone()),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let formula = formula.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "fmla"))?;
        Ok(Self { name, formula })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Path2DCommand {
    /// This element specifies the ending of a series of lines and curves in the creation path of a custom geometric
    /// shape. When this element is encountered, the generating application should consider the corresponding path
    /// closed. That is, any further lines or curves that follow this element should be ignored.
    ///
    /// # Note
    ///
    /// A path can be specified and not closed. A path such as this cannot however have any fill associated with it
    /// as it has not been considered a closed geometric path.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:custGeom>
    ///   <a:pathLst>
    ///     <a:path w="2824222" h="590309">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="428263"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="1620455" y="590309"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="2824222" y="173620"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1562582" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    ///
    /// In the above example there is specified a four sided geometric shape that has all straight sides. While we only
    /// see three lines being drawn via the lnTo element there are actually four sides because the last point of
    /// (x=1562585, y=0) is connected to the first point in the creation path via a lnTo element
    ///
    /// # Note
    ///
    /// When the last point in the creation path does not meet with the first point in the creation path the
    /// generating application should connect the last point with the first via a straight line, thus creating a closed shape
    /// geometry.
    Close,

    /// This element specifies a set of new coordinates to move the shape cursor to. This element is only used for
    /// drawing a custom geometry. When this element is utilized the pt element is used to specify a new set of shape
    /// coordinates that the shape cursor should be moved to. This does not draw a line or curve to this new position
    /// from the old position but simply move the cursor to a new starting position. It is only when a path drawing
    /// element such as lnTo is used that a portion of the path is drawn.
    ///
    /// # Xml example
    ///
    /// Consider the case where a user wishes to begin drawing a custom geometry not at the default starting
    /// coordinates of x=0 , y=0 but at coordinates further inset into the shape coordinate space. The following
    /// DrawingML would specify such a case.
    ///
    /// ```xml
    /// <a:custGeom>
    ///   <a:pathLst>
    ///     <a:path w="2824222" h="590309">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="428263"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="1620455" y="590309"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="2824222" y="173620"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1562582" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    ///
    /// Notice the moveTo element advances the y coordinates before any actual lines are drawn
    MoveTo(AdjPoint2D),

    /// This element specifies the drawing of a straight line from the current pen position to the new point specified.
    /// This line becomes part of the shape geometry, representing a side of the shape. The coordinate system used
    /// when specifying this line is the path coordinate system.
    LineTo(AdjPoint2D),

    /// This element specifies the existence of an arc within a shape path. It draws an arc with the specified parameters
    /// from the current pen position to the new point specified. An arc is a line that is bent based on the shape of a
    /// supposed circle. The length of this arc is determined by specifying both a start angle and an ending angle that
    /// act together to effectively specify an end point for the arc.
    ArcTo(Path2DArcTo),

    /// This element specifies to draw a quadratic bezier curve along the specified points. To specify a quadratic bezier
    /// curve there needs to be 2 points specified. The first is a control point used in the quadratic bezier calculation
    /// and the last is the ending point for the curve. The coordinate system used for this type of curve is the path
    /// coordinate system as this element is path specific.
    QuadBezierTo(AdjPoint2D, AdjPoint2D),

    /// This element specifies to draw a cubic bezier curve along the specified points. To specify a cubic bezier curve
    /// there needs to be 3 points specified. The first two are control points used in the cubic bezier calculation and the
    /// last is the ending point for the curve. The coordinate system used for this kind of curve is the path coordinate
    /// system as this element is path specific.
    CubicBezTo(AdjPoint2D, AdjPoint2D, AdjPoint2D),
}

impl XsdType for Path2DCommand {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let get_point_at = |index| {
            xml_node
                .child_nodes
                .get(index)
                .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "pt")))
                .and_then(AdjPoint2D::from_xml_element)
        };

        match xml_node.local_name() {
            "close" => Ok(Path2DCommand::Close),
            "moveTo" => Ok(Path2DCommand::LineTo(get_point_at(0)?)),
            "lnTo" => Ok(Path2DCommand::LineTo(get_point_at(0)?)),
            "arcTo" => Ok(Path2DCommand::ArcTo(Path2DArcTo::from_xml_element(xml_node)?)),
            "quadBezTo" => Ok(Path2DCommand::QuadBezierTo(get_point_at(0)?, get_point_at(1)?)),
            "cubicBezTo" => Ok(Path2DCommand::CubicBezTo(
                get_point_at(0)?,
                get_point_at(1)?,
                get_point_at(2)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_Path2DCommand",
            ))),
        }
    }
}

impl XsdChoice for Path2DCommand {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "close" | "moveTo" | "lnTo" | "arcTo" | "quadBezTo" | "cubicBezTo" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct GeomGuideList(pub Vec<GeomGuide>);

impl GeomGuideList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        Ok(Self(
            xml_node
                .child_nodes
                .iter()
                .filter(|child_node| child_node.local_name() == "gd")
                .map(GeomGuide::from_xml_element)
                .collect::<Result<Vec<_>>>()?,
        ))
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct CustomGeometry2D {
    /// This element specifies the adjust values that are applied to the specified shape. An adjust value is simply a guide
    /// that has a value based formula specified. That is, no calculation takes place for an adjust value guide. Instead,
    /// this guide specifies a parameter value that is used for calculations within the shape guides.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:xfrm>
    ///   <a:off x="3200400" y="1600200"/>
    ///   <a:ext cx="1705233" cy="679622"/>
    /// </a:xfrm>
    /// <a:custGeom>
    ///   <a:avLst>
    ///     <a:gd name="myGuide" fmla="val 2"/>
    ///   </a:avLst>
    ///   <a:gdLst/>
    ///   <a:ahLst/>
    ///   <a:cxnLst/>
    ///   <a:rect l="0" t="0" r="0" b="0"/>
    ///   <a:pathLst>
    ///     <a:path w="2" h="2">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="myGuide"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="2" y="myGuide"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    pub adjust_value_list: Option<GeomGuideList>,

    /// This element specifies all the guides that are used for this shape. A guide is specified by the gd element and
    /// defines a calculated value that can be used for the construction of the corresponding shape.
    ///
    /// # Note
    ///
    /// Guides that have a literal value formula specified via fmla="val x" above should only be used within the
    /// adjust_value_list as an adjust value for the shape. This however is not strictly enforced.
    pub guide_list: Option<GeomGuideList>,

    /// This element specifies the adjust handles that are applied to a custom geometry. These adjust handles specify
    /// points within the geometric shape that can be used to perform certain transform operations on the shape.
    ///
    /// # Example
    ///
    /// Consider the scenario where a custom geometry, an arrow in this case, has been drawn and adjust
    /// handles have been placed at the top left corner of both the arrow head and arrow body. The user interface can
    /// then be made to transform only certain parts of the shape by using the corresponding adjust handle.
    ///
    /// For instance if the user wished to change only the width of the arrow head then they would use the adjust
    /// handle located on the top left of the arrow head.
    pub adjust_handle_list: Option<Vec<AdjustHandle>>,

    /// This element specifies all the connection sites that are used for this shape. A connection site is specified by
    /// defining a point within the shape bounding box that can have a cxnSp element attached to it. These connection
    /// sites are specified using the shape coordinate system that is specified within the ext transform element.
    pub connection_site_list: Option<Vec<ConnectionSite>>,

    /// This element specifies the rectangular bounding box for text within a custGeom shape. The default for this
    /// rectangle is the bounding box for the shape. This can be modified using this elements four attributes to inset or
    /// extend the text bounding box.
    ///
    /// # Note
    ///
    /// Text specified to reside within this shape text rectangle can flow outside this bounding box. Depending on
    /// the autofit options within the txBody element the text might not entirely reside within this shape text rectangle.
    pub rect: Option<Box<GeomRect>>,

    /// This element specifies the entire path that is to make up a single geometric shape. The path_list can consist of
    /// many individual paths within it.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:custGeom>
    ///   <a:pathLst>
    ///     <a:path w="2824222" h="590309">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="428263"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="1620455" y="590309"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="2824222" y="173620"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1562582" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    ///
    /// In the above example there is specified a four sided geometric shape that has all straight sides. While we only
    /// see three lines being drawn via the lnTo element there are actually four sides because the last point of
    /// (x=1562585, y=0) is connected to the first point in the creation path via a lnTo element.
    ///
    /// # Note
    ///
    /// A geometry with multiple paths within it should be treated visually as if each path were a distinct shape.
    /// That is each creation path has its first point and last point joined to form a closed shape. However, the
    /// generating application should then connect the last point to the first point of the new shape. If a close element
    /// is encountered at the end of the previous creation path then this joining line should not be rendered by the
    /// generating application. The rendering should resume with the first line or curve on the new creation path.
    pub path_list: Vec<Path2D>,
}

impl CustomGeometry2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "avLst" => instance.adjust_value_list = Some(GeomGuideList::from_xml_element(child_node)?),
                    "gdLst" => instance.guide_list = Some(GeomGuideList::from_xml_element(child_node)?),
                    "ahLst" => {
                        instance.adjust_handle_list = Some(
                            child_node
                                .child_nodes
                                .iter()
                                .filter_map(AdjustHandle::try_from_xml_element)
                                .collect::<Result<Vec<_>>>()?,
                        )
                    }
                    "cxnLst" => {
                        instance.connection_site_list = Some(
                            child_node
                                .child_nodes
                                .iter()
                                .filter(|cxn_node| cxn_node.local_name() == "cxn")
                                .map(ConnectionSite::from_xml_element)
                                .collect::<Result<Vec<_>>>()?,
                        )
                    }
                    "rect" => instance.rect = Some(Box::new(GeomRect::from_xml_element(child_node)?)),
                    "pathLst" => {
                        instance.path_list = child_node
                            .child_nodes
                            .iter()
                            .filter(|path_node| path_node.local_name() == "path")
                            .map(Path2D::from_xml_element)
                            .collect::<Result<Vec<_>>>()?
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PresetGeometry2D {
    /// Specifies the preset geometry that is used for this shape. This preset can have any of the
    /// values in the enumerated list for ShapeType. This attribute is required in order for a
    /// preset geometry to be rendered.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:nvSpPr>
    ///     <p:cNvPr id="4" name="Sun 3"/>
    ///     <p:cNvSpPr/>
    ///     <p:nvPr/>
    ///   </p:nvSpPr>
    ///   <p:spPr>
    ///     <a:xfrm>
    ///       <a:off x="1981200" y="533400"/>
    ///       <a:ext cx="1143000" cy="1066800"/>
    ///     </a:xfrm>
    ///     <a:prstGeom prst="sun">
    ///     </a:prstGeom>
    ///   </p:spPr>
    /// </p:sp>
    /// ```
    ///
    /// In the above example a preset geometry has been used to define a shape. The shape
    /// utilized here is the sun shape.
    pub preset: ShapeType,
    pub adjust_value_list: Option<GeomGuideList>,
}

impl PresetGeometry2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset = xml_node
            .attributes
            .get("prst")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?
            .parse()?;

        let adjust_value_list = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "avLst")
            .map(GeomGuideList::from_xml_element)
            .transpose()?;

        Ok(Self {
            preset,
            adjust_value_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Geometry {
    /// This element specifies the existence of a custom geometric shape. This shape consists of a series of lines and
    /// curves described within a creation path. In addition to this there can also be adjust values, guides, adjust
    /// handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
    ///
    /// # Xml example
    ///
    /// Consider the scenario when a preset geometry does not accurately depict what must be displayed in
    /// the document. For this a custom geometry can be used to define most any 2-dimensional geometric shape.
    ///
    /// ```xml
    /// <a:custGeom>
    ///   <a:avLst/>
    ///   <a:gdLst/>
    ///   <a:ahLst/>
    ///   <a:cxnLst/>
    ///   <a:rect l="0" t="0" r="0" b="0"/>
    ///   <a:pathLst>
    ///     <a:path w="2650602" h="1261641">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="1261641"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="2650602" y="1261641"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="1226916" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    Custom(Box<CustomGeometry2D>),

    /// This element specifies when a preset geometric shape should be used instead of a custom geometric shape. The
    /// generating application should be able to render all preset geometries enumerated in the ShapeType enum.
    ///
    /// # Xml example
    ///
    /// Consider the scenario when a user does not wish to specify all the lines and curves that make up the
    /// desired shape but instead chooses to use a preset geometry. The following DrawingML would specify such a
    /// case.
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:nvSpPr>
    ///     <p:cNvPr id="4" name="My Preset Shape"/>
    ///     <p:cNvSpPr/>
    ///     <p:nvPr/>
    ///   </p:nvSpPr>
    ///   <p:spPr>
    ///     <a:xfrm>
    ///       <a:off x="1981200" y="533400"/>
    ///       <a:ext cx="1143000" cy="1066800"/>
    ///     </a:xfrm>
    ///     <a:prstGeom prst="heart">
    ///     </a:prstGeom>
    ///   </p:spPr>
    /// </p:sp>
    /// ```
    Preset(Box<PresetGeometry2D>),
}

impl XsdType for Geometry {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "custGeom" => Ok(Geometry::Custom(Box::new(CustomGeometry2D::from_xml_element(
                xml_node,
            )?))),
            "prstGeom" => Ok(Geometry::Preset(Box::new(PresetGeometry2D::from_xml_element(
                xml_node,
            )?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_Geometry").into()),
        }
    }
}

impl XsdChoice for Geometry {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "custGeom" | "prstGeom" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PresetTextShape {
    /// Specifies the preset geometry that is used for a shape warp on a piece of text. This preset
    /// can have any of the values in the enumerated list for TextShapeType. This attribute
    /// is required in order for a text warp to be rendered.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr wrap="none" rtlCol="0">
    ///       <a:prstTxWarp prst="textInflate">
    ///         </a:prstTxWarp>
    ///       <a:spAutoFit/>
    ///     </a:bodyPr>
    ///     <a:lstStyle/>
    ///     <a:p>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:p>
    ///   </p:txBody>
    /// </p:sp>
    /// ```
    ///
    /// In the above example a preset text shape geometry has been used to define the warping
    /// shape. The shape utilized here is the sun shape.
    pub preset: TextShapeType,

    /// The list of adjust values used to represent this preset text shape.
    pub adjust_value_list: Option<GeomGuideList>,
}

impl PresetTextShape {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset = xml_node
            .attributes
            .get("prst")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?
            .parse()?;

        let adjust_value_list = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "avLst")
            .map(GeomGuideList::from_xml_element)
            .transpose()?;

        Ok(Self {
            preset,
            adjust_value_list,
        })
    }
}

/// This element specifies the existence of a connection site on a custom shape. A connection site allows a cxnSp to
/// be attached to this shape. This connection is maintained when the shape is repositioned within the document. It
/// should be noted that this connection is placed within the shape bounding box using the transform coordinate
/// system which is also called the shape coordinate system, as it encompasses the entire shape. The width and
/// height for this coordinate system are specified within the ext transform element.
///
/// # Note
///
/// The transform coordinate system is different from a path coordinate system as it is per shape instead of
/// per path within the shape.
///
/// # Xml example
///
/// Consider the following custom geometry that has two connection sites specified. One connection is
/// located at the bottom left of the shape and the other at the bottom right. The following DrawingML would
/// describe such a custom geometry.
///
/// ```xml
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst/>
///   <a:ahLst/>
///   <a:cxnLst>
///     <a:cxn ang="0">
///       <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///       <a:pos x="1705233" y="679622"/>
///     </a:cxn>
///   </a:cxnLst>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="2" h="2">
///       <a:moveTo>
///         <a:pt x="0" y="2"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2" y="2"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct ConnectionSite {
    /// Specifies the incoming connector angle. This angle is the angle around the connection
    /// site that an incoming connector tries to be routed to. This allows connectors to know
    /// where the shape is in relation to the connection site and route connectors so as to avoid
    /// any overlap with the shape.
    pub angle: AdjAngle,
    pub position: AdjPoint2D,
}

impl ConnectionSite {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let angle = xml_node
            .attributes
            .get("ang")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "ang"))?
            .parse()?;

        let position = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pos")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "pos")))
            .and_then(AdjPoint2D::from_xml_element)?;

        Ok(Self { angle, position })
    }
}
