use crate::{
    error::MissingAttributeError,
    xml::{parse_xml_bool, XmlNode}
};
use super::simpletypes::{Angle, Coordinate, PositiveCoordinate};

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Debug, Copy, Clone, Eq, PartialEq)]
pub struct Point2D {
    /// Specifies a coordinate on the x-axis. The origin point for this coordinate shall be specified
    /// by the parent XML element.
    pub x: Coordinate,

    /// Specifies a coordinate on the x-axis. The origin point for this coordinate shall be specified
    /// by the parent XML element.
    pub y: Coordinate,
}

impl Point2D {
    pub fn new(x: Coordinate, y: Coordinate) -> Self {
        Self { x, y }
    }

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

#[derive(Debug, Copy, Clone, Eq, PartialEq)]
pub struct PositiveSize2D {
    /// Specifies the length of the extents rectangle in EMUs. This rectangle shall dictate the size
    /// of the object as displayed (the result of any scaling to the original object).
    pub width: PositiveCoordinate,
    /// Specifies the width of the extents rectangle in EMUs. This rectangle shall dictate the size
    /// of the object as displayed (the result of any scaling to the original object).
    pub height: PositiveCoordinate,
}

impl PositiveSize2D {
    pub fn new(width: PositiveCoordinate, height: PositiveCoordinate) -> Self {
        Self { width, height }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_width = None;
        let mut opt_height = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "cx" => opt_width = Some(value.parse::<PositiveCoordinate>()?),
                "cy" => opt_height = Some(value.parse::<PositiveCoordinate>()?),
                _ => (),
            }
        }

        let width = opt_width.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cx"))?;
        let height = opt_height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cy"))?;

        Ok(Self { width, height })
    }
}

#[derive(Default, Debug, Copy, Clone, PartialEq)]
pub struct Transform2D {
    /// Specifies the rotation of the Graphic Frame. The units for which this attribute is specified
    /// in reside within the simple type definition referenced below.
    pub rotate_angle: Option<Angle>,

    /// Specifies a horizontal flip. When true, this attribute defines that the shape is flipped
    /// horizontally about the center of its bounding box.
    ///
    /// Defaults to false
    pub flip_horizontal: Option<bool>,

    /// Specifies a vertical flip. When true, this attribute defines that the group is flipped
    /// vertically about the center of its bounding box.
    pub flip_vertical: Option<bool>,

    /// This element specifies the location of the bounding box of an object. Effects on an object are not included in this
    /// bounding box.
    pub offset: Option<Point2D>,

    /// This element specifies the size of the bounding box enclosing the referenced object.
    pub extents: Option<PositiveSize2D>,
}

impl Transform2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (key, value)| {
                match key.as_str() {
                    "rot" => instance.rotate_angle = Some(value.parse()?),
                    "flipH" => instance.flip_horizontal = Some(parse_xml_bool(value)?),
                    "flipV" => instance.flip_vertical = Some(parse_xml_bool(value)?),
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
                            "off" => instance.offset = Some(Point2D::from_xml_element(child_node)?),
                            "ext" => instance.extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Default, Debug, Copy, Clone, PartialEq)]
pub struct GroupTransform2D {
    /// Rotation. Specifies the clockwise rotation of a group in 1/64000 of a degree.
    ///
    /// Defaults to 0
    pub rotate_angle: Option<Angle>,

    /// Horizontal flip. When true, this attribute defines that the group is flipped horizontally
    /// about the center of its bounding box.
    ///
    /// Defaults to false
    pub flip_horizontal: Option<bool>,

    /// Vertical flip. When true, this attribute defines that the group is flipped vertically about
    /// the center of its bounding box.
    ///
    /// Defaults to false
    pub flip_vertical: Option<bool>,

    /// This element specifies the location of the bounding box of an object. Effects on an object are not included in this
    /// bounding box.
    pub offset: Option<Point2D>,

    /// This element specifies the size of the bounding box enclosing the referenced object.
    pub extents: Option<PositiveSize2D>,

    /// This element specifies the location of the child extents rectangle and is used for calculations of grouping, scaling,
    /// and rotation behavior of shapes placed within a group.
    pub child_offset: Option<Point2D>,

    /// This element specifies the size dimensions of the child extents rectangle and is used for calculations of grouping,
    /// scaling, and rotation behavior of shapes placed within a group.
    pub child_extents: Option<PositiveSize2D>,
}

impl GroupTransform2D {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "rot" => instance.rotate_angle = Some(value.parse()?),
                    "flipH" => instance.flip_horizontal = Some(parse_xml_bool(value)?),
                    "flipV" => instance.flip_vertical = Some(parse_xml_bool(value)?),
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
                            "off" => instance.offset = Some(Point2D::from_xml_element(child_node)?),
                            "ext" => instance.extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                            "chOff" => instance.child_offset = Some(Point2D::from_xml_element(child_node)?),
                            "chExt" => instance.child_extents = Some(PositiveSize2D::from_xml_element(child_node)?),
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}
