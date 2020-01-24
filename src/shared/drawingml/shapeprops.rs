use super::{
    colors::Color,
    simpletypes::{
        BlendMode, BlipCompression, Coordinate, EffectContainerType, FixedAngle, FixedPercentage, LineEndLength,
        LineEndType, LineEndWidth, PathShadeType, Percentage, PositiveCoordinate, PositiveFixedAngle,
        PositiveFixedPercentage, PositivePercentage, PresetLineDashVal, PresetPatternVal, PresetShadowVal,
        RectAlignment, TileFlipMode,
    },
};
use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    shared::relationship::RelationshipId,
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use log::trace;
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Default, Debug, Clone, PartialEq)]
pub struct RelativeRect {
    /// Specifies the left edge of the rectangle.
    pub left: Option<Percentage>,

    /// Specifies the top edge of the rectangle.
    pub top: Option<Percentage>,

    /// Specifies the right edge of the rectangle.
    pub right: Option<Percentage>,

    /// Specifies the bottom edge of the rectangle.
    pub bottom: Option<Percentage>,
}

impl RelativeRect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<RelativeRect> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "l" => instance.left = Some(value.parse()?),
                    "t" => instance.top = Some(value.parse()?),
                    "r" => instance.right = Some(value.parse()?),
                    "b" => instance.bottom = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element represents an Alpha Bi-Level Effect.
///
/// Alpha (Opacity) values less than the threshold are changed to 0 (fully transparent) and alpha values greater than
/// or equal to the threshold are changed to 100% (fully opaque).
#[derive(Debug, Clone, PartialEq)]
pub struct AlphaBiLevelEffect {
    // Specifies the threshold value for the alpha bi-level effect.
    pub threshold: PositiveFixedPercentage,
}

impl AlphaBiLevelEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaBiLevelEffect> {
        let threshold = xml_node
            .attributes
            .get("thresh")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "thresh"))?
            .parse()?;

        Ok(Self { threshold })
    }
}

/// This element represents an alpha inverse effect.
///
/// Alpha (opacity) values are inverted by subtracting from 100%.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct AlphaInverseEffect {
    pub color: Option<Color>,
}

impl AlphaInverseEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaInverseEffect> {
        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?;

        Ok(Self { color })
    }
}

/// This element represents an alpha modulate effect.
///
/// Effect alpha (opacity) values are multiplied by a fixed percentage. The effect container specifies an effect
/// containing alpha values to modulate.
#[derive(Debug, Clone, PartialEq)]
pub struct AlphaModulateEffect {
    pub container: EffectContainer,
}

impl AlphaModulateEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<AlphaModulateEffect> {
        let container = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cont")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "container")))
            .and_then(EffectContainer::from_xml_element)?;

        Ok(Self { container })
    }
}

/// This element represents an alpha modulate fixed effect.
///
/// Effect alpha (opacity) values are multiplied by a fixed percentage.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct AlphaModulateFixedEffect {
    /// Specifies the percentage amount to scale the alpha.
    ///
    /// Defaults to 100000
    pub amount: Option<PositivePercentage>,
}

impl AlphaModulateFixedEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let amount = xml_node.attributes.get("amt").map(|value| value.parse()).transpose()?;

        Ok(Self { amount })
    }
}

/// This element specifies an alpha outset/inset effect.
///
/// This is equivalent to an alpha ceiling, followed by alpha blur, followed by either an alpha ceiling (positive radius)
/// or alpha floor (negative radius).
#[derive(Default, Debug, Clone, PartialEq)]
pub struct AlphaOutsetEffect {
    /// Specifies the radius of outset/inset.
    pub radius: Option<Coordinate>,
}

impl AlphaOutsetEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius = xml_node.attributes.get("rad").map(|value| value.parse()).transpose()?;

        Ok(Self { radius })
    }
}

/// This element specifies an alpha replace effect.
///
/// Effect alpha (opacity) values are replaced by a fixed alpha.
#[derive(Debug, Clone, PartialEq)]
pub struct AlphaReplaceEffect {
    /// Specifies the new opacity value.
    pub alpha: PositiveFixedPercentage,
}

impl AlphaReplaceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let alpha = xml_node
            .attributes
            .get("a")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "a"))?
            .parse()?;

        Ok(Self { alpha })
    }
}

/// This element specifies a bi-level (black/white) effect. Input colors whose luminance is less than the specified
/// threshold value are changed to black. Input colors whose luminance are greater than or equal the specified
/// value are set to white. The alpha effect values are unaffected by this effect.
#[derive(Debug, Clone, PartialEq)]
pub struct BiLevelEffect {
    /// Specifies the luminance threshold for the Bi-Level effect. Values greater than or equal to
    /// the threshold are set to white. Values lesser than the threshold are set to black.
    pub threshold: PositiveFixedPercentage,
}

impl BiLevelEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let threshold = xml_node
            .attributes
            .get("thresh")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "thresh"))?
            .parse()?;

        Ok(Self { threshold })
    }
}

/// This element specifies a blend of several effects. The container specifies the raw effects to blend while the blend
/// mode specifies how the effects are to be blended.
#[derive(Debug, Clone, PartialEq)]
pub struct BlendEffect {
    /// Specifies how to blend the two effects.
    pub blend: BlendMode,
    pub container: EffectContainer,
}

impl BlendEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let blend = xml_node
            .attributes
            .get("blend")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "blend"))?
            .parse()?;

        let container = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "cont")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "cont")))
            .and_then(EffectContainer::from_xml_element)?;

        Ok(Self { blend, container })
    }
}

/// This element specifies a blur effect that is applied to the entire shape, including its fill. All color channels,
/// including alpha, are affected.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct BlurEffect {
    /// Specifies the radius of blur.
    ///
    /// Defaults to 0
    pub radius: Option<PositiveCoordinate>,

    /// Specifies whether the bounds of the object should be grown as a result of the blurring.
    /// True indicates the bounds are grown while false indicates that they are not.
    ///
    /// With grow set to false, the blur effect does not extend beyond the original bounds of the
    /// object
    ///
    /// With grow set to true, the blur effect can extend beyond the original bounds of the
    /// object
    ///
    /// Defaults to true
    pub grow: Option<bool>,
}

impl BlurEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "rad" => instance.radius = Some(value.parse()?),
                    "grow" => instance.grow = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies a Color Change Effect. Instances of clrFrom are replaced with instances of clrTo.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorChangeEffect {
    /// Specifies whether alpha values are considered for the effect. Effect alpha values are
    /// considered if use_alpha is true, else they are ignored.
    ///
    /// Defaults to true
    pub use_alpha: Option<bool>,

    /// This element specifies a solid color replacement value. All effect colors are changed to a fixed color. Alpha values
    /// are unaffected.
    pub color_from: Color,

    /// This element specifies the color which replaces the clrFrom in a clrChange effect. This is the "target" or "to"
    /// color in the color change effect.
    pub color_to: Color,
}

impl ColorChangeEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let use_alpha = xml_node.attributes.get("useA").map(|value| value.parse()).transpose()?;

        let mut color_from = None;
        let mut color_to = None;
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrFrom" => {
                    color_from = child_node
                        .child_nodes
                        .iter()
                        .find_map(Color::try_from_xml_element)
                        .transpose()?
                }
                "clrTo" => {
                    color_to = child_node
                        .child_nodes
                        .iter()
                        .find_map(Color::try_from_xml_element)
                        .transpose()?
                }
                _ => (),
            }
        }

        let color_from = color_from.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrFrom"))?;
        let color_to = color_to.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrTo"))?;

        Ok(Self {
            use_alpha,
            color_from,
            color_to,
        })
    }
}

/// This element specifies a solid color replacement value. All effect colors are changed to a fixed color. Alpha values
/// are unaffected.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorReplaceEffect {
    pub color: Color,
}

impl ColorReplaceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;

        Ok(Self { color })
    }
}

/// This element specifies a luminance effect. Brightness linearly shifts all colors closer to white or black.
/// Contrast scales all colors to be either closer or further apart.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct LuminanceEffect {
    /// Specifies the percent to change the brightness.
    pub brightness: Option<FixedPercentage>,

    /// Specifies the percent to change the contrast.
    pub contrast: Option<FixedPercentage>,
}

impl LuminanceEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "bright" => instance.brightness = Some(value.parse()?),
                    "contrast" => instance.contrast = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies a duotone effect.
///
/// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
#[derive(Debug, Clone, PartialEq)]
pub struct DuotoneEffect {
    pub colors: [Color; 2],
}

impl DuotoneEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut iterator = xml_node.child_nodes.iter().filter_map(Color::try_from_xml_element);

        let color1 = iterator
            .next()
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;

        let color2 = iterator
            .next()
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_Color"))?;

        // TODO(dam4rus): Check if node contains more than 2 color?
        Ok(Self {
            colors: [color1, color2],
        })
    }
}

/// This element specifies a fill which is one of blipFill, gradFill, grpFill, noFill, pattFill or solidFill.
#[derive(Debug, Clone, PartialEq)]
pub struct FillEffect {
    pub fill_properties: FillProperties,
}

impl FillEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fill_properties = xml_node
            .child_nodes
            .iter()
            .find_map(FillProperties::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;

        Ok(Self { fill_properties })
    }
}

/// This element specifies a fill overlay effect. A fill overlay can be used to specify an additional fill for an object and
/// blend the two fills together.
#[derive(Debug, Clone, PartialEq)]
pub struct FillOverlayEffect {
    /// Specifies how to blend the fill with the base effect.
    pub blend_mode: BlendMode,
    pub fill: FillProperties,
}

impl FillOverlayEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let blend_mode = xml_node
            .attributes
            .get("blend")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "blend"))?
            .parse()?;

        let fill = xml_node
            .child_nodes
            .iter()
            .find_map(FillProperties::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;

        Ok(Self { blend_mode, fill })
    }
}

/// This element specifies a glow effect, in which a color blurred outline is added outside the edges of the object.
#[derive(Debug, Clone, PartialEq)]
pub struct GlowEffect {
    /// Specifies the radius of the glow.
    ///
    /// Defaults to 0
    pub radius: Option<PositiveCoordinate>,
    pub color: Color,
}

impl GlowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius = xml_node.attributes.get("rad").map(|value| value.parse()).transpose()?;

        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;

        Ok(Self { radius, color })
    }
}

/// This element specifies a hue/saturation/luminance effect. The hue, saturation, and luminance can each be
/// adjusted relative to its current value.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct HslEffect {
    /// Specifies the number of degrees by which the hue is adjusted.
    ///
    /// Defaults to 0
    pub hue: Option<PositiveFixedAngle>,

    /// Specifies the percentage by which the saturation is adjusted.
    ///
    /// Defaults to 0
    pub saturation: Option<FixedPercentage>,

    /// Specifies the percentage by which the luminance is adjusted.
    ///
    /// Defaults to 0
    pub luminance: Option<FixedPercentage>,
}

impl HslEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "hue" => instance.hue = Some(value.parse()?),
                    "sat" => instance.saturation = Some(value.parse()?),
                    "lum" => instance.luminance = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies an inner shadow effect. A shadow is applied within the edges of the object according to
/// the parameters given by the attributes.
#[derive(Debug, Clone, PartialEq)]
pub struct InnerShadowEffect {
    /// Specifies the blur radius.
    ///
    /// Defaults to 0
    pub blur_radius: Option<PositiveCoordinate>,

    /// Specifies how far to offset the shadow.
    ///
    /// Defaults to 0
    pub distance: Option<PositiveCoordinate>,

    /// Specifies the direction to offset the shadow.
    ///
    /// Defaults to 0
    pub direction: Option<PositiveFixedAngle>,
    pub color: Color,
}

impl InnerShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;

        let mut blur_radius = None;
        let mut distance = None;
        let mut direction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "blurRad" => blur_radius = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            blur_radius,
            distance,
            direction,
            color,
        })
    }
}

/// This element specifies an Outer Shadow Effect.
#[derive(Debug, Clone, PartialEq)]
pub struct OuterShadowEffect {
    /// Specifies the blur radius of the shadow.
    ///
    /// Defaults to 0
    pub blur_radius: Option<PositiveCoordinate>,

    /// Specifies the how far to offset the shadow.
    ///
    /// Defaults to 0
    pub distance: Option<PositiveCoordinate>,

    /// Specifies the direction to offset the shadow.
    ///
    /// Defaults to 0
    pub direction: Option<PositiveFixedAngle>,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip.
    ///
    /// Defaults to 100_000
    pub scale_x: Option<Percentage>,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    ///
    /// Defaults to 100_000
    pub scale_y: Option<Percentage>,

    /// Specifies the horizontal skew angle.
    ///
    /// Defaults to 0
    pub skew_x: Option<FixedAngle>,

    /// Specifies the vertical skew angle.
    ///
    /// Defaults to 0
    pub skew_y: Option<FixedAngle>,

    /// Specifies shadow alignment; alignment happens first, effectively setting the origin for
    /// scale, skew, and offset.
    ///
    /// Defaults to RectAlignment::Bottom
    pub alignment: Option<RectAlignment>,

    /// Specifies whether the shadow rotates with the shape if the shape is rotated.
    ///
    /// Defaults to true
    pub rotate_with_shape: Option<bool>,
    pub color: Color,
}

impl OuterShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;

        let mut blur_radius = None;
        let mut distance = None;
        let mut direction = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut skew_x = None;
        let mut skew_y = None;
        let mut alignment = None;
        let mut rotate_with_shape = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "blurRad" => blur_radius = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                "sx" => scale_x = Some(value.parse()?),
                "sy" => scale_y = Some(value.parse()?),
                "kx" => skew_x = Some(value.parse()?),
                "ky" => skew_y = Some(value.parse()?),
                "algn" => alignment = Some(value.parse()?),
                "rotWithShape" => rotate_with_shape = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(Self {
            blur_radius,
            distance,
            direction,
            scale_x,
            scale_y,
            skew_x,
            skew_y,
            alignment,
            rotate_with_shape,
            color,
        })
    }
}

/// This element specifies that a preset shadow is to be used. Each preset shadow is equivalent to a specific outer
/// shadow effect. For each preset shadow, the color element, direction attribute, and distance attribute represent
/// the color, direction, and distance parameters of the corresponding outer shadow. Additionally, the
/// rotateWithShape attribute of corresponding outer shadow is always false. Other non-default parameters of
/// the outer shadow are dependent on the prst attribute.
#[derive(Debug, Clone, PartialEq)]
pub struct PresetShadowEffect {
    /// Specifies which preset shadow to use.
    pub preset: PresetShadowVal,

    /// Specifies how far to offset the shadow.
    ///
    /// Defaults to 0
    pub distance: Option<PositiveCoordinate>,

    /// Specifies the direction to offset the shadow.
    ///
    /// Defaults to 0
    pub direction: Option<PositiveFixedAngle>,
    pub color: Color,
}

impl PresetShadowEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice"))?;

        let mut preset = None;
        let mut distance = None;
        let mut direction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "prst" => preset = Some(value.parse()?),
                "dist" => distance = Some(value.parse()?),
                "dir" => direction = Some(value.parse()?),
                _ => (),
            }
        }

        let preset = preset.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "prst"))?;

        Ok(Self {
            preset,
            distance,
            direction,
            color,
        })
    }
}

/// This element specifies a reflection effect.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct ReflectionEffect {
    /// Specifies the blur radius.
    ///
    /// Defaults to 0
    pub blur_radius: Option<PositiveCoordinate>,

    /// Starting reflection opacity.
    ///
    /// Defaults to 100_000
    pub start_opacity: Option<PositiveFixedPercentage>,

    /// Specifies the start position (along the alpha gradient ramp) of the start alpha value.
    ///
    /// Defaults to 0
    pub start_position: Option<PositiveFixedPercentage>,

    /// Specifies the ending reflection opacity.
    ///
    /// Defaults to 0
    pub end_opacity: Option<PositiveFixedPercentage>,

    /// Specifies the end position (along the alpha gradient ramp) of the end alpha value.
    ///
    /// Defaults to 100_000
    pub end_position: Option<PositiveFixedPercentage>,

    /// Specifies how far to distance the shadow.
    ///
    /// Defaults to 0
    pub distance: Option<PositiveCoordinate>,

    /// Specifies the direction of the alpha gradient ramp relative to the shape itself.
    ///
    /// Defaults to 0
    pub direction: Option<PositiveFixedAngle>,

    /// Specifies the direction to offset the reflection.
    ///
    /// Defaults to 5_400_000
    pub fade_direction: Option<PositiveFixedAngle>,

    /// Specifies the horizontal scaling factor.
    ///
    /// Defaults to 100_000
    pub scale_x: Option<Percentage>,

    /// Specifies the vertical scaling factor.
    ///
    /// Defaults to 100_000
    pub scale_y: Option<Percentage>,

    /// Specifies the horizontal skew angle.
    ///
    /// Defaults to 0
    pub skew_x: Option<FixedAngle>,

    /// Specifies the vertical skew angle.
    ///
    /// Defaults to 0
    pub skew_y: Option<FixedAngle>,

    /// Specifies shadow alignment.
    ///
    /// Defaults to RectAlignment::Bottom
    pub alignment: Option<RectAlignment>,

    /// Specifies if the reflection rotates with the shape.
    ///
    /// Defaults to true
    pub rotate_with_shape: Option<bool>,
}

impl ReflectionEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "blurRad" => instance.blur_radius = Some(value.parse()?),
                    "stA" => instance.start_opacity = Some(value.parse()?),
                    "stPos" => instance.start_position = Some(value.parse()?),
                    "endA" => instance.end_opacity = Some(value.parse()?),
                    "endPos" => instance.end_position = Some(value.parse()?),
                    "dist" => instance.distance = Some(value.parse()?),
                    "dir" => instance.direction = Some(value.parse()?),
                    "fadeDir" => instance.fade_direction = Some(value.parse()?),
                    "sx" => instance.scale_x = Some(value.parse()?),
                    "sy" => instance.scale_y = Some(value.parse()?),
                    "kx" => instance.skew_x = Some(value.parse()?),
                    "ky" => instance.skew_y = Some(value.parse()?),
                    "algn" => instance.alignment = Some(value.parse()?),
                    "rotWithShape" => instance.rotate_with_shape = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies a relative offset effect. Sets up a new origin by offsetting relative to the size of the
/// previous effect.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct RelativeOffsetEffect {
    /// Specifies the X offset.
    ///
    /// Defaults to 0
    pub translate_x: Option<Percentage>,

    /// Specifies the Y offset.
    ///
    /// Defaults to 0
    pub translate_y: Option<Percentage>,
}

impl RelativeOffsetEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "tx" => instance.translate_x = Some(value.parse()?),
                    "ty" => instance.translate_y = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies a soft edge effect. The edges of the shape are blurred, while the fill is not affected.
#[derive(Debug, Clone, PartialEq)]
pub struct SoftEdgesEffect {
    /// Specifies the radius of blur to apply to the edges.
    pub radius: PositiveCoordinate,
}

impl SoftEdgesEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let radius = xml_node
            .attributes
            .get("rad")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "rad"))?
            .parse()?;

        Ok(Self { radius })
    }
}

/// This element specifies a tint effect. Shifts effect color values towards/away from hue by the specified amount.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TintEffect {
    /// Specifies the hue towards which to tint.
    ///
    /// Defaults to 0
    pub hue: Option<PositiveFixedAngle>,

    /// Specifies by how much the color value is shifted.
    ///
    /// Defaults to 0
    pub amount: Option<FixedPercentage>,
}

impl TintEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "hue" => instance.hue = Some(value.parse()?),
                    "amt" => instance.amount = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies a transform effect. The transform is applied to each point in the shape's geometry using
/// the following matrix:
///
/// sx          tan(kx)     tx      x
/// tan(ky)     sy          ty  *   y
/// 0           0           1       1
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TransformEffect {
    /// Specifies a percentage by which to horizontally scale the object.
    ///
    /// Defaults to 100_000
    pub scale_x: Option<Percentage>,

    /// Specifies a percentage by which to vertically scale the object.
    ///
    /// Defaults to 100_000
    pub scale_y: Option<Percentage>,

    /// Specifies an amount by which to shift the object along the x-axis.
    ///
    /// Defaults to 0
    pub translate_x: Option<Coordinate>,

    /// Specifies an amount by which to shift the object along the y-axis.
    ///
    /// Defaults to 0
    pub translate_y: Option<Coordinate>,

    /// Specifies the horizontal skew angle, defined as the angle between the top-left corner and
    /// bottom-left corner of the object's original bounding box. If positive, the bottom edge of
    /// the shape is positioned to the right relative to the top edge.
    ///
    /// Defaults to 0
    pub skew_x: Option<FixedAngle>,

    /// Specifies the vertical skew angle, defined as the angle between the top-left corner and
    /// top-right corner of the object's original bounding box. If positive, the right edge of the
    /// object is positioned lower relative to the left edge.
    ///
    /// Defaults to 0
    pub skew_y: Option<FixedAngle>,
}

impl TransformEffect {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "sx" => instance.scale_x = Some(value.parse()?),
                    "sy" => instance.scale_y = Some(value.parse()?),
                    "kx" => instance.skew_x = Some(value.parse()?),
                    "ky" => instance.skew_y = Some(value.parse()?),
                    "tx" => instance.translate_x = Some(value.parse()?),
                    "ty" => instance.translate_y = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

// TODO: maybe Box ReflectionEffect variant (sizeof==120)
#[derive(Debug, Clone, PartialEq)]
pub enum Effect {
    Container(EffectContainer),

    /// This element specifies a reference to an existing effect container.
    ///
    /// Its value can be the name of an effect container, or one of four
    /// special references:
    /// * fill - refers to the fill effect
    /// * line - refers to the line effect
    /// * fillLine - refers to the combined fill and line effects
    /// * children - refers to the combined effects from logical child shapes or text
    EffectReference(String),
    AlphaBiLevel(AlphaBiLevelEffect),

    /// This element represents an alpha ceiling effect.
    ///
    /// Alpha (opacity) values greater than zero are changed to 100%. In other words, anything partially opaque
    /// becomes fully opaque.
    AlphaCeiling,

    /// This element represents an alpha floor effect.
    ///
    /// Alpha (opacity) values less than 100% are changed to zero. In other words, anything partially transparent
    /// becomes fully transparent.
    AlphaFloor,
    AlphaInverse(AlphaInverseEffect),
    AlphaModulate(AlphaModulateEffect),
    AlphaModulateFixed(AlphaModulateFixedEffect),
    AlphaOutset(AlphaOutsetEffect),
    AlphaReplace(AlphaReplaceEffect),
    BiLevel(BiLevelEffect),
    Blend(BlendEffect),
    Blur(BlurEffect),
    ColorChange(ColorChangeEffect),
    ColorReplace(ColorReplaceEffect),
    Duotone(DuotoneEffect),
    Fill(FillEffect),
    FillOverlay(FillOverlayEffect),
    Glow(GlowEffect),

    /// This element specifies a gray scale effect. Converts all effect color values to a shade of gray, corresponding to
    /// their luminance. Effect alpha (opacity) values are unaffected.
    Grayscale,
    Hsl(HslEffect),
    InnerShadow(InnerShadowEffect),
    Luminance(LuminanceEffect),
    OuterShadow(OuterShadowEffect),
    PresetShadow(PresetShadowEffect),
    Reflection(ReflectionEffect),
    RelativeOffset(RelativeOffsetEffect),
    SoftEdges(SoftEdgesEffect),
    Tint(TintEffect),
    Transform(TransformEffect),
}

impl XsdType for Effect {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "cont" => Ok(Effect::Container(EffectContainer::from_xml_element(xml_node)?)),
            "effect" => {
                let reference = xml_node
                    .attributes
                    .get("ref")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "ref"))?
                    .clone();
                Ok(Effect::EffectReference(reference))
            }
            "alphaBiLevel" => Ok(Effect::AlphaBiLevel(AlphaBiLevelEffect::from_xml_element(xml_node)?)),
            "alphaCeiling" => Ok(Effect::AlphaCeiling),
            "alphaFloor" => Ok(Effect::AlphaFloor),
            "alphaInv" => Ok(Effect::AlphaInverse(AlphaInverseEffect::from_xml_element(xml_node)?)),
            "alphaMod" => Ok(Effect::AlphaModulate(AlphaModulateEffect::from_xml_element(xml_node)?)),
            "alphaModFix" => Ok(Effect::AlphaModulateFixed(AlphaModulateFixedEffect::from_xml_element(
                xml_node,
            )?)),
            "alphaOutset" => Ok(Effect::AlphaOutset(AlphaOutsetEffect::from_xml_element(xml_node)?)),
            "alphaRepl" => Ok(Effect::AlphaReplace(AlphaReplaceEffect::from_xml_element(xml_node)?)),
            "biLevel" => Ok(Effect::BiLevel(BiLevelEffect::from_xml_element(xml_node)?)),
            "blend" => Ok(Effect::Blend(BlendEffect::from_xml_element(xml_node)?)),
            "blur" => Ok(Effect::Blur(BlurEffect::from_xml_element(xml_node)?)),
            "clrChange" => Ok(Effect::ColorChange(ColorChangeEffect::from_xml_element(xml_node)?)),
            "clrRepl" => Ok(Effect::ColorReplace(ColorReplaceEffect::from_xml_element(xml_node)?)),
            "duotone" => Ok(Effect::Duotone(DuotoneEffect::from_xml_element(xml_node)?)),
            "fill" => Ok(Effect::Fill(FillEffect::from_xml_element(xml_node)?)),
            "fillOverlay" => Ok(Effect::FillOverlay(FillOverlayEffect::from_xml_element(xml_node)?)),
            "glow" => Ok(Effect::Glow(GlowEffect::from_xml_element(xml_node)?)),
            "grayscl" => Ok(Effect::Grayscale),
            "hsl" => Ok(Effect::Hsl(HslEffect::from_xml_element(xml_node)?)),
            "innerShdw" => Ok(Effect::InnerShadow(InnerShadowEffect::from_xml_element(xml_node)?)),
            "lum" => Ok(Effect::Luminance(LuminanceEffect::from_xml_element(xml_node)?)),
            "outerShdw" => Ok(Effect::OuterShadow(OuterShadowEffect::from_xml_element(xml_node)?)),
            "prstShdw" => Ok(Effect::PresetShadow(PresetShadowEffect::from_xml_element(xml_node)?)),
            "reflection" => Ok(Effect::Reflection(ReflectionEffect::from_xml_element(xml_node)?)),
            "relOff" => Ok(Effect::RelativeOffset(RelativeOffsetEffect::from_xml_element(
                xml_node,
            )?)),
            "softEdge" => Ok(Effect::SoftEdges(SoftEdgesEffect::from_xml_element(xml_node)?)),
            "tint" => Ok(Effect::Tint(TintEffect::from_xml_element(xml_node)?)),
            "xfrm" => Ok(Effect::Transform(TransformEffect::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_Effect"))),
        }
    }
}

impl XsdChoice for Effect {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "cont" | "effect" | "alphaBiLevel" | "alphaCeiling" | "alphaFloor" | "alphaInv" | "alphaMod"
            | "alphaModFix" | "alphaOutset" | "alphaRepl" | "biLevel" | "blend" | "blur" | "clrChange" | "clrRepl"
            | "duotone" | "fill" | "fillOverlay" | "glow" | "grayscl" | "hsl" | "innerShdw" | "lum" | "outerShdw"
            | "prstShdw" | "reflection" | "relOff" | "softEdge" | "tint" | "xfrm" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum BlipEffect {
    AlphaBiLevel(AlphaBiLevelEffect),

    /// This element represents an alpha ceiling effect.
    ///
    /// Alpha (opacity) values greater than zero are changed to 100%. In other words, anything partially opaque
    /// becomes fully opaque.
    AlphaCeiling,

    /// This element represents an alpha floor effect.
    ///
    /// Alpha (opacity) values less than 100% are changed to zero. In other words, anything partially transparent
    /// becomes fully transparent.
    AlphaFloor,
    AlphaInverse(AlphaInverseEffect),
    AlphaModulate(AlphaModulateEffect),
    AlphaModulateFixed(AlphaModulateFixedEffect),
    AlphaReplace(AlphaReplaceEffect),
    BiLevel(BiLevelEffect),
    Blur(BlurEffect),
    ColorChange(ColorChangeEffect),
    ColorReplace(ColorReplaceEffect),
    Duotone(DuotoneEffect),
    FillOverlay(FillOverlayEffect),

    /// This element specifies a gray scale effect. Converts all effect color values to a shade of gray, corresponding to
    /// their luminance. Effect alpha (opacity) values are unaffected.
    Grayscale,
    Hsl(HslEffect),
    Luminance(LuminanceEffect),
    Tint(TintEffect),
}

impl XsdType for BlipEffect {
    fn from_xml_element(xml_node: &XmlNode) -> Result<BlipEffect> {
        match xml_node.local_name() {
            "alphaBiLevel" => Ok(BlipEffect::AlphaBiLevel(AlphaBiLevelEffect::from_xml_element(
                xml_node,
            )?)),
            "alphaCeiling" => Ok(BlipEffect::AlphaCeiling),
            "alphaFloor" => Ok(BlipEffect::AlphaFloor),
            "alphaInv" => Ok(BlipEffect::AlphaInverse(AlphaInverseEffect::from_xml_element(
                xml_node,
            )?)),
            "alphaMod" => Ok(BlipEffect::AlphaModulate(AlphaModulateEffect::from_xml_element(
                xml_node,
            )?)),
            "alphaModFixed" => Ok(BlipEffect::AlphaModulateFixed(
                AlphaModulateFixedEffect::from_xml_element(xml_node)?,
            )),
            "alphaRepl" => Ok(BlipEffect::AlphaReplace(AlphaReplaceEffect::from_xml_element(
                xml_node,
            )?)),
            "biLevel" => Ok(BlipEffect::BiLevel(BiLevelEffect::from_xml_element(xml_node)?)),
            "blur" => Ok(BlipEffect::Blur(BlurEffect::from_xml_element(xml_node)?)),
            "clrChange" => Ok(BlipEffect::ColorChange(ColorChangeEffect::from_xml_element(xml_node)?)),
            "clrRepl" => Ok(BlipEffect::ColorReplace(ColorReplaceEffect::from_xml_element(
                xml_node,
            )?)),
            "duotone" => Ok(BlipEffect::Duotone(DuotoneEffect::from_xml_element(xml_node)?)),
            "fillOverlay" => Ok(BlipEffect::FillOverlay(FillOverlayEffect::from_xml_element(xml_node)?)),
            "grayscl" => Ok(BlipEffect::Grayscale),
            "hsl" => Ok(BlipEffect::Hsl(HslEffect::from_xml_element(xml_node)?)),
            "lum" => Ok(BlipEffect::Luminance(LuminanceEffect::from_xml_element(xml_node)?)),
            "tint" => Ok(BlipEffect::Tint(TintEffect::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_BlipEffect").into()),
        }
    }
}

impl XsdChoice for BlipEffect {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "alphaBiLevel" | "alphaCeiling" | "alphaFloor" | "alphaInv" | "alphaMod" | "alphaModFixed"
            | "alphaRepl" | "biLevel" | "blur" | "clrChange" | "clrRepl" | "duotone" | "fillOverlay" | "grayscl"
            | "hsl" | "lum" | "tint" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum EffectProperties {
    /// This element specifies a list of effects. Effects in an effectLst are applied in the default order by the rendering
    /// engine. The following diagrams illustrate the order in which effects are applied, both for shapes and for group
    /// shapes.
    ///
    /// # Note
    ///
    /// The output of many effects does not include the input shape. For effects that should be applied to the
    /// result of previous effects as well as the original shape, a container is used to group the inputs together.
    ///
    /// # Example
    ///
    /// Outer Shadow is applied both to the original shape and the original shape's glow. The result of blur
    /// contains the original shape, while the result of glow contains only the added glow. Therefore, a container that
    /// groups the blur result with the glow result is used as the input to Outer Shadow.
    EffectList(Box<EffectList>),

    /// This element specifies a list of effects. Effects are applied in the order specified by the container type (sibling or
    /// tree).
    ///
    /// # Note
    ///
    /// An effectDag element can contain multiple effect containers as child elements. Effect containers with
    /// different styles can be combined in an effectDag to define a directed acyclic graph (DAG) that specifies the order
    /// in which all effects are applied.
    EffectContainer(Box<EffectContainer>),
}

impl XsdType for EffectProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "effectLst" => Ok(EffectProperties::EffectList(Box::new(EffectList::from_xml_element(
                xml_node,
            )?))),
            "effectDag" => Ok(EffectProperties::EffectContainer(Box::new(
                EffectContainer::from_xml_element(xml_node)?,
            ))),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "EG_EffectProperties",
            ))),
        }
    }
}

impl XsdChoice for EffectProperties {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "effectLst" | "effectDag" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct EffectList {
    pub blur: Option<BlurEffect>,
    pub fill_overlay: Option<FillOverlayEffect>,
    pub glow: Option<GlowEffect>,
    pub inner_shadow: Option<InnerShadowEffect>,
    pub outer_shadow: Option<OuterShadowEffect>,
    pub preset_shadow: Option<PresetShadowEffect>,
    pub reflection: Option<ReflectionEffect>,
    pub soft_edges: Option<SoftEdgesEffect>,
}

impl EffectList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        trace!("parsing EffectList '{}'", xml_node.name);
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "blur" => instance.blur = Some(BlurEffect::from_xml_element(child_node)?),
                    "fillOverlay" => instance.fill_overlay = Some(FillOverlayEffect::from_xml_element(child_node)?),
                    "glow" => instance.glow = Some(GlowEffect::from_xml_element(child_node)?),
                    "innerShdw" => instance.inner_shadow = Some(InnerShadowEffect::from_xml_element(child_node)?),
                    "outerShdw" => instance.outer_shadow = Some(OuterShadowEffect::from_xml_element(child_node)?),
                    "prstShdw" => instance.preset_shadow = Some(PresetShadowEffect::from_xml_element(child_node)?),
                    "reflection" => instance.reflection = Some(ReflectionEffect::from_xml_element(child_node)?),
                    "softEdge" => instance.soft_edges = Some(SoftEdgesEffect::from_xml_element(child_node)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

/// This element specifies an Effect Container. It is a list of effects.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct EffectContainer {
    /// Specifies the kind of container, either sibling or tree.
    pub container_type: Option<EffectContainerType>,

    /// Specifies an optional name for this list of effects, so that it can be referred to later. Shall
    /// be unique across all effect trees and effect containers.
    pub name: Option<String>,

    /// Specifies the effects contained in this container.
    pub effects: Vec<Effect>,
}

impl EffectContainer {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<EffectContainer> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "type" => instance.container_type = Some(value.parse::<EffectContainerType>()?),
                    "name" => instance.name = Some(value.clone()),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|mut instance| {
                instance.effects = xml_node
                    .child_nodes
                    .iter()
                    .filter_map(Effect::try_from_xml_element)
                    .collect::<Result<Vec<_>>>()?;

                Ok(instance)
            })
    }
}

/// This element defines a gradient fill.
///
/// A gradient fill is a fill which is characterized by a smooth gradual transition from one color to the next. At its
/// simplest, it is a fill which transitions between two colors; or more generally, it can be a transition of any number
/// of colors.
///
/// The desired transition colors and locations are defined in the gradient stop list (gsLst) child element.
///
/// The other child element defines the properties of the gradient fill (there are two styles-- a linear shade style as
/// well as a path shade style)
#[derive(Default, Debug, Clone, PartialEq)]
pub struct GradientFillProperties {
    /// Specifies the direction(s) in which to flip the gradient while tiling.
    ///
    /// Normally a gradient fill encompasses the entire bounding box of the shape which
    /// contains the fill. However, with the tileRect element, it is possible to define a "tile"
    /// rectangle which is smaller than the bounding box. In this situation, the gradient fill is
    /// encompassed within the tile rectangle, and the tile rectangle is tiled across the bounding
    /// box to fill the entire area.
    pub flip: Option<TileFlipMode>,

    /// Specifies if a fill rotates along with a shape when the shape is rotated.
    pub rotate_with_shape: Option<bool>,

    /// The list of gradient stops that specifies the gradient colors and their relative positions in the color band.
    pub gradient_stop_list: Option<Vec<GradientStop>>,

    /// Specifies the shade properties.
    pub shade_properties: Option<ShadeProperties>,
    /// This element specifies a rectangular region of the shape to which the gradient is applied. This region is then
    /// tiled across the remaining area of the shape to complete the fill. The tile rectangle is defined by percentage
    /// offsets from the sides of the shape's bounding box.
    ///
    /// Each edge of the tile rectangle is defined by a percentage offset from the corresponding edge of the bounding
    /// box. A positive percentage specifies an inset, while a negative percentage specifies an outset.
    ///
    /// # Note
    ///
    /// For example, a left offset of 25% specifies that the left edge of the tile rectangle is located to the right of the
    /// bounding box's left edge by an amount equal to 25% of the bounding box's width.
    pub tile_rect: Option<RelativeRect>,
}

impl GradientFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_str() {
                    "flip" => instance.flip = Some(value.parse()?),
                    "rotWithShape" => instance.rotate_with_shape = Some(parse_xml_bool(value)?),
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
                            "gsLst" => {
                                let gradient_stop_list = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|gs_node| gs_node.local_name() == "gs")
                                    .map(GradientStop::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;

                                match gradient_stop_list.len() {
                                    len if len >= 2 => instance.gradient_stop_list = Some(gradient_stop_list),
                                    len => {
                                        return Err(Box::<dyn Error>::from(LimitViolationError::new(
                                            xml_node.name.clone(),
                                            "gsLst",
                                            2,
                                            MaxOccurs::Unbounded,
                                            len as u32,
                                        )))
                                    }
                                }
                            }
                            "tileRect" => instance.tile_rect = Some(RelativeRect::from_xml_element(child_node)?),
                            local_name if ShadeProperties::is_choice_member(local_name) => {
                                instance.shade_properties = Some(ShadeProperties::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

/// Blip
#[derive(Default, Debug, Clone, PartialEq)]
pub struct Blip {
    /// Specifies the identification information for an embedded picture. This attribute is used to
    /// specify an image that resides locally within the file.
    pub embed_rel_id: Option<RelationshipId>,

    /// Specifies the identification information for a linked picture. This attribute is used to
    /// specify an image that does not reside within this file.
    pub linked_rel_id: Option<RelationshipId>,

    /// Specifies the compression state with which the picture is stored. This allows the
    /// application to specify the amount of compression that has been applied to a picture.
    pub compression: Option<BlipCompression>,
    pub effects: Vec<BlipEffect>,
}

impl Blip {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "r:embed" => instance.embed_rel_id = Some(value.clone()),
                    "r:link" => instance.linked_rel_id = Some(value.clone()),
                    "cstate" => instance.compression = Some(value.parse::<BlipCompression>()?),
                    _ => (),
                }

                Ok(instance)
            })
            .and_then(|instance| {
                let effects = xml_node
                    .child_nodes
                    .iter()
                    .filter_map(BlipEffect::try_from_xml_element)
                    .collect::<Result<Vec<_>>>()?;

                Ok(Self { effects, ..instance })
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct BlipFillProperties {
    /// Specifies the DPI (dots per inch) used to calculate the size of the blip. If not present or
    /// zero, the DPI in the blip is used.
    ///
    /// # Note
    ///
    /// This attribute is primarily used to keep track of the picture quality within a
    /// document. There are different levels of quality needed for print than on-screen viewing
    /// and thus a need to track this information.
    pub dpi: Option<u32>,

    /// Specifies that the fill should rotate with the shape. That is, when the shape that has been
    /// filled with a picture and the containing shape (say a rectangle) is transformed with a
    /// rotation then the fill is transformed with the same rotation.
    pub rotate_with_shape: Option<bool>,

    /// This element specifies the existence of an image (binary large image or picture) and contains a reference to the
    /// image data.
    pub blip: Option<Box<Blip>>,

    /// This element specifies the portion of the blip used for the fill.
    ///
    /// Each edge of the source rectangle is defined by a percentage offset from the corresponding edge of the
    /// bounding box. A positive percentage specifies an inset, while a negative percentage specifies an outset.
    ///
    /// # Note
    ///
    /// For example, a left offset of 25% specifies that the left edge of the source rectangle is located to the right of the
    /// bounding box's left edge by an amount equal to 25% of the bounding box's width.
    pub source_rect: Option<RelativeRect>,

    /// Specifies the fill mode of this blip.
    pub fill_mode_properties: Option<FillModeProperties>,
}

impl BlipFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "dpi" => instance.dpi = Some(value.parse()?),
                    "rotWithShape" => instance.rotate_with_shape = Some(parse_xml_bool(value)?),
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
                            "blip" => instance.blip = Some(Box::new(Blip::from_xml_element(child_node)?)),
                            "srcRect" => instance.source_rect = Some(RelativeRect::from_xml_element(child_node)?),
                            local_name if FillModeProperties::is_choice_member(local_name) => {
                                instance.fill_mode_properties = Some(FillModeProperties::from_xml_element(child_node)?)
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

/// This element specifies a dash stop primitive. Dashing schemes are built by specifying an ordered list of dash stop
/// primitive. A dash stop primitive consists of a dash and a space.
#[derive(Debug, Clone, PartialEq)]
pub struct DashStop {
    /// Specifies the length of the dash relative to the line width.
    pub dash_length: PositivePercentage,

    /// Specifies the length of the space relative to the line width.
    pub space_length: PositivePercentage,
}

impl DashStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<DashStop> {
        let mut opt_dash_length = None;
        let mut opt_space_length = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "d" => opt_dash_length = Some(value.parse::<PositivePercentage>()?),
                "sp" => opt_space_length = Some(value.parse::<PositivePercentage>()?),
                _ => (),
            }
        }

        let dash_length = opt_dash_length.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "d"))?;
        let space_length = opt_space_length.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "sp"))?;

        Ok(Self {
            dash_length,
            space_length,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    /// The position of this gradient stop.
    pub position: PositiveFixedPercentage,

    /// The color of this gradient stop.
    pub color: Color,
}

impl GradientStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let position = xml_node
            .attributes
            .get("pos")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "pos"))?
            .parse()?;

        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;

        Ok(Self { position, color })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct LineEndProperties {
    /// Specifies the line end decoration, such as a triangle or arrowhead.
    pub end_type: Option<LineEndType>,

    /// Specifies the line end width in relation to the line width.
    pub width: Option<LineEndWidth>,

    /// Specifies the line end length in relation to the line width.
    pub length: Option<LineEndLength>,
}

impl LineEndProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<LineEndProperties> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "type" => instance.end_type = Some(value.parse::<LineEndType>()?),
                    "width" => instance.width = Some(value.parse::<LineEndWidth>()?),
                    "length" => instance.length = Some(value.parse::<LineEndLength>()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct LinearShadeProperties {
    /// Specifies the direction of color change for the gradient. To define this angle, let its value
    /// be x measured clockwise. Then ( -sin x, cos x ) is a vector parallel to the line of constant
    /// color in the gradient fill.
    pub angle: Option<PositiveFixedAngle>,

    /// Whether the gradient angle scales with the fill region. Mathematically, if this flag is true,
    /// then the gradient vector ( cos x , sin x ) is scaled by the width (w) and height (h) of the fill
    /// region, so that the vector becomes ( w cos x, h sin x ) (before normalization). Observe
    /// that now if the gradient angle is 45 degrees, the gradient vector is ( w, h ), which goes
    /// from top-left to bottom-right of the fill region. If this flag is false, the gradient angle is
    /// independent of the fill region and is not scaled using the manipulation described above.
    /// So a 45-degree gradient angle always give a gradient band whose line of constant color is
    /// parallel to the vector (1, -1).
    pub scaled: Option<bool>,
}

impl LinearShadeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "ang" => instance.angle = Some(value.parse::<PositiveFixedAngle>()?),
                    "scaled" => instance.scaled = Some(parse_xml_bool(value)?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct PathShadeProperties {
    /// Specifies the shape of the path to follow.
    pub path: Option<PathShadeType>,

    /// This element defines the "focus" rectangle for the center shade, specified relative to the fill tile rectangle. The
    /// center shade fills the entire tile except the margins specified by each attribute.
    ///
    /// Each edge of the center shade rectangle is defined by a percentage offset from the corresponding edge of the
    /// tile rectangle. A positive percentage specifies an inset, while a negative percentage specifies an outset.
    ///
    /// # Note
    ///
    /// For example, a left offset of 25% specifies that the left edge of the center shade rectangle is located to the right
    /// of the tile rectangle's left edge by an amount equal to 25% of the tile rectangle's width.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:path path="rect">
    ///   <a:fillToRect l="50000" r="50000" t="50000" b="50000"/>
    /// </a:path>
    /// ```
    ///
    /// In the above shape, the rectangle defined by fillToRect is a single point in the center of the shape. This creates
    /// the effect of the center shade focusing at a point in the center of the region.
    pub fill_to_rect: Option<RelativeRect>,
}

impl PathShadeProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let path = xml_node.attributes.get("path").map(|value| value.parse()).transpose()?;

        let fill_to_rect = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "fillToRect")
            .map(RelativeRect::from_xml_element)
            .transpose()?;

        Ok(Self { path, fill_to_rect })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ShadeProperties {
    /// This element specifies a linear gradient.
    Linear(LinearShadeProperties),

    /// This element defines that a gradient fill follows a path vs. a linear line.
    Path(PathShadeProperties),
}

impl XsdType for ShadeProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "lin" => Ok(ShadeProperties::Linear(LinearShadeProperties::from_xml_element(
                xml_node,
            )?)),
            "path" => Ok(ShadeProperties::Path(PathShadeProperties::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ShadeProperties").into()),
        }
    }
}

impl XsdChoice for ShadeProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "lin" | "path" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct PatternFillProperties {
    /// Specifies one of a set of preset patterns to fill the object.
    pub preset: Option<PresetPatternVal>,

    /// This element specifies the foreground color of a pattern fill.
    pub fg_color: Option<Color>,

    /// This element specifies the background color of a Pattern fill.
    pub bg_color: Option<Color>,
}

impl PatternFillProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let preset = xml_node.attributes.get("prst").map(|value| value.parse()).transpose()?;

        xml_node.child_nodes.iter().try_fold(
            Self {
                preset,
                ..Default::default()
            },
            |mut instance, child_node| {
                match child_node.local_name() {
                    "fgClr" => {
                        let fg_color = child_node
                            .child_nodes
                            .iter()
                            .find_map(Color::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;

                        instance.fg_color = Some(fg_color);
                    }
                    "bgClr" => {
                        let bg_color = child_node
                            .child_nodes
                            .iter()
                            .find_map(Color::try_from_xml_element)
                            .transpose()?
                            .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;

                        instance.bg_color = Some(bg_color);
                    }
                    _ => (),
                }

                Ok(instance)
            },
        )
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FillProperties {
    /// This element specifies that no fill is applied to the parent element.
    NoFill,

    /// This element specifies a solid color fill. The shape is filled entirely with the specified color.
    SolidFill(Color),

    /// This element specifies a gradient color fill.
    GradientFill(Box<GradientFillProperties>),

    /// This element specifies the type of picture fill that the picture object has. Because a picture has a picture fill
    /// already by default, it is possible to have two fills specified for a picture object. An example of this is shown
    /// below.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   ...
    ///   <p:blipFill>
    ///     <a:blip r:embed="rId2"/>
    ///     <a:stretch>
    ///       <a:fillRect/>
    ///     </a:stretch>
    ///   </p:blipFill>
    ///   ...
    /// </p:pic>
    /// ```
    BlipFill(Box<BlipFillProperties>),

    /// This element specifies a pattern fill. A repeated pattern is used to fill the object.
    PatternFill(Box<PatternFillProperties>),

    /// This element specifies a group fill. When specified, this setting indicates that the parent element is part of a
    /// group and should inherit the fill properties of the group.
    GroupFill,
}

impl XsdType for FillProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "noFill" => Ok(FillProperties::NoFill),
            "solidFill" => {
                let color = xml_node
                    .child_nodes
                    .iter()
                    .find_map(Color::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;

                Ok(FillProperties::SolidFill(color))
            }
            "gradFill" => Ok(FillProperties::GradientFill(Box::new(
                GradientFillProperties::from_xml_element(xml_node)?,
            ))),
            "blipFill" => Ok(FillProperties::BlipFill(Box::new(
                BlipFillProperties::from_xml_element(xml_node)?,
            ))),
            "pattFill" => Ok(FillProperties::PatternFill(Box::new(
                PatternFillProperties::from_xml_element(xml_node)?,
            ))),
            "grpFill" => Ok(FillProperties::GroupFill),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_FillProperties").into()),
        }
    }
}

impl XsdChoice for FillProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "noFill" | "solidFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum LineJoinProperties {
    /// This element specifies that lines joined together have a round join.
    Round,

    /// This element specifies a Bevel Line Join.
    ///
    /// A bevel joint specifies that an angle joint is used to connect lines.
    Bevel,

    /// This element specifies that a line join shall be mitered.
    ///
    /// The value specifies the amount by which lines is extended to form a miter join - otherwise miter
    /// joins can extend infinitely far (for lines which are almost parallel).
    Miter(Option<PositivePercentage>),
}

impl XsdType for LineJoinProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<LineJoinProperties> {
        match xml_node.local_name() {
            "round" => Ok(LineJoinProperties::Round),
            "bevel" => Ok(LineJoinProperties::Bevel),
            "miter" => {
                let lim = xml_node.attributes.get("lim").map(|value| value.parse()).transpose()?;

                Ok(LineJoinProperties::Miter(lim))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineJoinProperties").into()),
        }
    }
}

impl XsdChoice for LineJoinProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "round" | "bevel" | "miter" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct StretchInfoProperties {
    /// This element specifies a fill rectangle. When stretching of an image is specified, a source rectangle, srcRect, is
    /// scaled to fit the specified fill rectangle.
    ///
    /// Each edge of the fill rectangle is defined by a percentage offset from the corresponding edge of the shape's
    /// bounding box. A positive percentage specifies an inset, while a negative percentage specifies an outset.
    ///
    /// # Note
    ///
    /// For example, a left offset of 25% specifies that the left edge of the fill rectangle is located to the right of the
    /// bounding box's left edge by an amount equal to 25% of the bounding box's width.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:blipFill>
    ///   <a:blip r:embed="rId2"/>
    ///   <a:stretch>
    ///     <a:fillRect b="10000" r="25000"/>
    ///   </a:stretch>
    /// </a:blipFill>
    /// ```
    pub fill_rect: Option<RelativeRect>,
}

impl StretchInfoProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let fill_rect = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "fillRect")
            .map(RelativeRect::from_xml_element)
            .transpose()?;

        Ok(Self { fill_rect })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TileInfoProperties {
    /// Specifies additional horizontal offset after alignment.
    pub translate_x: Option<Coordinate>,

    /// Specifies additional vertical offset after alignment.
    pub translate_y: Option<Coordinate>,

    /// Specifies the amount to horizontally scale the srcRect.
    pub scale_x: Option<Percentage>,

    /// Specifies the amount to vertically scale the srcRect.
    pub scale_y: Option<Percentage>,

    /// Specifies the direction(s) in which to flip the source image while tiling. Images can be
    /// flipped horizontally, vertically, or in both directions to fill the entire region.
    pub flip_mode: Option<TileFlipMode>,

    /// Specifies where to align the first tile with respect to the shape. Alignment happens after
    /// the scaling, but before the additional offset.
    pub alignment: Option<RectAlignment>,
}

impl TileInfoProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "tx" => instance.translate_x = Some(value.parse()?),
                    "ty" => instance.translate_y = Some(value.parse()?),
                    "sx" => instance.scale_x = Some(value.parse()?),
                    "sy" => instance.scale_y = Some(value.parse()?),
                    "flip" => instance.flip_mode = Some(value.parse()?),
                    "algn" => instance.alignment = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FillModeProperties {
    /// This element specifies that a BLIP should be tiled to fill the available space. This element defines a "tile"
    /// rectangle within the bounding box. The image is encompassed within the tile rectangle, and the tile rectangle is
    /// tiled across the bounding box to fill the entire area.
    Tile(Box<TileInfoProperties>),

    /// This element specifies that a BLIP should be stretched to fill the target rectangle. The other option is a tile where
    /// a BLIP is tiled to fill the available area.
    Stretch(Box<StretchInfoProperties>),
}

impl XsdType for FillModeProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "tile" => Ok(FillModeProperties::Tile(Box::new(
                TileInfoProperties::from_xml_element(xml_node)?,
            ))),
            "stretch" => Ok(FillModeProperties::Stretch(Box::new(
                StretchInfoProperties::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_FillModeProperties").into()),
        }
    }
}

impl XsdChoice for FillModeProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "tile" | "stretch" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum LineFillProperties {
    /// This element specifies that no fill is applied to the parent element.
    NoFill,

    /// This element specifies a solid color fill. The shape is filled entirely with the specified color.
    SolidFill(Color),

    /// This element specifies a gradient color fill.
    GradientFill(Box<GradientFillProperties>),

    /// This element specifies a pattern color fill.
    PatternFill(Box<PatternFillProperties>),
}

impl XsdType for LineFillProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<LineFillProperties> {
        match xml_node.local_name() {
            "noFill" => Ok(LineFillProperties::NoFill),
            "solidFill" => {
                let color = xml_node
                    .child_nodes
                    .iter()
                    .find_map(Color::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;

                Ok(LineFillProperties::SolidFill(color))
            }
            "gradFill" => Ok(LineFillProperties::GradientFill(Box::new(
                GradientFillProperties::from_xml_element(xml_node)?,
            ))),
            "pattFill" => Ok(LineFillProperties::PatternFill(Box::new(
                PatternFillProperties::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineFillProperties").into()),
        }
    }
}

impl XsdChoice for LineFillProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "noFill" | "solidFill" | "gradFill" | "pattFill" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum LineDashProperties {
    /// This element specifies that a preset line dashing scheme should be used.
    PresetDash(PresetLineDashVal),

    /// This element specifies a custom dashing scheme. It is a list of dash stop elements which represent building block
    /// atoms upon which the custom dashing scheme is built.
    CustomDash(Vec<DashStop>),
}

impl XsdChoice for LineDashProperties {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "prstDash" | "custDash" => true,
            _ => false,
        }
    }
}

impl XsdType for LineDashProperties {
    fn from_xml_element(xml_node: &XmlNode) -> Result<LineDashProperties> {
        match xml_node.local_name() {
            "prstDash" => {
                let val = xml_node
                    .attributes
                    .get("val")
                    .map(|value| value.parse())
                    .transpose()?
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

                Ok(LineDashProperties::PresetDash(val))
            }
            "custDash" => {
                let dash_vec = xml_node
                    .child_nodes
                    .iter()
                    .filter(|child_node| child_node.local_name() == "ds")
                    .map(DashStop::from_xml_element)
                    .collect::<Result<Vec<_>>>()?;

                Ok(LineDashProperties::CustomDash(dash_vec))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_LineDashProperties").into()),
        }
    }
}
