use super::{
    sharedstylesheet::ColorMapping,
    simpletypes::{
        parse_hex_color_rgb, Angle, FixedPercentage, HexColorRGB, Percentage, PositiveFixedAngle,
        PositiveFixedPercentage, PositivePercentage, PresetColorVal, SchemeColorVal, SystemColorVal,
    },
    util::XmlNodeExt,
};
use crate::{
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[repr(C)]
#[derive(Debug, Clone, PartialEq)]
pub enum ColorTransform {
    /// This element specifies a lighter version of its input color. A 10% tint is 10% of the input color combined with
    /// 90% white.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (BC, FF, BC)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:tint val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Tint(PositiveFixedPercentage),

    /// This element specifies a darker version of its input color. A 10% shade is 10% of the input color combined with
    /// 90% black.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, BC, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:shade val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Shade(PositiveFixedPercentage),

    /// This element specifies that the color rendered should be the complement of its input color with the complement
    /// being defined as such. Two colors are called complementary if, when mixed they produce a shade of grey. For
    /// instance, the complement of red which is RGB (255, 0, 0) is cyan which is RGB (0, 255, 255).
    ///
    /// Primary colors and secondary colors are typically paired in this way:
    /// * red and cyan (where cyan is the mixture of green and blue)
    /// * green and magenta (where magenta is the mixture of red and blue)
    /// * blue and yellow (where yellow is the mixture of red and green)
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="FF0000">
    ///     <a:comp/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Complement,

    /// This element specifies the inverse of its input color.
    ///
    /// # Xml example
    ///
    /// The inverse of red (1, 0, 0) is cyan (0, 1, 1).
    ///
    /// The following represents cyan, the inverse of red:
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="FF0000">
    ///     <a:inv/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Inverse,

    /// This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
    /// primaries.
    Grayscale,

    /// This element specifies its input color with the specific opacity, but with its color unchanged.
    ///
    /// # Xml example
    ///
    /// The following represents a green solid fill which is 50% opaque
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:alpha val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Alpha(PositiveFixedPercentage),

    /// This element specifies a more or less opaque version of its input color. Increases or decreases the input alpha
    /// percentage by the specified percentage offset. A 10% alpha offset increases a 50% opacity to 60%. A -10% alpha
    /// offset decreases a 50% opacity to 40%. The transformed alpha values are limited to a range of 0 to 100%. A 10%
    /// alpha offset increase to a 100% opaque object still results in 100% opacity.
    ///
    /// # Xml example
    ///
    /// The following represents a green solid fill which is 90% opaque
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:alphaOff val="-10000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    AlphaOffset(FixedPercentage),

    /// This element specifies a more or less opaque version of its input color. An alpha modulate never increases the
    /// alpha beyond 100%. A 200% alpha modulate makes a input color twice as opaque as before. A 50% alpha
    /// modulate makes a input color half as opaque as before.
    ///
    /// # Xml example
    ///
    /// The following represents a green solid fill which is 50% opaque
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:alphaMod val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    AlphaModulate(PositivePercentage),

    /// This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
    ///
    /// # Xml example
    ///
    /// The following two solid fills are equivalent.
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="50000">
    /// </a:solidFill>
    /// <a:solidFill>
    ///   <a:hslClr hue="0" sat="100000" lum="50000">
    ///     <a:hue val="14400000"/>
    ///   <a:hslClr/>
    /// </a:solidFill>
    /// ```
    Hue(PositiveFixedAngle),

    /// This element specifies the input color with its hue shifted, but with its saturation and luminance unchanged.
    ///
    /// # Xml example
    ///
    /// The following increases the hue angular value by 10 degrees.
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="0" sat="100000" lum="50000"/>
    ///   <a:hueOff val="600000"/>
    /// </a:solidFill>
    /// ```
    HueOffset(Angle),

    /// This element specifies the input color with its hue modulated by the given percentage. A 50% hue modulate
    /// decreases the angular hue value by half. A 200% hue modulate doubles the angular hue value.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="50000">
    ///     <a:hueMod val="50000"/>
    ///   </a:hslClr>
    /// </a:solidFill>
    /// ```
    HueModulate(PositivePercentage),

    /// This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
    /// Typically saturation values fall in the range [0%, 100%].
    ///
    /// # Xml example
    ///
    /// The following two solid fills are equivalent:
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="50000">
    /// </a:solidFill>
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="0" lum="50000">
    ///     <a:sat val="100000"/>
    ///   <a:hslClr/>
    /// </a:solidFill>
    /// ```
    Saturation(Percentage),

    /// This element specifies the input color with its saturation shifted, but with its hue and luminance unchanged. A
    /// 10% offset to 20% saturation yields 30% saturation.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (19, E5, 19)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:satOff val="-20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    SaturationOffset(Percentage),

    /// This element specifies the input color with its saturation modulated by the given percentage. A 50% saturation
    /// modulate reduces the saturation by half. A 200% saturation modulate doubles the saturation.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (66, 99, 66)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:satMod val="20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    SaturationModulate(Percentage),

    /// This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
    /// Typically luminance values fall in the range [0%, 100%].
    ///
    /// # Xml example
    ///
    /// The following two solid fills are equivalent:
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="50000">
    /// </a:solidFill>
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="0">
    ///     <a:lum val="50000"/>
    ///   <a:hslClr/>
    /// </a:solidFill>
    /// ```
    Luminance(Percentage),

    /// This element specifies the input color with its luminance shifted, but with its hue and saturation unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, 99, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:lumOff val="-20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    LuminanceOffset(Percentage),

    /// This element specifies the input color with its luminance modulated by the given percentage. A 50% luminance
    /// modulate reduces the luminance by half. A 200% luminance modulate doubles the luminance.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, 75, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:lumMod val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    LuminanceModulate(Percentage),

    /// This element specifies the input color with the specified red component, but with its green and blue color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (FF, FF, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:red val="100000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Red(Percentage),

    /// This element specifies the input color with its red component shifted, but with its green and blue color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    /// to value RRGGBB= (CC, 00, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="FF0000">
    ///     <a:redOff val="-20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    RedOffset(Percentage),

    /// This element specifies the input color with its red component modulated by the given percentage. A 50% red
    /// modulate reduces the red component by half. A 200% red modulate doubles the red component.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    /// to value RRGGBB= (80, 00, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="FF0000">
    ///     <a:redMod val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    RedModulate(Percentage),

    /// This elements specifies the input color with the specified green component, but with its red and blue color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    /// to value RRGGBB= (00, FF, FF)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="0000FF">
    ///     <a:green val="100000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Green(Percentage),

    /// This element specifies the input color with its green component shifted, but with its red and blue color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, CC, 00).
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:greenOff val="-20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    GreenOffset(Percentage),

    /// This element specifies the input color with its green component modulated by the given percentage. A 50%
    /// green modulate reduces the green component by half. A 200% green modulate doubles the green component.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, 80, 00)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:greenMod val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    GreenModulate(Percentage),

    /// This element specifies the input color with the specific blue component, but with the red and green color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    /// to value RRGGBB= (00, FF, FF)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:blue val="100000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    Blue(Percentage),

    /// This element specifies the input color with its blue component shifted, but with its red and green color
    /// components unchanged.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    /// to value RRGGBB= (00, 00, CC)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="00FF00">
    ///     <a:blueOff val="-20000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    BlueOffset(Percentage),

    /// This element specifies the input color with its blue component modulated by the given percentage. A 50% blue
    /// modulate reduces the blue component by half. A 200% blue modulate doubles the blue component.
    ///
    /// # Xml example
    ///
    /// The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    /// to value RRGGBB= (00, 00, 80)
    /// ```xml
    /// <a:solidFill>
    ///   <a:srgbClr val="0000FF">
    ///     <a:blueMod val="50000"/>
    ///   </a:srgbClr>
    /// </a:solidFill>
    /// ```
    BlueModulate(Percentage),

    /// This element specifies that the output color rendered by the generating application should be the sRGB gamma
    /// shift of the input color.
    Gamma,

    /// This element specifies that the output color rendered by the generating application should be the inverse sRGB
    /// gamma shift of the input color.
    InverseGamma,
}

impl XsdType for ColorTransform {
    fn from_xml_element(xml_node: &XmlNode) -> Result<ColorTransform> {
        match xml_node.local_name() {
            "tint" => Ok(ColorTransform::Tint(xml_node.parse_val_attribute()?)),
            "shade" => Ok(ColorTransform::Shade(xml_node.parse_val_attribute()?)),
            "comp" => Ok(ColorTransform::Complement),
            "inv" => Ok(ColorTransform::Inverse),
            "gray" => Ok(ColorTransform::Grayscale),
            "alpha" => Ok(ColorTransform::Alpha(xml_node.parse_val_attribute()?)),
            "alphaOff" => Ok(ColorTransform::AlphaOffset(xml_node.parse_val_attribute()?)),
            "alphaMod" => Ok(ColorTransform::AlphaModulate(xml_node.parse_val_attribute()?)),
            "hue" => Ok(ColorTransform::Hue(xml_node.parse_val_attribute()?)),
            "hueOff" => Ok(ColorTransform::HueOffset(xml_node.parse_val_attribute()?)),
            "hueMod" => Ok(ColorTransform::HueModulate(xml_node.parse_val_attribute()?)),
            "sat" => Ok(ColorTransform::Saturation(xml_node.parse_val_attribute()?)),
            "satOff" => Ok(ColorTransform::SaturationOffset(xml_node.parse_val_attribute()?)),
            "satMod" => Ok(ColorTransform::SaturationModulate(xml_node.parse_val_attribute()?)),
            "lum" => Ok(ColorTransform::Luminance(xml_node.parse_val_attribute()?)),
            "lumOff" => Ok(ColorTransform::LuminanceOffset(xml_node.parse_val_attribute()?)),
            "lumMod" => Ok(ColorTransform::LuminanceModulate(xml_node.parse_val_attribute()?)),
            "red" => Ok(ColorTransform::Red(xml_node.parse_val_attribute()?)),
            "redOff" => Ok(ColorTransform::RedOffset(xml_node.parse_val_attribute()?)),
            "redMod" => Ok(ColorTransform::RedModulate(xml_node.parse_val_attribute()?)),
            "green" => Ok(ColorTransform::Green(xml_node.parse_val_attribute()?)),
            "greenOff" => Ok(ColorTransform::GreenOffset(xml_node.parse_val_attribute()?)),
            "greenMod" => Ok(ColorTransform::GreenModulate(xml_node.parse_val_attribute()?)),
            "blue" => Ok(ColorTransform::Blue(xml_node.parse_val_attribute()?)),
            "blueOff" => Ok(ColorTransform::BlueOffset(xml_node.parse_val_attribute()?)),
            "blueMod" => Ok(ColorTransform::BlueModulate(xml_node.parse_val_attribute()?)),
            "gamma" => Ok(ColorTransform::Gamma),
            "invGamma" => Ok(ColorTransform::InverseGamma),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorTransform").into()),
        }
    }
}

impl XsdChoice for ColorTransform {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "tint" | "shade" | "comp" | "inv" | "gray" | "alpha" | "alphaOff" | "alphaMod" | "hue" | "hueOff"
            | "hueMod" | "sat" | "satOff" | "satMod" | "lum" | "lumOff" | "lumMod" | "red" | "redOff" | "redMod"
            | "green" | "greenOff" | "greenMod" | "blue" | "blueOff" | "blueMod" | "gamma" | "invGamma" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ScRgbColor {
    /// Specifies the percentage of red.
    pub r: Percentage,

    /// Specifies the percentage of green.
    pub g: Percentage,

    /// Specifies the percentage of blue.
    pub b: Percentage,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl ScRgbColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<ScRgbColor> {
        let mut opt_r = None;
        let mut opt_g = None;
        let mut opt_b = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r" => opt_r = Some(value.parse()?),
                "g" => opt_g = Some(value.parse()?),
                "b" => opt_b = Some(value.parse()?),
                _ => (),
            }
        }

        let r = opt_r.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?;
        let g = opt_g.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "g"))?;
        let b = opt_b.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            r,
            g,
            b,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SRgbColor {
    pub value: u32,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl SRgbColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SRgbColor> {
        let value = xml_node
            .attributes
            .get("val")
            .ok_or_else(|| Box::<dyn Error>::from(MissingAttributeError::new(xml_node.name.clone(), "val")))
            .and_then(|value| u32::from_str_radix(value, 16).map_err(Box::from))?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct HslColor {
    /// Specifies the angular value describing the wavelength. Expressed in 1/6000ths of a
    /// degree.
    pub hue: PositiveFixedAngle,

    /// Specifies the saturation referring to the purity of the hue. Expressed as a percentage with
    /// 0% referring to grey, 100% referring to the purest form of the hue.
    pub saturation: Percentage,

    /// Specifies the luminance referring to the lightness or darkness of the color. Expressed as a
    /// percentage with 0% referring to maximal dark (black) and 100% referring to maximal
    /// white.
    pub luminance: Percentage,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl HslColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<HslColor> {
        let mut opt_h = None;
        let mut opt_s = None;
        let mut opt_l = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "hue" => opt_h = Some(value.parse::<PositiveFixedAngle>()?),
                "sat" => opt_s = Some(value.parse()?),
                "lum" => opt_l = Some(value.parse()?),
                _ => (),
            }
        }

        let hue = opt_h.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hue"))?;
        let saturation = opt_s.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "sat"))?;
        let luminance = opt_l.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "lum"))?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            hue,
            saturation,
            luminance,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SystemColor {
    /// Specifies the system color value.
    pub value: SystemColorVal,

    /// Specifies the color value that was last computed by the generating application.
    pub last_color: Option<HexColorRGB>,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl SystemColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SystemColor> {
        let mut opt_val = None;
        let mut last_color = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "val" => opt_val = Some(value.parse()?),
                "lastClr" => last_color = Some(parse_hex_color_rgb(value)?),
                _ => (),
            }
        }

        let value = opt_val.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            value,
            last_color,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PresetColor {
    pub value: PresetColorVal,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl PresetColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<PresetColor> {
        let value = xml_node.parse_val_attribute()?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SchemeColor {
    pub value: SchemeColorVal,

    /// Color transforms to apply to this color
    pub color_transforms: Vec<ColorTransform>,
}

impl SchemeColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<SchemeColor> {
        let value = xml_node
            .attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?
            .parse()?;

        let color_transforms = xml_node
            .child_nodes
            .iter()
            .filter_map(ColorTransform::try_from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            value,
            color_transforms,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Color {
    /// This element specifies a color using the red, green, blue RGB color model. Each component, red, green, and blue
    /// is expressed as a percentage from 0% to 100%. A linear gamma of 1.0 is assumed.
    ///
    /// Specifies the level of red as expressed by a percentage offset increase or decrease relative to the input color.
    ///
    /// # Xml example
    ///
    /// The following represent the same color
    /// ```xml
    /// <a:solidFill>
    ///   <a:scrgbClr r="50000" g="50000" b="50000"/>
    /// </a:solidFill>
    /// <a:solidFill>
    ///   <a:srgbClr val="BCBCBC"/>
    /// </a:solidFill>
    /// ```
    ScRgbColor(ScRgbColor),

    /// This element specifies a color using the red, green, blue RGB color model. Red, green, and blue is expressed as
    /// sequence of hex digits, RRGGBB. A perceptual gamma of 2.2 is used.
    ///
    /// Specifies the level of red as expressed by a percentage offset increase or decrease relative to the input color.
    ///
    /// # Xml example
    ///
    /// The following represent the same color
    /// ```xml
    /// <a:solidFill>
    ///   <a:scrgbClr r="50000" g="50000" b="50000"/>
    /// </a:solidFill>
    /// <a:solidFill>
    ///   <a:srgbClr val="BCBCBC"/>
    /// </a:solidFill>
    /// ```
    SRgbColor(SRgbColor),

    /// This element specifies a color using the HSL color model. A perceptual gamma of 2.2 is assumed.
    ///
    /// Hue refers to the dominant wavelength of color, saturation refers to the purity of its hue, and luminance refers
    /// to its lightness or darkness.
    ///
    /// As with all colors, colors defined with the HSL color model can have color transforms applied to it.
    ///
    /// # Xml example
    ///
    /// The color blue having RGB value RRGGBB = (00, 00, 80) is equivalent to
    /// ```xml
    /// <a:solidFill>
    ///   <a:hslClr hue="14400000" sat="100000" lum="50000">
    /// </a:solidFill>
    /// ```
    HslColor(HslColor),

    /// This element specifies a color bound to predefined operating system elements.
    ///
    /// # Xml example
    ///
    /// The following represents the default color used for displaying text in a window.
    /// ```xml
    /// <a:solidFill>
    ///   <a:sysClr val="windowText"/>
    /// </a:solidFill>
    /// ```
    SystemColor(SystemColor),

    /// This element specifies a color bound to a user's theme. As with all elements which define a color, it is possible to
    /// apply a list of color transforms to the base color defined.
    ///
    /// # Xml example
    ///
    /// <a:solidFill>
    ///   <a:schemeClr val="lt1"/>
    /// </a:solidFill>
    SchemeColor(SchemeColor),

    /// This element specifies a color which is bound to one of a predefined collection of colors.
    ///
    /// # Xml example
    ///
    /// The following defines a solid fill bound to the "black" preset color.
    /// <a:solidFill>
    ///   <a:prstClr val="black">
    /// </a:solidFill>
    PresetColor(PresetColor),
}

impl XsdType for Color {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Color> {
        match xml_node.local_name() {
            "scrgbClr" => Ok(Color::ScRgbColor(ScRgbColor::from_xml_element(xml_node)?)),
            "srgbClr" => Ok(Color::SRgbColor(SRgbColor::from_xml_element(xml_node)?)),
            "hslClr" => Ok(Color::HslColor(HslColor::from_xml_element(xml_node)?)),
            "sysClr" => Ok(Color::SystemColor(SystemColor::from_xml_element(xml_node)?)),
            "schemeClr" => Ok(Color::SchemeColor(SchemeColor::from_xml_element(xml_node)?)),
            "prstClr" => Ok(Color::PresetColor(PresetColor::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_ColorChoice").into()),
        }
    }
}

impl XsdChoice for Color {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "scrgbClr" | "srgbClr" | "hslClr" | "sysClr" | "schemeClr" | "prstClr" => true,
            _ => false,
        }
    }
}
/// This element defines a custom color. The custom colors are used within a custom color list to define custom
/// colors that are extra colors that can be appended to a theme. This is useful within corporate scenarios where
/// there is a set corporate color palette from which to work.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomColor {
    /// The name of the color shown in the color picker.
    pub name: Option<String>,

    /// The color represented by this custom color.
    pub color: Color,
}

impl CustomColor {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node.attributes.get("name").cloned();
        let color = xml_node
            .child_nodes
            .iter()
            .find(|child_node| Color::is_choice_member(child_node.local_name()))
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "EG_ColorChoice")))
            .and_then(Color::from_xml_element)?;

        Ok(Self { name, color })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ColorMappingOverride {
    /// This element is a part of a choice for which color mapping is used within the document.
    /// If this element is specified, then we specifically use the color mapping defined in the master.
    UseMaster,

    /// This element provides an override for the color mapping in a document. When defined, this color mapping is
    /// used in place of the already defined color mapping, or master color mapping. This color mapping is defined in
    /// the same manner as the other mappings within this document.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
    ///   accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
    ///   accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    /// ```
    Override(Box<ColorMapping>),
}

impl XsdType for ColorMappingOverride {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "masterClrMapping" => Ok(ColorMappingOverride::UseMaster),
            "overrideClrMapping" => Ok(ColorMappingOverride::Override(Box::new(
                ColorMapping::from_xml_element(xml_node)?,
            ))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "CT_ColorMappingOverride").into()),
        }
    }
}

impl XsdChoice for ColorMappingOverride {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "masterClrMapping" | "overrideClrMapping" => true,
            _ => false,
        }
    }
}
