use crate::{
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError},
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};
use super::{
    colors::Color,
    core::{LineProperties, ShapeProperties, ShapeStyle},
    shapeprops::{EffectProperties, FillProperties},
    simpletypes::{FontCollectionIndex, StyleMatrixColumnIndex, TextTypeFace},
    text::{bodyformatting::TextBodyProperties, bullet::TextListStyle, runformatting::TextFont},
};
use log::trace;
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, PartialEq)]
pub struct EffectStyleItem {
    pub effect_props: EffectProperties,
    // TODO implement
    //pub scene_3d: Option<Scene3D>,
    //pub shape_3d: Option<Shape3D>,
}

impl EffectStyleItem {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        trace!("parsing EffectStyleItem '{}'", xml_node.name);
        let mut effect_props = None;

        for child_node in &xml_node.child_nodes {
            if EffectProperties::is_choice_member(child_node.local_name()) {
                effect_props = Some(EffectProperties::from_xml_element(child_node)?);
            }
        }

        let effect_props =
            effect_props.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_EffectProperties"))?;

        Ok(Self { effect_props })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct StyleMatrixReference {
    /// Specifies the style matrix index of the style referred to.
    pub index: StyleMatrixColumnIndex,

    /// Specifies the color associated with this style matrix reference.
    pub color: Option<Color>,
}

impl StyleMatrixReference {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let index = xml_node
            .attributes
            .get("idx")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?
            .parse()?;

        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?;

        Ok(Self { index, color })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct StyleMatrix {
    /// Defines the name for the format scheme. The name is simply a human readable string
    /// which identifies the format scheme in the user interface.
    pub name: Option<String>,

    /// This element defines a set of three fill styles that are used within a theme. The three fill styles are arranged in
    /// order from subtle to moderate to intense.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <fillStyleLst>
    ///   <solidFill>
    ///   ...
    ///   </solidFill>
    ///   <gradFill rotWithShape="1">
    ///   ...
    ///   </gradFill>
    ///   <gradFill rotWithShape="1">
    ///   ...
    ///   </gradFill>
    /// </fillStyleLst>
    /// ```
    ///
    /// In this example, we see three fill styles being defined within the fill style list. The first style is the subtle style and
    /// defines simply a solid fill. The second and third styles (moderate and intense fills respectively) define gradient
    /// fills.
    pub fill_style_list: Vec<FillProperties>,

    /// This element defines a list of three line styles for use within a theme. The three line styles are arranged in order
    /// from subtle to moderate to intense versions of lines. This list makes up part of the style matrix.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <lnStyleLst>
    ///   <ln w="9525" cap="flat" cmpd="sng" algn="ctr">
    ///     <solidFill>
    ///       <schemeClr val="phClr">
    ///         <shade val="50000"/>
    ///         <satMod val="103000"/>
    ///       </schemeClr>
    ///     </solidFill>
    ///     <prstDash val="solid"/>
    ///   </ln>
    ///   <ln w="25400" cap="flat" cmpd="sng" algn="ctr">
    ///     <solidFill>
    ///       <schemeClr val="phClr"/>
    ///     </solidFill>
    ///     <prstDash val="solid"/>
    ///   </ln>
    ///   <ln w="38100" cap="flat" cmpd="sng" algn="ctr">
    ///     <solidFill>
    ///       <schemeClr val="phClr"/>
    ///     </solidFill>
    ///     <prstDash val="solid"/>
    ///   </ln>
    /// </lnStyleLst>
    /// ```
    ///
    /// In this example, we see three lines defined within a line style list. The first line corresponds to the subtle line,
    /// the second to the moderate, and the third corresponds to the intense line defined in the theme.
    pub line_style_list: Vec<LineProperties>,

    /// This element defines a set of three effect styles that create the effect style list for a theme. The effect styles are
    /// arranged in order of subtle to moderate to intense.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <effectStyleLst>
    ///   <effectStyle>
    ///     <effectLst>
    ///       <outerShdw blurRad="57150" dist="38100" dir="5400000"
    ///       algn="ctr" rotWithShape="0">
    ///       ...
    ///       </outerShdw>
    ///     </effectLst>
    ///   </effectStyle>
    ///   <effectStyle>
    ///     <effectLst>
    ///       <outerShdw blurRad="57150" dist="38100" dir="5400000"
    ///       algn="ctr" rotWithShape="0">
    ///       ...
    ///       </outerShdw>
    ///     </effectLst>
    ///   </effectStyle>
    ///   <effectStyle>
    ///     <effectLst>
    ///       <outerShdw blurRad="57150" dist="38100" dir="5400000"
    ///       algn="ctr" rotWithShape="0">
    ///       ...
    ///       </outerShdw>
    ///     </effectLst>
    ///     <scene3d>
    ///     ...
    ///     </scene3d>
    ///     <sp3d prstMaterial="powder">
    ///     ...
    ///     </sp3d>
    ///   </effectStyle>
    /// </effectStyleLst>
    /// ```
    ///
    /// In this example, we see three effect styles defined. The first two (subtle and moderate) define an outer shadow
    /// as the effect, while the third effect style (intense) defines an outer shadow along with 3D properties which are
    /// to be applied to the object as well.
    pub effect_style_list: Vec<EffectStyleItem>,

    /// This element defines a list of background fills that are used within a theme. The background fills consist of three
    /// fills, arranged in order from subtle to moderate to intense.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <bgFillStyleLst>
    ///   <solidFill>
    ///   ...
    ///   </solidFill>
    ///   <gradFill rotWithShape="1">
    ///   ...
    ///   </gradFill>
    ///   <blipFill>
    ///   ...
    ///   </blipFill>
    /// </bgFillStyleLst>
    /// ```
    ///
    /// In this example, we see that the list contains a solid fill for the subtle fill, a gradient fill for the moderate fill and
    /// an image fill for the intense background fill.
    pub bg_fill_style_list: Vec<FillProperties>,
}

impl StyleMatrix {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        trace!("parsing StyleMatrix '{}'", xml_node.name);
        let name = xml_node.attributes.get("name").cloned();
        let mut fill_style_list = None;
        let mut line_style_list = None;
        let mut effect_style_list = None;
        let mut bg_fill_style_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "fillStyleLst" => {
                    let vec = child_node
                        .child_nodes
                        .iter()
                        .filter_map(FillProperties::try_from_xml_element)
                        .collect::<Result<Vec<_>>>()?;

                    fill_style_list = match vec.len() {
                        len if len >= 3 => Some(vec),
                        len => {
                            return Err(Box::new(LimitViolationError::new(
                                String::from("fillStyleLst"),
                                "EG_FillProperties",
                                3,
                                MaxOccurs::Unbounded,
                                len as u32,
                            )))
                        }
                    };
                }
                "lnStyleLst" => {
                    let vec = child_node
                        .child_nodes
                        .iter()
                        .filter(|child_node| child_node.local_name() == "ln")
                        .map(LineProperties::from_xml_element)
                        .collect::<Result<Vec<_>>>()?;

                    line_style_list = match vec.len() {
                        len if len >= 3 => Some(vec),
                        len => {
                            return Err(Box::new(LimitViolationError::new(
                                String::from("lnStyleLst"),
                                "ln",
                                3,
                                MaxOccurs::Unbounded,
                                len as u32,
                            )))
                        }
                    };
                }
                "effectStyleLst" => {
                    let vec = child_node
                        .child_nodes
                        .iter()
                        .filter(|child_node| child_node.local_name() == "effectStyle")
                        .map(EffectStyleItem::from_xml_element)
                        .collect::<Result<Vec<_>>>()?;

                    effect_style_list = match vec.len() {
                        len if len >= 3 => Some(vec),
                        len => {
                            return Err(Box::new(LimitViolationError::new(
                                String::from("effectStyleLst"),
                                "effectStyle",
                                3,
                                MaxOccurs::Unbounded,
                                len as u32,
                            )))
                        }
                    };
                }
                "bgFillStyleLst" => {
                    let vec = child_node
                        .child_nodes
                        .iter()
                        .filter_map(FillProperties::try_from_xml_element)
                        .collect::<Result<Vec<_>>>()?;

                    bg_fill_style_list = match vec.len() {
                        len if len >= 3 => Some(vec),
                        len => {
                            return Err(Box::new(LimitViolationError::new(
                                String::from("bgFillStyleLst"),
                                "EG_FillProperties",
                                3,
                                MaxOccurs::Unbounded,
                                len as u32,
                            )))
                        }
                    };
                }
                _ => (),
            }
        }

        let fill_style_list =
            fill_style_list.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fillStyleLst"))?;

        let line_style_list =
            line_style_list.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lnStyleLst"))?;

        let effect_style_list =
            effect_style_list.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "effectStyleLst"))?;

        let bg_fill_style_list =
            bg_fill_style_list.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bgFillStyleLst"))?;

        Ok(Self {
            name,
            fill_style_list,
            line_style_list,
            effect_style_list,
            bg_fill_style_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SupplementalFont {
    /// Specifies the script, or language, in which the typeface is supposed to be used.
    ///
    /// # Note
    ///
    /// It is recommended that script names as specified in ISO 15924 are used.
    pub script: String,

    /// Specifies the font face to use.
    pub typeface: TextTypeFace,
}

impl SupplementalFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut script = None;
        let mut typeface = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "script" => script = Some(value.clone()),
                "typeface" => typeface = Some(value.clone()),
                _ => (),
            }
        }

        let script = script.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "script"))?;
        let typeface = typeface.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "typeface"))?;

        Ok(Self { script, typeface })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FontReference {
    /// Specifies the identifier of the font to reference.
    pub index: FontCollectionIndex,
    pub color: Option<Color>,
}

impl FontReference {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let index = xml_node
            .attributes
            .get("idx")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "idx"))?
            .parse()?;

        let color = xml_node
            .child_nodes
            .iter()
            .find_map(Color::try_from_xml_element)
            .transpose()?;

        Ok(Self { index, color })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FontScheme {
    /// The name of the font scheme shown in the user interface.
    pub name: String,

    /// This element defines the set of major fonts which are to be used under different languages or locals.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <majorFont>
    /// <latin typeface="Calibri"/>
    ///   <ea typeface="Arial"/>
    ///   <cs typeface="Arial"/>
    ///   <font script="Jpan" typeface="MS Pゴシック "/>
    ///   <font script="Hang" typeface="HY중고딕"/>
    ///   <font script="Hans" typeface="隶 书"/>
    ///   <font script="Hant" typeface="微軟正黑體 "/>
    ///   <font script="Arab" typeface="Traditional Arabic"/>
    ///   <font script="Hebr" typeface="Arial"/>
    ///   <font script="Thai" typeface="Cordia New"/>
    ///   <font script="Ethi" typeface="Nyala"/>
    ///   <font script="Beng" typeface="Vrinda"/>
    ///   <font script="Gujr" typeface="Shruti"/>
    ///   <font script="Khmr" typeface="DaunPenh"/>
    ///   <font script="Knda" typeface="Tunga"/>
    /// </majorFont>
    /// ```
    ///
    /// In this example, we see the latin, east asian, and complex script fonts defined along with many fonts for
    /// different locals.
    pub major_font: Box<FontCollection>,

    /// This element defines the set of minor fonts that are to be used under different languages or locals.
    ///
    /// ```xml
    /// <minorFont>
    ///   <latin typeface="Calibri"/>
    ///   <ea typeface="Arial"/>
    ///   <cs typeface="Arial"/>
    ///   <font script="Jpan" typeface="MS Pゴシック "/>
    ///   <font script="Hang" typeface="HY중고딕"/>
    ///   <font script="Hans" typeface="隶 书"/>
    ///   <font script="Hant" typeface="微軟正黑體 "/>
    ///   <font script="Arab" typeface="Traditional Arabic"/>
    ///   <font script="Hebr" typeface="Arial"/>
    ///   <font script="Thai" typeface="Cordia New"/>
    ///   <font script="Ethi" typeface="Nyala"/>
    ///   <font script="Beng" typeface="Vrinda"/>
    ///   <font script="Gujr" typeface="Shruti"/>
    ///   <font script="Khmr" typeface="DaunPenh"/>
    ///   <font script="Knda" typeface="Tunga"/>
    /// </minorFont>
    /// ```
    ///
    /// In this example, we see the latin, east asian, and complex script fonts defined along with many fonts for
    /// different locals.
    pub minor_font: Box<FontCollection>,
}

impl FontScheme {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node
            .attributes
            .get("name")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?
            .clone();

        let mut major_font = None;
        let mut minor_font = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "majorFont" => major_font = Some(Box::new(FontCollection::from_xml_element(child_node)?)),
                "minorFont" => minor_font = Some(Box::new(FontCollection::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let major_font = major_font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "majorFont"))?;
        let minor_font = minor_font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "minorFont"))?;

        Ok(Self {
            name,
            major_font,
            minor_font,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DefaultShapeDefinition {
    /// This element specifies the visual shape properties that can be applied to a shape.
    pub shape_properties: Box<ShapeProperties>,
    pub text_body_properties: Box<TextBodyProperties>,
    pub text_list_style: Box<TextListStyle>,

    /// This element specifies the style information for a shape.
    pub shape_style: Option<Box<ShapeStyle>>,
}

impl DefaultShapeDefinition {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut shape_properties = None;
        let mut text_body_properties = None;
        let mut text_list_style = None;
        let mut shape_style = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "spPr" => shape_properties = Some(Box::new(ShapeProperties::from_xml_element(child_node)?)),
                "bodyPr" => text_body_properties = Some(Box::new(TextBodyProperties::from_xml_element(child_node)?)),
                "lstStyle" => text_list_style = Some(Box::new(TextListStyle::from_xml_element(child_node)?)),
                "style" => shape_style = Some(Box::new(ShapeStyle::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let shape_properties =
            shape_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "spPr"))?;
        let text_body_properties =
            text_body_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "bodyPr"))?;
        let text_list_style =
            text_list_style.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lstStyle"))?;

        Ok(Self {
            shape_properties,
            text_body_properties,
            text_list_style,
            shape_style,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FontCollection {
    /// Specifies the font used for latin characters.
    pub latin: TextFont,

    /// Specifies the font used for east asian characters.
    pub east_asian: TextFont,

    /// Specifies the font used for complex characters.
    pub complex_script: TextFont,

    /// This element defines a list of font within the styles area of DrawingML. A font is defined by a script along
    /// with a typeface.
    pub supplemental_font_list: Vec<SupplementalFont>,
}

impl FontCollection {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut opt_latin = None;
        let mut opt_ea = None;
        let mut opt_cs = None;
        let mut supplemental_font_list = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "latin" => opt_latin = Some(TextFont::from_xml_element(child_node)?),
                "ea" => opt_ea = Some(TextFont::from_xml_element(child_node)?),
                "cs" => opt_cs = Some(TextFont::from_xml_element(child_node)?),
                "font" => supplemental_font_list.push(SupplementalFont::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let latin = opt_latin.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "latin"))?;
        let east_asian = opt_ea.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "ea"))?;
        let complex_script = opt_cs.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "cs"))?;

        Ok(Self {
            latin,
            east_asian,
            complex_script,
            supplemental_font_list,
        })
    }
}
