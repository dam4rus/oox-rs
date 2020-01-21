use super::{paragraphs::TextParagraphProperties, runformatting::TextFont};
use crate::{
    shared::drawingml::{
        colors::Color,
        shapeprops::Blip,
        simpletypes::{TextAutonumberScheme, TextBulletSizePercent, TextBulletStartAtNum, TextFontSize},
    },
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone, PartialEq)]
pub enum TextBulletColor {
    /// This element specifies that the color of the bullets for a paragraph should be of the same color as the text run
    /// within which each bullet is contained.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///     <a:buClrTx>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The color of the above bullet follows the default text color of the text for the run of text shown above since no
    /// specific text color was specified.
    FollowText,

    /// This element specifies the color to be used on bullet characters within a given paragraph. The color is specified
    /// using the numerical RGB color format.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buClr>
    ///         <a:srgbClr val="FFFF00"/>
    ///       </a:buClr>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The color of the above bullet does not follow the text color but instead has a yellow color specified by
    /// val="FFFF00". This color should only apply to the actual bullet character and not to the text within the bullet.
    Color(Color),
}

impl XsdType for TextBulletColor {
    fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletColor> {
        match xml_node.local_name() {
            "buClrTx" => Ok(TextBulletColor::FollowText),
            "buClr" => {
                let color = xml_node
                    .child_nodes
                    .iter()
                    .find_map(Color::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "color"))?;

                Ok(TextBulletColor::Color(color))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletColor").into()),
        }
    }
}

impl XsdChoice for TextBulletColor {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "buClrTx" | "buClr" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextBulletSize {
    /// This element specifies that the size of the bullets for a paragraph should be of the same point size as the text run
    /// within which each bullet is contained.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buSzTx>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The size of the above bullet follows the default text size of the text for the run of text shown above since no
    /// specific text size was specified.
    FollowText,

    /// This element specifies the size in percentage of the surrounding text to be used on bullet characters within a
    /// given paragraph.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buSzPct val="111%"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The size of the above bullet follows the text size in that it is always rendered at 111% the size of the text within
    /// the given text run. This is specified by val="111%", with a restriction on the values not being less than 25% or
    /// more than 400%. This percentage size should only apply to the actual bullet character and not to the text within
    /// the bullet.
    Percent(TextBulletSizePercent),

    /// This element specifies the size in points to be used on bullet characters within a given paragraph. The size is
    /// specified using the points where 100 is equal to 1 point font and 1200 is equal to 12 point font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buSzPts val="1400"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The size of the above bullet does not follow the text size of the text within the given text run. The bullets size is
    /// specified by val="1400", which corresponds to a point size of 14. This bullet size should only apply to the actual
    /// bullet character and not to the text within the bullet.
    Point(TextFontSize),
}

impl XsdType for TextBulletSize {
    fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletSize> {
        match xml_node.local_name() {
            "buSzTx" => Ok(TextBulletSize::FollowText),
            "buSzPct" => {
                let val = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?
                    .parse()?;

                Ok(TextBulletSize::Percent(val))
            }
            "buSzPts" => {
                let val = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?
                    .parse()?;

                Ok(TextBulletSize::Point(val))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletSize").into()),
        }
    }
}

impl XsdChoice for TextBulletSize {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "buSzTx" | "buSzPct" | "buSzPts" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextBulletTypeface {
    /// This element specifies that the font of the bullets for a paragraph should be of the same font as the text run
    /// within which each bullet is contained.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buFontTx>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The font of the above bullet follows the default text font of the text for the run of text shown above since no
    /// specific text font was specified.
    FollowText,

    /// This element specifies the font to be used on bullet characters within a given paragraph. The font is specified
    /// using the typeface that it is registered as within the generating application.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buFont typeface="Arial"/>
    ///       <a:buChar char="g"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The font of the above bullet does not follow the text font but instead has Arial font specified by
    /// typeface="Arial". This font should only apply to the actual bullet character and not to the text within the bullet.
    Font(TextFont),
}

impl XsdType for TextBulletTypeface {
    fn from_xml_element(xml_node: &XmlNode) -> Result<TextBulletTypeface> {
        match xml_node.local_name() {
            "buFontTx" => Ok(TextBulletTypeface::FollowText),
            "buFont" => Ok(TextBulletTypeface::Font(TextFont::from_xml_element(xml_node)?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBulletTypeface").into()),
        }
    }
}

impl XsdChoice for TextBulletTypeface {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "buFontTx" | "buFont" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextBullet {
    /// This element specifies that the paragraph within which it is applied is to have no bullet formatting applied to it.
    /// That is to say that there should be no bulleting found within the paragraph where this element is specified.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buNone/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph is formatted with no bullets.
    None,

    /// This element specifies that automatic numbered bullet points should be applied to a paragraph. These are not
    /// just numbers used as bullet points but instead automatically assigned numbers that are based on both
    /// buAutoNum attributes and paragraph level.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buAutoNum type="arabicPeriod"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr lvl="1"…>
    ///       <a:buAutoNum type="arabicPeriod"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 2</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buAutoNum type="arabicPeriod"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 3</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// For the above text there are a total of three bullet points. Two of which are at lvl="0" and one at lvl="1". Due to
    /// this breakdown of levels, the numbering sequence that should be automatically applied is 1, 1, 2 as is shown in
    /// the picture above.
    AutoNumbered(TextAutonumberedBullet),

    /// This element specifies that a character be applied to a set of bullets. These bullets are allowed to be any
    /// character in any font that the system is able to support. If no bullet font is specified along with this element then
    /// the paragraph font is used.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buFont typeface="Calibri"/>
    ///       <a:buChar char="g"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr lvl="1"…>
    ///       <a:buFont typeface="Calibri"/>
    ///       <a:buChar char="g"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 2</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buFont typeface="Calibri"/>
    ///       <a:buChar char="g"/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 3</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// For the above text there are a total of three bullet points. Two of which are at lvl="0" and one at lvl="1".
    /// Because the same character is specified for each bullet the levels do not stand out here. The only difference is
    /// the indentation as shown in the picture above.
    Character(String),

    /// This element specifies that a picture be applied to a set of bullets. This element allows for any standard picture
    /// format graphic to be used instead of the typical bullet characters. This opens up the possibility for bullets to be
    /// anything the generating application would seek to apply.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buBlip>
    ///         <a:blip r:embed="rId2"/>
    ///       </a:buBlip>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 1</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr lvl="1"…>
    ///       <a:buBlip>
    ///         <a:blip r:embed="rId2"/>
    ///       </a:buBlip>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 2</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:buBlip>
    ///         <a:blip r:embed="rId2"/>
    ///       </a:buBlip>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Bullet 3</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// For the above text there are a total of three bullet points. Two of which are at lvl="0" and one at lvl="1".
    /// Because the same picture is specified for each bullet the levels do not stand out here. The only difference is the
    /// indentation as shown in the picture above.
    Picture(Box<Blip>),
}

impl XsdType for TextBullet {
    fn from_xml_element(xml_node: &XmlNode) -> Result<TextBullet> {
        match xml_node.local_name() {
            "buNone" => Ok(TextBullet::None),
            "buAutoNum" => Ok(TextBullet::AutoNumbered(TextAutonumberedBullet::from_xml_element(
                xml_node,
            )?)),
            "buChar" => {
                let character = xml_node
                    .attributes
                    .get("char")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "char"))?
                    .clone();

                Ok(TextBullet::Character(character))
            }
            "buBlip" => {
                let blip = xml_node
                    .child_nodes
                    .iter()
                    .find(|child_node| child_node.local_name() == "blip")
                    .ok_or_else(|| {
                        Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "EG_TextBullet"))
                    })
                    .and_then(Blip::from_xml_element)?;

                Ok(TextBullet::Picture(Box::new(blip)))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextBullet").into()),
        }
    }
}

impl XsdChoice for TextBullet {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "buNone" | "buAutoNum" | "buChar" | "buBlip" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextAutonumberedBullet {
    /// Specifies the numbering scheme that is to be used. This allows for the describing of
    /// formats other than strictly numbers. For instance, a set of bullets can be represented by a
    /// series of Roman numerals instead of the standard 1,2,3,etc. number set.
    pub scheme: TextAutonumberScheme,

    /// Specifies the number that starts a given sequence of automatically numbered bullets.
    /// When the numbering is alphabetical, the number should map to the appropriate letter.
    /// For instance 1 maps to 'a', 2 to 'b' and so on. If the numbers are larger than 26, then
    /// multiple letters should be used. For instance 27 should be represented as 'aa' and
    /// similarly 53 should be 'aaa'.
    pub start_at: Option<TextBulletStartAtNum>,
}

impl TextAutonumberedBullet {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextAutonumberedBullet> {
        let mut scheme = None;
        let mut start_at = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "type" => scheme = Some(value.parse()?),
                "startAt" => start_at = Some(value.parse()?),
                _ => (),
            }
        }

        let scheme = scheme.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self { scheme, start_at })
    }
}

/// This element specifies the list of styles associated with this body of text.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextListStyle {
    /// This element specifies the paragraph properties that are to be applied when no other paragraph properties have
    /// been specified. If this attribute is omitted, then it is left to the application to decide the set of default paragraph
    /// properties that should be applied.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:lstStyle>
    ///     <a:defPPr>
    ///       <a:buNone/>
    ///     </a:defPPr>
    ///   </a:lstStyle>
    ///   <a:p>
    ///     …
    ///     <a:t>Sample Text</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph follows the properties described in defPPr if no overriding properties are specified within
    /// the pPr element.
    pub def_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="0". There
    /// are a total of 9 level text property elements allowed, levels 0-8. It is recommended that the order in which this
    /// and other level property elements are specified be in order of increasing level. That is lvl2pPr should come
    /// before lvl3pPr. This allows the lower level properties to take precedence over the higher level ones because
    /// they are parsed first
    ///
    /// # Xml example
    ///
    /// Consider the following DrawingML code that would specify a paragraph to follow the level style
    /// defined in lvl1pPr and thus create a paragraph of text that has no bullets and is right aligned.
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:lstStyle>
    ///     <a:lvl1pPr algn="r">
    ///       <a:buNone/>
    ///     </a:lvl1pPr>
    ///   </a:lstStyle>
    ///   <a:p>
    ///     <a:pPr lvl="0">
    ///     </a:pPr>
    ///     …
    ///     <a:t>Some text</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// # Note
    ///
    /// To resolve conflicting paragraph properties the linear hierarchy of paragraph properties should be
    /// examined starting first with the pPr element. The rule here is that properties that are defined at a level closer to
    /// the actual text should take precedence. That is if there is a conflicting property between the pPr and lvl1pPr
    /// elements then the pPr property should take precedence because in the property hierarchy it is closer to the
    /// actual text being represented.
    pub lvl1_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="1".
    pub lvl2_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="2".
    pub lvl3_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="3".
    pub lvl4_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="4".
    pub lvl5_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="5".
    pub lvl6_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="6".
    pub lvl7_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="7".
    pub lvl8_paragraph_props: Option<Box<TextParagraphProperties>>,

    /// This element specifies all paragraph level text properties for all elements that have the attribute lvl="8".
    pub lvl9_paragraph_props: Option<Box<TextParagraphProperties>>,
}

impl TextListStyle {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "defPPr" => {
                        instance.def_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl1pPr" => {
                        instance.lvl1_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl2pPr" => {
                        instance.lvl2_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl3pPr" => {
                        instance.lvl3_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl4pPr" => {
                        instance.lvl4_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl5pPr" => {
                        instance.lvl5_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl6pPr" => {
                        instance.lvl6_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl7pPr" => {
                        instance.lvl7_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl8pPr" => {
                        instance.lvl8_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "lvl9pPr" => {
                        instance.lvl9_paragraph_props =
                            Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}
