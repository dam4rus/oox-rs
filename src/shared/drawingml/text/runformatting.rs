use super::paragraphs::{TextCharacterProperties, TextField, TextLineBreak};
use crate::{
    shared::drawingml::{
        core::LineProperties,
        shapeprops::FillProperties,
        simpletypes::{Panose, TextTypeFace},
    },
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Debug, Clone, PartialEq)]
pub struct TextFont {
    /// Specifies the typeface, or name of the font that is to be used. The typeface is a string
    /// name of the specific font that should be used in rendering the presentation. If this font is
    /// not available within the font list of the generating application than font substitution logic
    /// should be utilized in order to select an alternate font.
    pub typeface: TextTypeFace,

    /// Specifies the Panose-1 classification number for the current font using the mechanism
    /// defined in §5.2.7.17 of ISO/IEC 14496-22.
    pub panose: Option<Panose>,

    /// Specifies the font pitch as well as the font family for the corresponding font. Because the
    /// value of this attribute is determined by an octet value this value shall be interpreted as
    /// follows:
    ///
    /// | Value     | Description                                   |
    /// |-----------|-----------------------------------------------|
    /// | 0x00      | DEFAULT PITCH + UNKNOWN FONT FAMILY           |
    /// | 0x01      | FIXED PITCH + UNKNOWN FONT FAMILY             |
    /// | 0x02      | VARIABLE PITCH + UNKNOWN FONT FAMILY          |
    /// | 0x10      | DEFAULT PITCH + ROMAN FONT FAMILY             |
    /// | 0x11      | FIXED PITCH + ROMAN FONT FAMILY               |
    /// | 0x12      | VARIABLE PITCH + ROMAN FONT FAMILY            |
    /// | 0x20      | DEFAULT PITCH + SWISS FONT FAMILY             |
    /// | 0x21      | FIXED PITCH + SWISS FONT FAMILY               |
    /// | 0x22      | VARIABLE PITCH + SWISS FONT FAMILY            |
    /// | 0x30      | DEFAULT PITCH + MODERN FONT FAMILY            |
    /// | 0x31      | FIXED PITCH + MODERN FONT FAMILY              |
    /// | 0x32      | VARIABLE PITCH + MODERN FONT FAMILY           |
    /// | 0x40      | DEFAULT PITCH + SCRIPT FONT FAMILY            |
    /// | 0x41      | FIXED PITCH + SCRIPT FONT FAMILY              |
    /// | 0x42      | VARIABLE PITCH + SCRIPT FONT FAMILY           |
    /// | 0x50      | DEFAULT PITCH + DECORATIVE FONT FAMILY        |
    /// | 0x51      | FIXED PITCH + DECORATIVE FONT FAMIL           |
    /// | 0x52      | VARIABLE PITCH + DECORATIVE FONT FAMILY       |
    ///
    /// This information is determined by querying the font when present and shall not be
    /// modified when the font is not available. This information can be used in font substitution
    /// logic to locate an appropriate substitute font when this font is not available.
    ///
    /// Defaults to 0x00
    ///
    /// # Note
    ///
    /// Although the attribute name is pitchFamily, the integer value of this attribute
    /// specifies the font family with higher 4 bits and the font pitch with lower 4 bits.
    pub pitch_family: Option<i32>,

    /// Specifies the character set which is supported by the parent font. This information can be
    /// used in font substitution logic to locate an appropriate substitute font when this font is
    /// not available. This information is determined by querying the font when present and shall
    /// not be modified when the font is not available.
    ///
    /// The value of this attribute shall be interpreted as follows:
    ///
    /// | Value     | Description                                                               |
    /// |-----------|---------------------------------------------------------------------------|
    /// | 0x00      | Specifies the ANSI character set. (IANA name *iso-8859-1*)                |
    /// | 0x01      | Specifies the default character set.                                      |
    /// | 0x02      | Specifies the Symbol character set. This value specifies that the         |
    /// |           | characters in the Unicode private use area (U+FF00 to U+FFFF) of the      |
    /// |           | font should be used to display characters in the range U+0000 to          |
    /// |           | U+00FF.                                                                   |
    /// | 0x4D      | Specifies a Macintosh (Standard Roman) character set. (IANA name          |
    /// |           | *macintosh*)                                                              |
    /// | 0x80      | Specifies the JIS character set. (IANA name *shift_jis*)                  |
    /// | 0x81      | Specifies the Hangul character set. (IANA name *ks_c_5601-1987*)          |
    /// | 0x82      | Specifies a Johab character set. (IANA name *KS C-5601-1992*)             |
    /// | 0x86      | Specifies the GB-2312 character set. (IANA name *GBK*)                    |
    /// | 0x88      | Specifies the Chinese Big Five character set. (IANA name *Big5*)          |
    /// | 0xA1      | Specifies a Greek character set. (IANA name *windows-1253*)               |
    /// | 0xA2      | Specifies a Turkish character set. (IANA name *iso-8859-9*)               |
    /// | 0xA3      | Specifies a Vietnamese character set. (IANA name *windows-1258*)          |
    /// | 0xB1      | Specifies a Hebrew character set. (IANA name *windows-1255*)              |
    /// | 0xB2      | Specifies an Arabic character set. (IANA name *windows-1256*)             |
    /// | 0xBA      | Specifies a Baltic character set. (IANA name *windows-1257*)              |
    /// | 0xCC      | Specifies a Russian character set. (IANA name *windows-1251*)             |
    /// | 0xDE      | Specifies a Thai character set. (IANA name *windows-874*)                 |
    /// | 0xEE      | Specifies an Eastern European character set. (IANA name *windows-1250*)   |
    /// | 0xFF      | Specifies an OEM character set not defined by ECMA-376.                   |
    /// | _         | Application-defined, can be ignored.                                      |
    ///
    /// Defaults to 0x01
    pub charset: Option<i32>,
}

impl TextFont {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextFont> {
        let mut typeface = None;
        let mut panose = None;
        let mut pitch_family = None;
        let mut charset = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "typeface" => typeface = Some(value.clone()),
                "panose" => panose = Some(value.clone()),
                "pitchFamily" => pitch_family = Some(value.parse::<i32>()?),
                "charset" => charset = Some(value.parse::<i32>()?),
                _ => (),
            }
        }

        let typeface = typeface.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "typeface"))?;

        Ok(Self {
            typeface,
            panose,
            pitch_family,
            charset,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextRun {
    /// This element specifies the presence of a run of text within the containing text body. The run element is the
    /// lowest level text separation mechanism within a text body. A text run can contain text run properties associated
    /// with the run. If no properties are listed then properties specified in the defRPr element are used.
    ///
    /// # Xml example
    ///
    /// Consider the case where the user would like to describe a text body that contains two runs of text and
    /// would like one to be bold and the other not. The following DrawingML would specify such a text body.
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:r>
    ///     <a:rPr b="1">
    ///     </a:rPr>
    ///     <a:t>Some text</a:t>
    ///   </a:r>
    ///   …
    ///   <a:r>
    ///     <a:rPr/>
    ///     <a:t>Some text</a:t>
    ///   </a:r>
    /// </p:txBody>
    /// ```
    ///
    /// The above text body has the first run be formatted bold and the second normally.
    RegularTextRun(Box<RegularTextRun>),

    /// This element specifies the existence of a vertical line break between two runs of text within a paragraph. In
    /// addition to specifying a vertical space between two runs of text, this element can also have run properties
    /// specified via the rPr child element. This sets the formatting of text for the line break so that if text is later
    /// inserted there that a new run can be generated with the correct formatting.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       …
    ///       <a:t>Text Run 1.</a:t>
    ///       …
    ///     </a:r>
    ///     <a:br/>
    ///       <a:r>
    ///       …
    ///       <a:t>Text Run 2.</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// This paragraph has two runs of text laid out in a vertical fashion with a line break in between them. This line
    /// break acts much like a carriage return would within a normal run of text.
    LineBreak(Box<TextLineBreak>),

    /// This element specifies a text field which contains generated text that the application should update periodically.
    /// Each piece of text when it is generated is given a unique identification number that is used to refer to a specific
    /// field. At the time of creation the text field indicates the kind of text that should be used to update this field. This
    /// update type is used so that all applications that did not create this text field can still know what kind of text it
    /// should be updated with. Thus the new application can then attach an update type to the text field id for
    /// continual updating.
    ///
    /// # Xml example
    ///
    /// Consider a slide within a presentation that needs to have the slide number placed on the slide. The
    /// following DrawingML can be used to describe such a situation.
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr/>
    ///   <a:lstStyle/>
    ///   <a:p>
    ///     <a:fld id="{424CEEAC-8F67-4238-9622-1B74DC6E8318}" type="slidenum">
    ///       <a:rPr lang="en-US" smtClean="0"/>
    ///       <a:pPr/>
    ///       <a:t>3</a:t>
    ///     </a:fld>
    ///     <a:endParaRPr lang="en-US"/>
    ///   </a:p>
    /// </p:txBody>
    /// ```
    TextField(Box<TextField>),
}

impl XsdType for TextRun {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "r" => Ok(TextRun::RegularTextRun(Box::new(RegularTextRun::from_xml_element(
                xml_node,
            )?))),
            "br" => Ok(TextRun::LineBreak(Box::new(TextLineBreak::from_xml_element(xml_node)?))),
            "fld" => Ok(TextRun::TextField(Box::new(TextField::from_xml_element(xml_node)?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextRun").into()),
        }
    }
}

impl XsdChoice for TextRun {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "r" | "br" | "fld" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RegularTextRun {
    /// This element contains all run level text properties for the text runs within a containing paragraph.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:p>
    ///   …
    ///   <a:rPr u="sng"/>
    ///   …
    ///   <a:t>Some Text</a:t>
    ///   …
    /// </a:p>
    /// ```
    ///
    /// The run of text described above is formatting with a single underline of text matching color.
    pub char_properties: Option<Box<TextCharacterProperties>>,

    /// This element specifies the actual text for this text run. This is the text that is formatted using all specified body,
    /// paragraph and run properties. This element shall be present within a run of text.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     …
    ///     <a:r>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above DrawingML specifies a text body containing a single paragraph, containing a single run which contains
    /// the actual text specified with the <a:t> element.
    pub text: String,
}

impl RegularTextRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut char_properties = None;
        let mut text = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => char_properties = Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?)),
                "t" => text = child_node.text.clone(),
                _ => (),
            }
        }

        let text = text.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "t"))?;
        Ok(Self { char_properties, text })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextUnderlineLine {
    /// This element specifies that the stroke style of an underline for a run of text should be of the same as the text run
    /// within which it is contained.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:uLnTx>
    ///       </a:rPr>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The underline stroke of the above text follows the stroke of the run text within which it resides.
    FollowText,

    /// This element specifies the properties for the stroke of the underline that is present within a run of text.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:uLn algn="r">
    ///       </a:rPr>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    Line(Box<LineProperties>),
}

impl XsdType for TextUnderlineLine {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "uLnTx" => Ok(TextUnderlineLine::FollowText),
            "uLn" => Ok(TextUnderlineLine::Line(Box::new(LineProperties::from_xml_element(
                xml_node,
            )?))),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextUnderlineLine").into()),
        }
    }
}

impl XsdChoice for TextUnderlineLine {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "uLnTx" | "uLn" => true,
            _ => false,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextUnderlineFill {
    /// This element specifies that the fill color of an underline for a run of text should be of the same color as the text
    /// run within which it is contained.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:uFillTx>
    ///       </a:rPr>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    FollowText,

    /// This element specifies the fill color of an underline for a run of text.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:uFill>
    ///           <a:solidFill>
    ///             <a:srgbClr val="FFFF00"/>
    ///           </a:solidFill>
    ///         </a:uFill>
    ///       </a:rPr>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    ///
    Fill(FillProperties),
}

impl TextUnderlineFill {
    pub fn is_choice_member(name: &str) -> bool {
        match name {
            "uFillTx" | "uFill" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "uFillTx" => Ok(TextUnderlineFill::FollowText),
            "uFill" => {
                let fill_properties = xml_node
                    .child_nodes
                    .iter()
                    .find_map(FillProperties::try_from_xml_element)
                    .transpose()?
                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "EG_FillProperties"))?;

                Ok(TextUnderlineFill::Fill(fill_properties))
            }
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextUnderlineFill").into()),
        }
    }
}
