use super::{
    bullet::{TextBullet, TextBulletColor, TextBulletSize, TextBulletTypeface},
    runformatting::{TextFont, TextRun, TextUnderlineFill, TextUnderlineLine},
};
use crate::{
    shared::drawingml::{
        colors::Color,
        core::{Hyperlink, LineProperties},
        shapeprops::{EffectProperties, FillProperties},
        simpletypes::{
            Coordinate32, Guid, Percentage, TextAlignType, TextCapsType, TextFontAlignType, TextFontSize, TextIndent,
            TextIndentLevelType, TextLanguageID, TextMargin, TextNonNegativePoint, TextPoint, TextSpacingPercent,
            TextSpacingPoint, TextStrikeType, TextTabAlignType, TextUnderlineType,
        },
        util::XmlNodeExt,
    },
    error::{LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};
use std::error::Error;

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextLineBreak {
    pub char_properties: Option<Box<TextCharacterProperties>>,
}

impl TextLineBreak {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let char_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .map(TextCharacterProperties::from_xml_element)
            .transpose()?
            .map(Box::new);

        Ok(Self { char_properties })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextField {
    /// Specifies the unique to this document, host specified token that is used to identify the
    /// field. This token is generated when the text field is created and persists in the file as the
    /// same token until the text field is removed. Any application should check the document
    /// for conflicting tokens before assigning a new token to a text field.
    pub id: Guid,

    /// Specifies the type of text that should be used to update this text field. This is used to
    /// inform the rendering application what text it should use to update this text field. There
    /// are no specific syntax restrictions placed on this attribute. The generating application can
    /// use it to represent any text that should be updated before rendering the presentation.
    ///
    /// Reserved values:
    ///
    /// |Value          |Description                                            |
    /// |---------------|-------------------------------------------------------|
    /// |slidenum       |presentation slide number                              |
    /// |datetime       |default date time format for the rendering application |
    /// |datetime1      |MM/DD/YYYY date time format                            |
    /// |datetime2      |Day, Month DD, YYYY date time format                   |
    /// |datetime3      |DD Month YYYY date time format                         |
    /// |datetime4      |Month DD, YYYY date time format                        |
    /// |datetime5      |DD-Mon-YY date time format                             |
    /// |datetime6      |Month YY date time format                              |
    /// |datetime7      |Mon-YY date time format                                |
    /// |datetime8      |MM/DD/YYYY hh:mm AM/PM date time format                |
    /// |datetime9      |MM/DD/YYYY hh:mm:ss AM/PM date time format             |
    /// |datetime10     |hh:mm date time format                                 |
    /// |datetime11     |hh:mm:ss date time format                              |
    /// |datetime12     |hh:mm AM/PM date time format                           |
    /// |datetime13     |hh:mm:ss: AM/PM date time format                       |
    pub field_type: Option<String>,

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

    /// Specifies the paragraph properties for this text field
    pub paragraph_properties: Option<Box<TextParagraph>>,

    /// The text of this text field.
    pub text: Option<String>,
}

impl TextField {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut field_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.clone()),
                "type" => field_type = Some(value.clone()),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        let mut char_properties = None;
        let mut paragraph_properties = None;
        let mut text = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => char_properties = Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?)),
                "pPr" => paragraph_properties = Some(Box::new(TextParagraph::from_xml_element(child_node)?)),
                "t" => text = child_node.text.clone(),
                _ => (),
            }
        }

        Ok(Self {
            id,
            field_type,
            char_properties,
            paragraph_properties,
            text,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextParagraphProperties {
    /// Specifies the left margin of the paragraph. This is specified in addition to the text body
    /// inset and applies only to this text paragraph. That is the text body inset and the marL
    /// attributes are additive with respect to the text position. If this attribute is omitted, then a
    /// value of 347663 is implied.
    pub margin_left: Option<TextMargin>,

    /// Specifies the right margin of the paragraph. This is specified in addition to the text body
    /// inset and applies only to this text paragraph. That is the text body inset and the marR
    /// attributes are additive with respect to the text position. If this attribute is omitted, then a
    /// value of 0 is implied.
    pub margin_right: Option<TextMargin>,

    /// Specifies the particular level text properties that this paragraph follows. The value for this
    /// attribute is numerical and formats the text according to the corresponding level
    /// paragraph properties that are listed within the lstStyle element. Since there are nine
    /// separate level properties defined, this tag has an effective range of 0-8 = 9 available
    /// values.
    ///
    /// # Xml example
    ///
    /// Consider the following DrawingML. This would specify that this paragraph
    /// should follow the lvl2pPr formatting style because once again lvl="1" is considered to be
    /// level 2.
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr lvl="1" …/>
    ///     …
    ///     <a:t>Sample text</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// # Note
    ///
    /// To resolve conflicting paragraph properties the linear hierarchy of paragraph
    /// properties should be examined starting first with the pPr element. The rule here is that
    /// properties that are defined at a level closer to the actual text should take precedence.
    /// That is if there is a conflicting property between the pPr and lvl1pPr elements then the
    /// pPr property should take precedence because in the property hierarchy it is closer to the
    /// actual text being represented.
    pub level: Option<TextIndentLevelType>,

    /// Specifies the indent size that is applied to the first line of text in the paragraph. An
    /// indentation of 0 is considered to be at the same location as marL attribute. If this
    /// attribute is omitted, then a value of -342900 is implied.
    ///
    /// # Xml example
    ///
    /// Consider the scenario where the user now wanted to add a paragraph
    /// indentation to the first line of text in their two column format book.
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr numCol="2" spcCol="914400"…/>
    ///     <a:normAutofit/>
    ///   </a:bodyPr>
    ///   …
    ///   <a:p>
    ///     <a:pPr marL="0" indent="571500" algn="just">
    ///       <a:buNone/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Here is some…</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// By adding the indent attribute the user has effectively added a first line indent to this
    /// paragraph of text.
    pub indent: Option<TextIndent>,

    /// Specifies the alignment that is to be applied to the paragraph. Possible values for this
    /// include left, right, centered, justified and distributed. If this attribute is omitted, then a
    /// value of left is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where the user wishes to have two columns of text that
    /// have a justified alignment, much like text within a book. The following DrawingML could
    /// describe this.
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr numCol="2" spcCol="914400"…/>
    ///     <a:normAutofit/>
    ///   </a:bodyPr>
    ///   …
    ///   <a:p>
    ///     <a:pPr marL="0" algn="just">
    ///       <a:buNone/>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Sample Text …</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub align: Option<TextAlignType>,

    /// Specifies the default size for a tab character within this paragraph. This attribute should
    /// be used to describe the spacing of tabs within the paragraph instead of a leading
    /// indentation tab. For indentation tabs there are the marL and indent attributes to assist
    /// with this.
    ///
    /// # Xml example
    ///
    /// Consider the case where a paragraph contains numerous tabs that need to be
    /// of a specific size. The following DrawingML would describe this.
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr defTabSz="376300" …/>
    ///     …
    ///     <a:t>Sample Text …</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub default_tab_size: Option<Coordinate32>,

    /// Specifies whether the text is right-to-left or left-to-right in its flow direction. If this
    /// attribute is omitted, then a value of 0, or left-to-right is implied.
    ///
    /// # Xml example
    ///
    /// Consider the following example of a text body with two lines of text. In this
    /// example, both lines contain English and Arabic text, however, the second line has the
    /// rtl attribute set to true whereas the first line does not set the rtl attribute.
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:t>Test </a:t>
    ///     </a:r>
    ///     <a:r>
    ///       <a:rPr>
    ///         <a:rtl w:val="1"/>
    ///       </a:rPr>
    ///       <a:t> تجربة </a:t>
    ///     </a:r>
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr rtl="1"/>
    ///     <a:r>
    ///       <a:rPr>
    ///         <a:rtl w:val="0"/>
    ///       </a:rPr>
    ///       <a:t>Test </a:t>
    ///     </a:r>
    ///     <a:r>
    ///       <a:t> تجربة </a:t>
    ///     </a:r>
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub rtl: Option<bool>,

    /// Specifies whether an East Asian word can be broken in half and wrapped onto the next
    /// line without a hyphen being added. To determine whether an East Asian word can be
    /// broken the presentation application would use the kinsoku settings here. This attribute
    /// is to be used specifically when there is a word that cannot be broken into multiple pieces
    /// without a hyphen. That is it is not present within the existence of normal breakable East
    /// Asian words but is when a special case word arises that should not be broken for a line
    /// break. If this attribute is omitted, then a value of 1 or true is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where the presentation contains a long word that must not
    /// be divided with a line break. Instead it should be placed, in whole on a new line so that it
    /// can fit. The picture below shows a normal paragraph where a long word has been broken
    /// for a line break. The second picture shown below shows that same paragraph with the
    /// long word specified to not allow a line break. The resulting DrawingML is as follows.
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr eaLnBrk="0" …/>
    ///     …
    ///     <a:t>Sample text (Long word)</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub east_asian_line_break: Option<bool>,

    /// Determines where vertically on a line of text the actual words are positioned. This deals
    /// with vertical placement of the characters with respect to the baselines. For instance
    /// having text anchored to the top baseline, anchored to the bottom baseline, centered in
    /// between, etc. To understand this attribute and it's use it is helpful to understand what
    /// baselines are. A diagram describing these different cases is shown below. If this attribute
    /// is omitted, then a value of base is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where the user wishes to represent the chemical compound
    /// of a water molecule. For this they need to make sure the H, the 2, and the O are all in the
    /// correct position and are of the correct size. The results below can be achieved through
    /// the DrawingML shown below.
    ///
    /// ```xml
    /// <a:txtBody>
    ///   …
    ///   <a:pPr fontAlgn="b" …/>
    ///   …
    ///   <a:r>
    ///     <a:rPr …/>
    ///     <a:t>H </a:t>
    ///   </a:r>
    ///   <a:r>
    ///     <a:rPr sz="1200" …/>
    ///     <a:t>2</a:t>
    ///   </a:r>
    ///   <a:r>
    ///     <a:rPr …/>
    ///     <a:t>O</a:t>
    ///   </a:r>
    ///   …
    /// </p:txBody>
    /// ```
    pub font_align: Option<TextFontAlignType>,

    /// Specifies whether a Latin word can be broken in half and wrapped onto the next line
    /// without a hyphen being added. This attribute is to be used specifically when there is a
    /// word that cannot be broken into multiple pieces without a hyphen. It is not present
    /// within the existence of normal breakable Latin words but is when a special case word
    /// arises that should not be broken for a line break. If this attribute is omitted, then a value
    /// of 1 or true is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where the presentation contains a long word that must not
    /// be divided with a line break. Instead it should be placed, in whole on a new line so that it
    /// can fit. The picture below shows a normal paragraph where a long word has been broken
    /// for a line break. The second picture shown below shows that same paragraph with the
    /// long word specified to not allow a line break. The resulting DrawingML is as follows.
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr latinLnBrk="0" …/>
    ///     …
    ///     <a:t>Sample text (Long word)</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub latin_line_break: Option<bool>,

    /// Specifies whether punctuation is to be forcefully laid out on a line of text or put on a
    /// different line of text. That is, if there is punctuation at the end of a run of text that should
    /// be carried over to a separate line does it actually get carried over. A true value allows for
    /// hanging punctuation forcing the punctuation to not be carried over and a value of false
    /// allows the punctuation to be carried onto the next text line. If this attribute is omitted,
    /// then a value of 0, or false is implied.
    pub hanging_punctuations: Option<bool>,

    /// This element specifies the vertical line spacing that is to be used within a paragraph. This can be specified in two
    /// different ways, percentage spacing and font point spacing. If this element is omitted then the spacing between
    /// two lines of text should be determined by the point size of the largest piece of text within a line.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:p>
    ///     <a:pPr>
    ///       <a:lnSpc>
    ///         <a:spcPct val="200%"/>
    ///       </a:lnSpc>
    ///     </a:pPr>
    ///     <a:r>
    ///       <a:rPr lang="en-US" dirty="0" smtClean="0"/>
    ///       <a:t>Some</a:t>
    ///     </a:r>
    ///     <a:br>
    ///       <a:rPr lang="en-US" smtClean="0"/>
    ///     </a:br>
    ///     <a:r>
    ///      <a:rPr lang="en-US" dirty="0" smtClean="0"/>
    ///      <a:t>Text</a:t>
    ///     </a:r>
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// This paragraph has two lines of text that have percentage based vertical spacing. This kind of spacing should
    /// change based on the size of the text involved as its size is calculated as a percentage of this.
    pub line_spacing: Option<TextSpacing>,

    /// This element specifies the amount of vertical white space that is present before a paragraph. This space is
    /// specified in either percentage or points via the child elements spcPct and spcPts.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:spcBef>
    ///         <a:spcPts val="1800"/>
    ///       </a:spcBef>
    ///       <a:spcAft>
    ///         <a:spcPts val="600"/>
    ///       </a:spcAft>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Sample Text</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph of text is formatted to have a spacing both before and after the paragraph text. The
    /// spacing before is a size of 18 points, or value=1800 and the spacing after is a size of 6 points, or value=600.
    pub space_before: Option<TextSpacing>,

    /// This element specifies the amount of vertical white space that is present after a paragraph. This space is
    /// specified in either percentage or points via the child elements spcPct and spcPts.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:spcBef>
    ///         <a:spcPts val="1800"/>
    ///       </a:spcBef>
    ///       <a:spcAft>
    ///         <a:spcPts val="600"/>
    ///       </a:spcAft>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Sample Text</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph of text is formatted to have a spacing both before and after the paragraph text. The
    /// spacing before is a size of 18 points, or value=1800 and the spacing after is a size of 6 points, or value=600.
    pub space_after: Option<TextSpacing>,

    /// Specifies the color of the bullet for this paragraph.
    pub bullet_color: Option<TextBulletColor>,

    /// Specifies the size of the bullet for this paragraph.
    pub bullet_size: Option<TextBulletSize>,

    /// Specifies the font properties of the bullet for this paragraph.
    pub bullet_typeface: Option<TextBulletTypeface>,

    /// Specifies the bullet's properties for this paragraph.
    pub bullet: Option<TextBullet>,

    /// This element specifies the list of all tab stops that are to be used within a paragraph. These tabs should be used
    /// when describing any custom tab stops within the document. If these are not specified then the default tab stops
    /// of the generating application should be used.
    pub tab_stop_list: Option<Vec<TextTabStop>>,

    /// This element contains all default run level text properties for the text runs within a containing paragraph. These
    /// properties are to be used when overriding properties have not been defined within the rPr element.
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
    pub default_run_properties: Option<Box<TextCharacterProperties>>,
}

impl TextParagraphProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextParagraphProperties> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "marL" => instance.margin_left = Some(value.parse()?),
                    "marR" => instance.margin_right = Some(value.parse()?),
                    "lvl" => instance.level = Some(value.parse()?),
                    "indent" => instance.indent = Some(value.parse()?),
                    "algn" => instance.align = Some(value.parse()?),
                    "defTabSz" => instance.default_tab_size = Some(value.parse()?),
                    "rtl" => instance.rtl = Some(parse_xml_bool(value)?),
                    "eaLnBrk" => instance.east_asian_line_break = Some(parse_xml_bool(value)?),
                    "fontAlgn" => instance.font_align = Some(value.parse()?),
                    "latinLnBrk" => instance.latin_line_break = Some(parse_xml_bool(value)?),
                    "hangingPunct" => instance.hanging_punctuations = Some(parse_xml_bool(value)?),
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
                            "lnSpc" => {
                                instance.line_spacing = Some(
                                    child_node
                                        .child_nodes
                                        .iter()
                                        .find_map(TextSpacing::try_from_xml_element)
                                        .transpose()?
                                        .ok_or_else(|| {
                                            MissingChildNodeError::new(child_node.name.clone(), "EG_TextSpacing")
                                        })?,
                                );
                            }
                            "spcBef" => {
                                instance.space_before = Some(
                                    child_node
                                        .child_nodes
                                        .iter()
                                        .find_map(TextSpacing::try_from_xml_element)
                                        .transpose()?
                                        .ok_or_else(|| {
                                            MissingChildNodeError::new(child_node.name.clone(), "EG_TextSpacing")
                                        })?,
                                );
                            }
                            "spcAft" => {
                                instance.space_after = Some(
                                    child_node
                                        .child_nodes
                                        .iter()
                                        .find_map(TextSpacing::try_from_xml_element)
                                        .transpose()?
                                        .ok_or_else(|| {
                                            MissingChildNodeError::new(child_node.name.clone(), "EG_TextSpacing")
                                        })?,
                                );
                            }
                            "tabLst" => {
                                let vec = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|tab_stop_node| tab_stop_node.local_name() == "tab")
                                    .map(TextTabStop::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;

                                instance.tab_stop_list = match vec.len() {
                                    len if len <= 32 => Some(vec),
                                    len => {
                                        return Err(Box::<dyn Error>::from(LimitViolationError::new(
                                            xml_node.name.clone(),
                                            "tabLst",
                                            0,
                                            MaxOccurs::Value(32),
                                            len as u32,
                                        )))
                                    }
                                };
                            }
                            "defRPr" => {
                                instance.default_run_properties =
                                    Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?))
                            }
                            local_name if TextBulletColor::is_choice_member(local_name) => {
                                instance.bullet_color = Some(TextBulletColor::from_xml_element(child_node)?);
                            }
                            local_name if TextBulletSize::is_choice_member(local_name) => {
                                instance.bullet_size = Some(TextBulletSize::from_xml_element(child_node)?);
                            }
                            local_name if TextBulletTypeface::is_choice_member(local_name) => {
                                instance.bullet_typeface = Some(TextBulletTypeface::from_xml_element(child_node)?);
                            }
                            local_name if TextBulletTypeface::is_choice_member(local_name) => {
                                instance.bullet = Some(TextBullet::from_xml_element(child_node)?);
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextParagraph {
    /// This element contains all paragraph level text properties for the containing paragraph. These paragraph
    /// properties should override any and all conflicting properties that are associated with the paragraph in question.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:p>
    ///   <a:pPr marL="0" algn="ctr">
    ///     <a:buNone/>
    ///   </a:pPr>
    ///   …
    ///   <a:t>Some Text</a:t>
    ///   …
    /// </a:p>
    /// ```
    ///
    /// The paragraph described above is formatting with a left margin of 0 and has all of text runs contained within it
    /// centered about the horizontal median of the bounding box for the text body.
    ///
    /// # Note
    ///
    /// To resolve conflicting paragraph properties the linear hierarchy of paragraph properties should be
    /// examined starting first with the pPr element. The rule here is that properties that are defined at a level closer to
    /// the actual text should take precedence. That is if there is a conflicting property between the pPr and lvl1pPr
    /// elements then the pPr property should take precedence because in the property hierarchy it is closer to the
    /// actual text being represented.
    pub properties: Option<Box<TextParagraphProperties>>,

    /// The list of text runs in this paragraph.
    pub text_run_list: Vec<TextRun>,

    /// This element specifies the text run properties that are to be used if another run is inserted after the last run
    /// specified. This effectively saves the run property state so that it can be applied when the user enters additional
    /// text. If this element is omitted, then the application can determine which default properties to apply. It is
    /// recommended that this element be specified at the end of the list of text runs within the paragraph so that an
    /// orderly list is maintained.
    pub end_paragraph_char_properties: Option<Box<TextCharacterProperties>>,
}

impl TextParagraph {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "pPr" => {
                        instance.properties = Some(Box::new(TextParagraphProperties::from_xml_element(child_node)?))
                    }
                    "endParaRPr" => {
                        instance.end_paragraph_char_properties =
                            Some(Box::new(TextCharacterProperties::from_xml_element(child_node)?))
                    }
                    local_name if TextRun::is_choice_member(local_name) => {
                        instance.text_run_list.push(TextRun::from_xml_element(child_node)?);
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextCharacterProperties {
    /// Specifies whether the numbers contained within vertical text continue vertically with the
    /// text or whether they are to be displayed horizontally while the surrounding characters
    /// continue in a vertical fashion. If this attribute is omitted, than a value of 0, or false is
    /// assumed.
    pub kumimoji: Option<bool>,

    /// Specifies the language to be used when the generating application is displaying the user
    /// interface controls. If this attribute is omitted, than the generating application can select a
    /// language of its choice.
    pub language: Option<TextLanguageID>,

    /// Specifies the alternate language to use when the generating application is displaying the
    /// user interface controls. If this attribute is omitted, than the lang attribute is used here.
    pub alternative_language: Option<TextLanguageID>,

    /// Specifies the size of text within a text run. Whole points are specified in increments of
    /// 100 starting with 100 being a point size of 1. For instance a font point size of 12 would be
    /// 1200 and a font point size of 12.5 would be 1250. If this attribute is omitted, than the
    /// value in defRPr should be used.
    pub font_size: Option<TextFontSize>,

    /// Specifies whether a run of text is formatted as bold text. If this attribute is omitted, than
    /// a value of 0, or false is assumed.
    pub bold: Option<bool>,

    /// Specifies whether a run of text is formatted as italic text. If this attribute is omitted, than
    /// a value of 0, or false is assumed.
    pub italic: Option<bool>,

    /// Specifies whether a run of text is formatted as underlined text. If this attribute is omitted,
    /// than no underline is assumed.
    pub underline: Option<TextUnderlineType>,

    /// Specifies whether a run of text is formatted as strikethrough text. If this attribute is
    /// omitted, than no strikethrough is assumed.
    pub strikethrough: Option<TextStrikeType>,

    /// Specifies the minimum font size at which character kerning occurs for this text run.
    /// Whole points are specified in increments of 100 starting with 100 being a point size of 1.
    /// For instance a font point size of 12 would be 1200 and a font point size of 12.5 would be
    /// 1250. If this attribute is omitted, than kerning occurs for all font sizes down to a 0 point
    /// font.
    pub kerning: Option<TextNonNegativePoint>,

    /// Specifies the capitalization that is to be applied to the text run. This is a render-only
    /// modification and does not affect the actual characters stored in the text run. This
    /// attribute is also distinct from the toggle function where the actual characters stored in
    /// the text run are changed.
    pub capitalization: Option<TextCapsType>,

    /// Specifies the spacing between characters within a text run. This spacing is specified
    /// numerically and should be consistently applied across the entire run of text by the
    /// generating application. Whole points are specified in increments of 100 starting with 100
    /// being a point size of 1. For instance a font point size of 12 would be 1200 and a font point
    /// size of 12.5 would be 1250. If this attribute is omitted than a value of 0 or no adjustment
    /// is assumed.
    pub spacing: Option<TextPoint>,

    /// Specifies the normalization of height that is to be applied to the text run. This is a renderonly
    /// modification and does not affect the actual characters stored in the text run. This
    /// attribute is also distinct from the toggle function where the actual characters stored in
    /// the text run are changed. If this attribute is omitted, than a value of 0, or false is
    /// assumed.
    pub normalize_heights: Option<bool>,

    /// Specifies the baseline for both the superscript and subscript fonts. The size is specified
    /// using a percentage where 1% is equal to 1 percent of the font size and 100% is equal to
    /// 100 percent font of the font size.
    pub baseline: Option<Percentage>,

    /// Specifies that a run of text has been selected by the user to not be checked for mistakes.
    /// Therefore if there are spelling, grammar, etc mistakes within this text the generating
    /// application should ignore them.
    pub no_proofing: Option<bool>,

    /// Specifies that the content of a text run has changed since the proofing tools have last
    /// been run. Effectively this flags text that is to be checked again by the generating
    /// application for mistakes such as spelling, grammar, etc.
    ///
    /// Defaults to true
    pub dirty: Option<bool>,

    /// Specifies that when this run of text was checked for spelling, grammar, etc. that a
    /// mistake was indeed found. This allows the generating application to effectively save the
    /// state of the mistakes within the document instead of having to perform a full pass check
    /// upon opening the document.
    ///
    /// Defaults to false
    pub spelling_error: Option<bool>,

    /// Specifies whether or not a text run has been checked for smart tags. This attribute acts
    /// much like the dirty attribute dose for the checking of spelling, grammar, etc. A value of
    /// true here indicates to the generating application that this text run should be checked for
    /// smart tags. If this attribute is omitted, than a value of 0, or false is assumed.
    pub smarttag_clean: Option<bool>,

    /// Specifies a smart tag identifier for a run of text. This ID is unique throughout the
    /// presentation and is used to reference corresponding auxiliary information about the
    /// smart tag.
    ///
    /// Defaults to 0
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr/>
    ///   <a:lstStyle/>
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr lang="en-US" dirty="0" smtId="1"/>
    ///       <a:t>CNTS</a:t>
    ///     </a:r>
    ///     <a:endParaRPr lang="en-US" dirty="0"/>
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// The text run has a smtId attribute value of 1, which denotes that the text should be
    /// inspected for smart tag information, which in this case maps to a stock ticker symbol.
    pub smarttag_id: Option<u32>,

    /// Specifies the link target name that is used to reference to the proper link properties in a
    /// custom XML part within the document.
    pub bookmark_link_target: Option<String>,

    /// Specifies the outline properties of this character
    pub line_properties: Option<Box<LineProperties>>,

    /// Specifies the fill properties of this character
    pub fill_properties: Option<FillProperties>,

    /// Specifies the effect that should be applied to this character
    pub effect_properties: Option<EffectProperties>,

    /// This element specifies the highlight color that is present for a run of text.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:highlight>
    ///           <a:srgbClr val="FFFF00"/>
    ///         </a:highlight>
    ///       </a:rPr>
    ///       …
    ///       <a:t>Sample Text</a:t>
    ///       …
    ///     </a:r>
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    pub highlight_color: Option<Color>,

    /// Specifies the line properties of the underline for this character.
    pub text_underline_line: Option<TextUnderlineLine>,

    /// Specifies the fill properties of the underline for this character.
    pub text_underline_fill: Option<TextUnderlineFill>,

    /// This element specifies that a Latin font be used for a specific run of text. This font is specified with a typeface
    /// attribute much like the others but is specifically classified as a Latin font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:r>
    ///   <a:rPr …>
    ///     <a:latin typeface="Sample Font"/>
    ///   </a:rPr>
    ///   <a:t>Sample Text</a:t>
    /// </a:r>
    /// ```
    pub latin_font: Option<TextFont>,

    /// This element specifies that an East Asian font be used for a specific run of text. This font is specified with a
    /// typeface attribute much like the others but is specifically classified as an East Asian font.
    ///
    /// If the specified font is not available on a system being used for rendering, then the attributes of this element can
    /// be utilized to select an alternative font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:r>
    ///   <a:rPr …>
    ///     <a:ea typeface="Sample Font"/>
    ///   </a:rPr>
    ///   <a:t>Sample Text</a:t>
    /// </a:r>
    /// ```
    pub east_asian_font: Option<TextFont>,

    /// This element specifies that a complex script font be used for a specific run of text. This font is specified with a
    /// typeface attribute much like the others but is specifically classified as a complex script font.
    ///
    /// If the specified font is not available on a system being used for rendering, then the attributes of this element can
    /// be utilized to select an alternative font.
    pub complex_script_font: Option<TextFont>,

    /// This element specifies that a symbol font be used for a specific run of text. This font is specified with a typeface
    /// attribute much like the others but is specifically classified as a symbol font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:r>
    ///   <a:rPr …>
    ///     <a:sym typeface="Sample Font"/>
    ///   </a:rPr>
    ///   <a:t>Sample Text</a:t>
    /// </a:r>
    /// ```
    ///
    /// The above run of text is rendered using the symbol font "Sample Font".
    pub symbol_font: Option<TextFont>,

    /// Specifies the on-click hyperlink information to be applied to a run of text. When the hyperlink text is clicked the
    /// link is fetched.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:hlinkClick r:id="rId2" tooltip="Some Sample Text"/>
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
    /// The above run of text is a hyperlink that points to the resource pointed at by rId2 within this slides relationship
    /// file. Additionally this text should display a tooltip when the mouse is hovered over the run of text.
    pub hyperlink_click: Option<Box<Hyperlink>>,

    /// Specifies the mouse-over hyperlink information to be applied to a run of text. When the mouse is hovered over
    /// this hyperlink text the link is fetched.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:r>
    ///       <a:rPr …>
    ///         <a:hlinkMouseOver r:id="rId2" tooltip="Some Sample Text"/>
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
    /// The above run of text is a hyperlink that points to the resource pointed at by rId2 within this slides relationship
    /// file. Additionally this text should display a tooltip when the mouse is hovered over the run of text.
    pub hyperlink_mouse_over: Option<Box<Hyperlink>>,

    /// This element specifies whether the contents of this run shall have right-to-left characteristics. Specifically, the
    /// following behaviors are applied when this element’s val attribute is true (or an equivalent):
    ///
    /// * Formatting – When the contents of this run are displayed, all characters shall be treated as complex
    ///   script characters. This means that the values of the cs element (§21.1.2.3.1) shall be used to determine
    ///   the font face.
    ///
    /// * Character Directionality Override – When the contents of this run are displayed, this property acts as a
    ///   right-to-left override for characters which are classified as follows (using the Unicode Character
    ///   Database):
    ///
    ///   * Weak types except European Number, European Number Terminator, Common Number Separator,
    ///     Arabic Number and (for Hebrew text) European Number Separator when constituting part of a
    ///     number
    ///
    ///   * Neutral types
    ///
    /// * This element provides information used to resolve the (Unicode) classifications of individual characters
    ///   as either L, R, AN or EN. Once this is determined, the line should be displayed subject to the
    ///   recommendation of the Unicode Bidirectional Algorithm in reordering resolved levels.
    ///
    ///   # Rationale
    ///
    ///   This override allows applications to store and utilize higher-level information beyond that
    ///   implicitly derived from the Unicode Bidirectional algorithm. For example, if the string “first second”
    ///   appears in a right-to-left paragraph inside a document, the Unicode algorithm would always result in
    ///   “first second” at display time (since the neutral character is surrounded by strongly classified
    ///   characters). However, if the whitespace was entered using a right-to-left input method (e.g. a Hebrew
    ///   keyboard), then that character could be classified as RTL using this property, allowing the display of
    ///   “second first” in a right-to-left paragraph, since the user explicitly asked for the space in a right-to-left
    ///   context.
    ///
    /// This property shall not be used with strong left-to-right text. Any behavior under that condition is unspecified.
    /// This property, when off, should not be used with strong right-to-left text. Any behavior under that condition is
    /// unspecified.
    ///
    /// If this element is not present, the default value is to leave the formatting applied at previous level in the style
    /// hierarchy. If this element is never applied in the style hierarchy, then right to left characteristics shall not be
    /// applied to the contents of this run.
    ///
    /// # Xml example
    ///
    /// Consider the following DrawingML visual content: “first second, أولى ثاني ”. This content might
    /// appear as follows within its parent paragraph:
    /// ```xml
    /// <a:p>
    ///   <a:r>
    ///     <a:t>first second, </w:t>
    ///   </a:r>
    ///   <a:r>
    ///     <a:rPr>
    ///       <a:rtl/>
    ///     </a:rPr>
    ///     <a:t> أولى </a:t>
    ///   </a:r>
    ///   <a:r>
    ///     <a:rPr>
    ///       <a:rtl/>
    ///     </a:rPr>
    ///     <a:t> </a:t>
    ///   </a:r>
    ///   <a:r>
    ///     <a:rPr>
    ///       <a:rtl/>
    ///     </a:rPr>
    ///     <a:t> ثاني </a:t>
    ///   </a:r>
    /// </a:p>
    /// ```
    ///
    /// The presence of the rtl element on the second, third, and fourth runs specifies that:
    ///
    /// * The formatting on those runs is specified using the complex-script property variants.
    /// * The whitespace character is treated as right-to-left.
    ///
    /// Note that the second, third and fourth runs could be joined as one run with the rtl element specified.
    pub rtl: Option<bool>,
}

impl TextCharacterProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextCharacterProperties> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "kumimoji" => instance.kumimoji = Some(parse_xml_bool(value)?),
                    "lang" => instance.language = Some(value.clone()),
                    "altLang" => instance.alternative_language = Some(value.clone()),
                    "sz" => instance.font_size = Some(value.parse()?),
                    "b" => instance.bold = Some(parse_xml_bool(value)?),
                    "i" => instance.italic = Some(parse_xml_bool(value)?),
                    "u" => instance.underline = Some(value.parse()?),
                    "strike" => instance.strikethrough = Some(value.parse()?),
                    "kern" => instance.kerning = Some(value.parse()?),
                    "cap" => instance.capitalization = Some(value.parse()?),
                    "spc" => instance.spacing = Some(value.parse()?),
                    "normalizeH" => instance.normalize_heights = Some(parse_xml_bool(value)?),
                    "baseline" => instance.baseline = Some(value.parse()?),
                    "noProof" => instance.no_proofing = Some(parse_xml_bool(value)?),
                    "dirty" => instance.dirty = Some(parse_xml_bool(value)?),
                    "err" => instance.spelling_error = Some(parse_xml_bool(value)?),
                    "smtClean" => instance.smarttag_clean = Some(parse_xml_bool(value)?),
                    "smtId" => instance.smarttag_id = Some(value.parse()?),
                    "bmk" => instance.bookmark_link_target = Some(value.clone()),
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
                            "ln" => {
                                instance.line_properties = Some(Box::new(LineProperties::from_xml_element(child_node)?))
                            }
                            "highlight" => {
                                let color = child_node
                                    .child_nodes
                                    .iter()
                                    .find_map(Color::try_from_xml_element)
                                    .transpose()?
                                    .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "CT_Color"))?;

                                instance.highlight_color = Some(color);
                            }
                            "latin" => instance.latin_font = Some(TextFont::from_xml_element(child_node)?),
                            "ea" => instance.east_asian_font = Some(TextFont::from_xml_element(child_node)?),
                            "cs" => instance.complex_script_font = Some(TextFont::from_xml_element(child_node)?),
                            "sym" => instance.symbol_font = Some(TextFont::from_xml_element(child_node)?),
                            "hlinkClick" => {
                                instance.hyperlink_click = Some(Box::new(Hyperlink::from_xml_element(child_node)?))
                            }
                            "hlinkMouseOver" => {
                                instance.hyperlink_mouse_over = Some(Box::new(Hyperlink::from_xml_element(child_node)?))
                            }
                            "rtl" => {
                                instance.rtl = child_node.text.as_ref().map(parse_xml_bool).transpose()?;
                            }
                            local_name if FillProperties::is_choice_member(local_name) => {
                                instance.fill_properties = Some(FillProperties::from_xml_element(child_node)?);
                            }
                            local_name if EffectProperties::is_choice_member(local_name) => {
                                instance.effect_properties = Some(EffectProperties::from_xml_element(child_node)?);
                            }
                            local_name if TextUnderlineLine::is_choice_member(local_name) => {
                                instance.text_underline_line = Some(TextUnderlineLine::from_xml_element(child_node)?);
                            }
                            local_name if TextUnderlineFill::is_choice_member(local_name) => {
                                instance.text_underline_fill = Some(TextUnderlineFill::from_xml_element(child_node)?);
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextSpacing {
    /// This element specifies the amount of white space that is to be used between lines and paragraphs in the form of
    /// a percentage of the text size. The text size that is used to calculate the spacing here is the text for each run, with
    /// the largest text size having precedence. That is if there is a run of text with 10 point font and within the same
    /// paragraph on the same line there is a run of text with a 12 point font size then the 12 point should be used to
    /// calculate the spacing to be used.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:spcBef>
    ///         <a:spcPct val="200%"/>
    ///       </a:spcBef>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Sample Text</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph of text is formatted to have a spacing before the paragraph text. This spacing is 200% of
    /// the size of the largest text on each line.
    Percent(TextSpacingPercent),

    /// This element specifies the amount of white space that is to be used between lines and paragraphs in the form of
    /// a text point size. The size is specified using points where 100 is equal to 1 point font and 1200 is equal to 12
    /// point.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   …
    ///   <a:p>
    ///     <a:pPr …>
    ///       <a:spcBef>
    ///         <a:spcPts val="1400"/>
    ///       </a:spcBef>
    ///     </a:pPr>
    ///     …
    ///     <a:t>Sample Text</a:t>
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// The above paragraph of text is formatted to have a spacing before the paragraph text. This spacing is a size of 14
    /// points due to val="1400".
    Point(TextSpacingPoint),
}

impl XsdType for TextSpacing {
    fn from_xml_element(xml_node: &XmlNode) -> Result<TextSpacing> {
        match xml_node.local_name() {
            "spcPct" => Ok(TextSpacing::Percent(xml_node.get_val_attribute()?.parse()?)),
            "spcPts" => Ok(TextSpacing::Point(xml_node.get_val_attribute()?.parse()?)),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextSpacing").into()),
        }
    }
}

impl XsdChoice for TextSpacing {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "spcPct" | "spcPts" => true,
            _ => false,
        }
    }
}

/// This element specifies a single tab stop to be used on a line of text when there are one or more tab characters
/// present within the text. When there is more than one present than they should be utilized in increasing position
/// order which is specified via the pos attribute.
///
/// # Xml example
///
/// ```xml
/// <p:txBody>
///   …
///   <a:p>
///     <a:pPr …>
///       <a:tabLst>
///         <a:tab pos="2292350" algn="l"/>
///         <a:tab pos="2627313" algn="l"/>
///         <a:tab pos="2743200" algn="l"/>
///         <a:tab pos="2974975" algn="l"/>
///       </a:tabLst>
///     </a:pPr>
///     …
///     <a:t>Sample Text</a:t>
///     …
///   </a:p>
///   …
/// </p:txBody>
/// ```
///
/// The paragraph within which this <a:tab> information resides has a total of 4 unique tab stops that should be
/// listed in order of increasing position. Along with specifying the tab position each tab allows for the specifying of
/// an alignment.
#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextTabStop {
    /// Specifies the position of the tab stop relative to the left margin. If this attribute is omitted
    /// then the application default for tab stops is used.
    pub position: Option<Coordinate32>,

    /// Specifies the alignment that is to be applied to text using this tab stop. If this attribute is
    /// omitted then the application default for the generating application.
    pub alignment: Option<TextTabAlignType>,
}

impl TextTabStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<TextTabStop> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "pos" => instance.position = Some(value.parse::<Coordinate32>()?),
                    "algn" => instance.alignment = Some(value.parse::<TextTabAlignType>()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}
