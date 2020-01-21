use crate::{
    shared::drawingml::{
        shapedefs::PresetTextShape,
        simpletypes::{
            Angle, Coordinate32, PositiveCoordinate32, TextAnchoringType, TextColumnCount, TextFontScalePercent,
            TextHorizontalOverflowType, TextSpacingPercent, TextVertOverflowType, TextVerticalType, TextWrappingType,
        },
    },
    error::NotGroupMemberError,
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::{XsdChoice, XsdType},
};

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextBodyProperties {
    /// Specifies the rotation that is being applied to the text within the bounding box. If it not
    /// specified, the rotation of the accompanying shape is used. If it is specified, then this is
    /// applied independently from the shape. That is the shape can have a rotation applied in
    /// addition to the text itself having a rotation applied to it. If this attribute is omitted, then a
    /// value of 0, is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where a shape has a rotation of 5400000, or 90 degrees
    /// clockwise applied to it. In addition to this, the text body itself has a rotation of -5400000,
    /// or 90 degrees counter-clockwise applied to it. Then the resulting shape would appear to
    /// be rotated but the text within it would appear as though it had not been rotated at all.
    /// The DrawingML specifying this would look like the following:
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:spPr>
    ///     <a:xfrm rot="5400000">
    ///     …
    ///     </a:xfrm>
    ///   </p:spPr>
    ///   …
    ///   <p:txBody>
    ///     <a:bodyPr rot="-5400000" … />
    ///     …
    ///     (Some text)
    ///     …
    ///   </p:txBody>
    /// </p:sp>
    /// ```
    pub rotate_angle: Option<Angle>,

    /// Specifies whether the before and after paragraph spacing defined by the user is to be
    /// respected. While the spacing between paragraphs is helpful, it is additionally useful to be
    /// able to set a flag as to whether this spacing is to be followed at the edges of the text
    /// body, in other words the first and last paragraphs in the text body. More precisely since
    /// this is a text body level property it should only effect the before paragraph spacing of the
    /// first paragraph and the after paragraph spacing of the last paragraph for a given text
    /// body. If this attribute is omitted, then a value of 0, or false is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where spacing has been defined between multiple
    /// paragraphs within a text body using the spcBef and spcAft paragraph spacing attributes.
    /// For this text body however the user would like to not have this followed for the edge
    /// paragraphs and thus we have the following DrawingML.
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr spcFirstLastPara="0" … />
    ///   …
    ///   <a:p>
    ///     <a:pPr>
    ///       <a:spcBef>
    ///         <a:spcPts val="1800"/>
    ///       </a:spcBef>
    ///       <a:spcAft>
    ///         <a:spcPts val="600"/>
    ///       </a:spcAft>
    ///     </a:pPr>
    ///     …
    ///     (Some text)
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     <a:pPr>
    ///       <a:spcBef>
    ///         <a:spcPts val="1800"/>
    ///       </a:spcBef>
    ///       <a:spcAft>
    ///         <a:spcPts val="600"/>
    ///       </a:spcAft>
    ///     </a:pPr>
    ///     …
    ///     (Some text)
    ///     …
    ///   </a:p>
    ///   …
    /// </p:txBody>
    /// ```
    pub paragraph_spacing: Option<bool>,

    /// Determines whether the text can flow out of the bounding box vertically. This is used to
    /// determine what happens in the event that the text within a shape is too large for the
    /// bounding box it is contained within. If this attribute is omitted, then a value of overflow
    /// is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where we have multiply paragraphs within a shape and the
    /// second causes text to flow outside the shape. By applying the clip value of the
    /// vertOverflow attribute as a body property this overflowing text is now cut off instead of
    /// extending beyond the bounds of the shape.
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr vertOverflow="clip" … />
    ///   …
    ///   <a:p>
    ///     …
    ///     (Some text)
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     …
    ///     (Some longer text)
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub vertical_overflow: Option<TextVertOverflowType>,

    /// Determines whether the text can flow out of the bounding box horizontally. This is used
    /// to determine what happens in the event that the text within a shape is too large for the
    /// bounding box it is contained within. If this attribute is omitted, then a value of overflow
    /// is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where we have multiply paragraphs within a shape and the
    /// second is greater in length and causes text to flow outside the shape. By applying the clip
    /// value of the horzOverflow attribute as a body property this overflowing text now is cut
    /// off instead of extending beyond the bounds of the shape.
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr horzOverflow="clip" … />
    ///   …
    ///   <a:p>
    ///   …
    ///   (Some text)
    ///   …
    ///   </a:p>
    ///   <a:p>
    ///   …
    ///   (Some more text)
    ///   …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    pub horizontal_overflow: Option<TextHorizontalOverflowType>,

    /// Determines if the text within the given text body should be displayed vertically. If this
    /// attribute is omitted, then a value of horz, or no vertical text is implied.
    ///
    /// # Xml example
    ///
    /// Consider the case where the user needs to display text that appears vertical
    /// and has a right to left flow with respect to its columns.
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr vert="wordArtVertRtl" … />
    ///   …
    ///   <a:p>
    ///     …
    ///     <a:t>This is</a:t>
    ///     …
    ///   </a:p>
    ///   <a:p>
    ///     …
    ///     <a:t>some text.</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// In the above sample DrawingML there are two paragraphs denoting a separation
    /// between the text otherwise which are known as either a line or paragraph break.
    /// Because wordArtVertRtl is used here this text is not only displayed in a stacked manner
    /// flowing from top to bottom but also have the first paragraph be displayed to the right of
    /// the second. This is because it is both vertical text and right to left.
    pub vertical_type: Option<TextVerticalType>,

    /// Specifies the wrapping options to be used for this text body. If this attribute is omitted,
    /// then a value of square is implied which wraps the text using the bounding text box.
    pub wrap_type: Option<TextWrappingType>,

    /// Specifies the left inset of the bounding rectangle. Insets are used just as internal margins
    /// for text boxes within shapes. If this attribute is omitted, then a value of 91440 or 0.1
    /// inches is implied.
    pub left_inset: Option<Coordinate32>,

    /// Specifies the top inset of the bounding rectangle. Insets are used just as internal margins
    /// for text boxes within shapes. If this attribute is omitted, then a value of 45720 or 0.05
    /// inches is implied.
    pub top_inset: Option<Coordinate32>,

    /// Specifies the right inset of the bounding rectangle. Insets are used just as internal
    /// margins for text boxes within shapes. If this attribute is omitted, then a value of 91440 or
    /// 0.1 inches is implied.
    pub right_inset: Option<Coordinate32>,

    /// Specifies the bottom inset of the bounding rectangle. Insets are used just as internal
    /// margins for text boxes within shapes. If this attribute is omitted, a value of 45720 or 0.05
    /// inches is implied.
    pub bottom_inset: Option<Coordinate32>,

    /// Specifies the number of columns of text in the bounding rectangle. When applied to a
    /// text run this property takes the width of the bounding box for the text and divides it by
    /// the number of columns specified. These columns are then treated as overflow containers
    /// in that when the previous column has been filled with text the next column acts as the
    /// repository for additional text. When all columns have been filled and text still remains
    /// then the overflow properties set for this text body are used and the text is reflowed to
    /// make room for additional text. If this attribute is omitted, then a value of 1 is implied.
    pub column_count: Option<TextColumnCount>,

    /// Specifies the space between text columns in the text area. This should only apply when
    /// there is more than 1 column present. If this attribute is omitted, then a value of 0 is
    /// implied.
    pub space_between_columns: Option<PositiveCoordinate32>,

    /// Specifies whether columns are used in a right-to-left or left-to-right order. The usage of
    /// this attribute only sets the column order that is used to determine which column
    /// overflow text should go to next. If this attribute is omitted, then a value of 0 or falseis
    /// implied in which case text starts in the leftmost column and flow to the right.
    ///
    /// # Note
    ///
    /// This attribute in no way determines the direction of text but merely the direction
    /// in which multiple columns are used.
    pub rtl_columns: Option<bool>,

    /// Specifies that text within this textbox is converted text from a WordArt object. This is
    /// more of a backwards compatibility attribute that is useful to the application from a
    /// tracking perspective. WordArt was the former way to apply text effects and therefore
    /// this attribute is useful in document conversion scenarios. If this attribute is omitted, then
    /// a value of 0 or false is implied.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr wrap="none" fromWordArt="1" …/>
    ///   …
    /// </p:txBody>
    /// ```
    ///
    /// Because of the presence of the fromWordArt attribute the text within this shape can be
    /// mapped back to the corresponding WordArt during document conversion.
    pub is_from_word_art: Option<bool>,

    /// Specifies the anchoring position of the txBody within the shape. If this attribute is
    /// omitted, then a value of t, or top is implied.
    pub anchor: Option<TextAnchoringType>,

    /// Specifies the centering of the text box. The way it works fundamentally is to determine
    /// the smallest possible "bounds box" for the text and then to center that "bounds box"
    /// accordingly. This is different than paragraph alignment, which aligns the text within the
    /// "bounds box" for the text. This flag is compatible with all of the different kinds of
    /// anchoring. If this attribute is omitted, then a value of 0 or false is implied.
    ///
    /// # Example
    ///
    /// The text within this shape has been both vertically centered with the anchor
    /// attribute and horizontally centered with the anchorCtr attribute.
    ///
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr anchor="ctr" anchorCtr="1" … />
    ///   …
    /// </p:txBody>
    /// ```
    pub anchor_center: Option<bool>,

    /// Forces the text to be rendered anti-aliased regardless of the font size. Certain fonts can
    /// appear grainy around their edges unless they are anti-aliased. Therefore this attribute
    /// allows for the specifying of which bodies of text should always be anti-aliased and which
    /// ones should not. If this attribute is omitted, then a value of 0 or false is implied.
    pub force_antialias: Option<bool>,

    /// Specifies whether text should remain upright, regardless of the transform applied to it
    /// and the accompanying shape transform. If this attribute is omitted, then a value of 0, or
    /// false is implied.
    pub upright: Option<bool>,

    /// Specifies that the line spacing for this text body is decided in a simplistic manner using
    /// the font scene. If this attribute is omitted, a value of 0 or false is implied.
    pub compatible_line_spacing: Option<bool>,

    /// This element specifies when a preset geometric shape should be used to transform a piece of text. This
    /// operation is known formally as a text warp. The generating application should be able to render all preset
    /// geometries enumerated in the TextShapeType list.
    ///
    /// Using any of the presets listed under the ST_TextShapeType list below it is possible to apply a text warp to a run
    /// of DrawingML text via the following steps.
    ///
    /// If you look at any of the text warps in the file format you notice that each consists of two paths. This
    /// corresponds to a top path (first one specified) and a bottom path (second one specified). Now the top path and
    /// the bottom path represent the top line and base line that the text needs to be warped to. This is done in the
    /// following way:
    ///
    /// 1. Compute the rectangle that the unwarped text resides in. (tightest possible rectangle around text, no
    ///    white space except for “space characters”)
    /// 2. Take each of the quadratic and cubic Bezier curves that are used to calculate the original character and
    ///    change their end points and control points by the following method…
    /// 3. Move a vertical line horizontally along the original text rectangle and find the horizontal percentage that
    ///    a given end point or control point lives at. (.5 for the middle for instance)
    /// 4. Now do the same thing for this point vertically. Find the vertical percentage that this point lives at with
    ///    the top and bottom of this text rectangle being the respective top and bottom bounds. (0.0 and 1.0
    ///    respectively)
    /// 5. Now that we have the percentages for a given point in a Bezier equation we can map that to the new
    ///    point in the warped text environment.
    /// 6. Going back to the top and bottom paths specified in the file format we can take these and flatten them
    ///    out to a straight arc (top and bottom might be different lengths)
    /// 7. After they are straight we can measure them both horizontally to find the same percentage point that
    ///    we found within the original text rectangle. (0.5 let’s say)
    /// 8. So then we measure 50% along the top path and 50% along the bottom path, putting the paths back to
    ///    their original curvy shapes.
    /// 9. Once we have these two points we can draw a line between them that serves as our vertical line in the
    ///    original text rectangle (This might not be truly vertical as 50% on the top does not always line up
    ///    with 50% on the bottom. end)
    /// 10. Taking this new line we then follow it from top to bottom the vertical percentage amount that we got
    ///     from step 4.
    /// 11. This is then the new point that should be used in place of the old point in the original text rectangle.
    /// 12. We then continue doing these same steps for each of the end points and control points within the body
    ///     of text. (is applied to a whole body of text only)
    ///
    /// # Xml example
    ///
    /// Consider the case where the user wishes to accent a piece of text by warping it's shape. For this to
    /// occur a preset shape is chosen from the TextShapeType list and applied to the entire body of text.
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr wrap="none" rtlCol="0">
    ///       <a:prstTxWarp prst="textInflate">
    ///       </a:prstTxWarp>
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
    /// # Note
    ///
    /// Horizontal percentages begin at 0.0 and continue to 1.0, left to right. Vertical percentages begin at 0.0
    /// and continue to 1.0, top to bottom.
    ///
    /// Since this is a shape it does have both a shape coordinate system and a path coordinate system.
    pub preset_text_warp: Option<Box<PresetTextShape>>,

    /// Specifies the method of auto fitting this text body.
    pub auto_fit_type: Option<TextAutoFit>,
    // TODO implement
    //pub scene_3d: Option<Scene3D>,
    //pub text_3d: Option<Text3D>,
}

impl TextBodyProperties {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "rot" => instance.rotate_angle = Some(value.parse()?),
                    "spcFirstLastPara" => instance.paragraph_spacing = Some(parse_xml_bool(value)?),
                    "vertOverflow" => instance.vertical_overflow = Some(value.parse()?),
                    "horzOverflow" => instance.horizontal_overflow = Some(value.parse()?),
                    "vert" => instance.vertical_type = Some(value.parse()?),
                    "wrap" => instance.wrap_type = Some(value.parse()?),
                    "lIns" => instance.left_inset = Some(value.parse()?),
                    "tIns" => instance.top_inset = Some(value.parse()?),
                    "rIns" => instance.right_inset = Some(value.parse()?),
                    "bIns" => instance.bottom_inset = Some(value.parse()?),
                    "numCol" => instance.column_count = Some(value.parse()?),
                    "spcCol" => instance.space_between_columns = Some(value.parse()?),
                    "rtlCol" => instance.rtl_columns = Some(parse_xml_bool(value)?),
                    "fromWordArt" => instance.is_from_word_art = Some(parse_xml_bool(value)?),
                    "anchor" => instance.anchor = Some(value.parse()?),
                    "anchorCtr" => instance.anchor_center = Some(parse_xml_bool(value)?),
                    "forceAA" => instance.force_antialias = Some(parse_xml_bool(value)?),
                    "upright" => instance.upright = Some(parse_xml_bool(value)?),
                    "compatLnSpc" => instance.compatible_line_spacing = Some(parse_xml_bool(value)?),
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
                            "prstTxWarp" => {
                                instance.preset_text_warp =
                                    Some(Box::new(PresetTextShape::from_xml_element(child_node)?))
                            }
                            local_name if TextAutoFit::is_choice_member(local_name) => {
                                instance.auto_fit_type = Some(TextAutoFit::from_xml_element(child_node)?);
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum TextAutoFit {
    /// This element specifies that text within the text body should not be auto-fit to the bounding box. Auto-fitting is
    /// when text within a text box is scaled in order to remain inside the text box. If this element is omitted, then
    /// noAutofit or auto-fit off is implied.
    ///
    /// # Xml example
    ///
    /// Consider a text box where the user wishes to have the text extend outside the bounding box. The
    /// following DrawingML would describe this.
    /// ```xml
    /// <p:txBody>
    ///   <a:bodyPr wrap="none" rtlCol="0">
    ///     <a:noAutofit/>
    ///   </a:bodyPr>
    ///   <a:p>
    ///     …
    ///     <a:t>Some text</a:t>
    ///     …
    ///   </a:p>
    /// </p:txBody>
    /// ```
    NoAutoFit,

    /// This element specifies that text within the text body should be normally auto-fit to the bounding box. Autofitting
    /// is when text within a text box is scaled in order to remain inside the text box. If this element is omitted,
    /// then noAutofit or auto-fit off is implied.
    ///
    /// # Xml example
    ///
    /// Consider the situation where a user is building a diagram and needs to have the text for each shape
    /// that they are using stay within the bounds of the shape. An easy way this might be done is by using
    /// normAutofit. The following DrawingML illustrates how this might be accomplished.
    /// ```xml
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr rtlCol="0" anchor="ctr">
    ///       <a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>
    ///     </a:bodyPr>
    ///     …
    ///     <a:p>
    ///       …
    ///       <a:t>Diagram Object 1</a:t>
    ///       …
    ///     </a:p>
    ///   </p:txBody>
    /// </p:sp>
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr rtlCol="0" anchor="ctr">
    ///       <a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>
    ///     </a:bodyPr>
    ///     …
    ///     <a:p>
    ///       …
    ///       <a:t>Diagram Object 2</a:t>
    ///       …
    ///     </a:p>
    ///   </p:txBody>
    /// </p:sp>
    /// ```
    ///
    /// In the above example there are two shapes that have normAutofit turned on so that when the user types more
    /// text within the shape that the text actually resizes to accommodate the new data. For the application to know
    /// how and to what degree the text should be resized two attributes are set for the auto-fit resize logic.
    NormalAutoFit(TextNormalAutoFit),

    /// This element specifies that a shape should be auto-fit to fully contain the text described within it. Auto-fitting is
    /// when text within a shape is scaled in order to contain all the text inside. If this element is omitted, then
    /// NoAutoFit or auto-fit off is implied.
    ///
    /// # Xml example
    ///
    /// Consider the situation where a user is building a diagram and needs to have the text for each shape
    /// that they are using stay within the bounds of the shape. An easy way this might be done is by using ShapeAutoFit.
    /// The following DrawingML illustrates how this might be accomplished.
    ///
    /// ```xml
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr rtlCol="0" anchor="ctr">
    ///       <a:spAutoFit/>
    ///     </a:bodyPr>
    ///     …
    ///     <a:p>
    ///       …
    ///       <a:t>Diagram Object 1</a:t>
    ///       …
    ///     </a:p>
    ///   </p:txBody>
    /// </p:sp>
    /// <p:sp>
    ///   <p:txBody>
    ///     <a:bodyPr rtlCol="0" anchor="ctr">
    ///       <a:spAutoFit/>
    ///     </a:bodyPr>
    ///     …
    ///     <a:p>
    ///       …
    ///       <a:t>Diagram Object 2</a:t>
    ///       …
    ///     </a:p>
    ///   </p:txBody>
    /// </p:sp>
    /// ```
    ///
    /// In the above example there are two shapes that have ShapeAutoFit turned on so that when the user types more
    /// text within the shape that the shape actually resizes to accommodate the new data.
    ShapeAutoFit,
}

impl XsdType for TextAutoFit {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "noAutofit" => Ok(TextAutoFit::NoAutoFit),
            "normAutofit" => Ok(TextAutoFit::NormalAutoFit(TextNormalAutoFit::from_xml_element(
                xml_node,
            )?)),
            "spAutoFit" => Ok(TextAutoFit::ShapeAutoFit),
            _ => Err(NotGroupMemberError::new(xml_node.name.clone(), "EG_TextAutofit").into()),
        }
    }
}

impl XsdChoice for TextAutoFit {
    fn is_choice_member<T: AsRef<str>>(name: T) -> bool {
        match name.as_ref() {
            "noAutofit" | "normAutofit" | "spAutoFit" => true,
            _ => false,
        }
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct TextNormalAutoFit {
    /// Specifies the percentage of the original font size to which each run in the text body is
    /// scaled. In order to auto-fit text within a bounding box it is sometimes necessary to
    /// decrease the font size by a certain percentage. Using this attribute the font within a text
    /// box can be scaled based on the value provided. A value of 100% scales the text to 100%,
    /// while a value of 1% scales the text to 1%. If this attribute is omitted, then a value of 100%
    /// is implied.
    ///
    /// Defaults to 100000
    pub font_scale: Option<TextFontScalePercent>,

    /// Specifies the percentage amount by which the line spacing of each paragraph in the text
    /// body is reduced. The reduction is applied by subtracting it from the original line spacing
    /// value. Using this attribute the vertical spacing between the lines of text can be scaled by
    /// a percent amount. A value of 100% reduces the line spacing by 100%, while a value of 1%
    /// reduces the line spacing by one percent. If this attribute is omitted, then a value of 0% is
    /// implied.
    ///
    /// Defaults to 0
    ///
    /// # Note
    ///
    /// This attribute applies only to paragraphs with percentage line spacing.
    pub line_spacing_reduction: Option<TextSpacingPercent>,
}

impl TextNormalAutoFit {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut font_scale = None;
        let mut line_spacing_reduction = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "fontScale" => font_scale = Some(value.parse::<TextFontScalePercent>()?),
                "lnSpcReduction" => line_spacing_reduction = Some(value.parse::<TextSpacingPercent>()?),
                _ => (),
            }
        }

        Ok(Self {
            font_scale,
            line_spacing_reduction,
        })
    }
}
