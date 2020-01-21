use crate::{
    shared::{
        drawingml::{
            coordsys::PositiveSize2D,
            simpletypes::{Percentage, PositiveCoordinate32},
            text::{bullet::TextListStyle, runformatting::TextFont},
        },
        relationship::RelationshipId,
        sharedtypes::ConformanceClass,
    },
    error::{MissingAttributeError, MissingChildNodeError},
    xml::{parse_xml_bool, XmlNode},
};
use std::{
    error::Error,
    io::{Read, Seek},
    str::FromStr,
};

pub type Result<T> = ::std::result::Result<T, Box<dyn Error>>;

/// This simple type specifies the allowed numbering for the slide identifier.
///
/// Values represented by this type are restricted to: 256 <= n <= 2147483648
pub type SlideId = u32;
/// This simple type sets the bounds for the slide layout id value. This layout id is used to identify the different slide
/// layout designs.
///
/// Values represented by this type are restricted to: 2147483648 <= n
pub type SlideLayoutId = u32;
/// This simple type specifies the allowed numbering for the slide master identifier.
///
/// Values represented by this type are restricted to: 2147483648 <= n
pub type SlideMasterId = u32;
/// This simple type specifies constraints for value of the Bookmark ID seed.
///
/// Values represented by this type are restricted to: 1 <= n <= 2147483648
pub type BookmarkIdSeed = u32;
/// This simple type specifies the slide size coordinate in EMUs (English Metric Units).AsRef
///
/// Values represented by this type are restricted to: 914400 <= n <= 51206400
pub type SlideSizeCoordinate = PositiveCoordinate32;
/// This simple type specifies a name, such as for a comment author or custom show.
pub type Name = String;

/// This simple type specifies the kind of slide size that the slide should be optimized for.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum SlideSizeType {
    /// Slide size should be optimized for 35mm film output
    #[strum(serialize = "mm35")]
    Mm35,
    /// Slide size should be optimized for A3 output
    #[strum(serialize = "a3")]
    A3,
    /// Slide size should be optimized for A4 output
    #[strum(serialize = "a4")]
    A4,
    /// Slide size should be optimized for B4ISO output
    #[strum(serialize = "b4ISO")]
    B4ISO,
    /// Slide size should be optimized for B4JIS output
    #[strum(serialize = "b4JIS")]
    B4JIS,
    /// Slide size should be optimized for B5ISO output
    #[strum(serialize = "b5ISO")]
    B5ISO,
    /// Slide size should be optimized for B5JIS output
    #[strum(serialize = "b5JIS")]
    B5JIS,
    /// Slide size should be optimized for banner output
    #[strum(serialize = "banner")]
    Banner,
    /// Slide size should be optimized for custom output
    #[strum(serialize = "custom")]
    Custom,
    /// Slide size should be optimized for hagaki card output
    #[strum(serialize = "hagakiCard")]
    HagakiCard,
    /// Slide size should be optimized for ledger output
    #[strum(serialize = "ledger")]
    Ledger,
    /// Slide size should be optimized for letter output
    #[strum(serialize = "letter")]
    Letter,
    /// Slide size should be optimized for overhead output
    #[strum(serialize = "overhead")]
    Overhead,
    /// Slide size should be optimized for 16x10 screen output
    #[strum(serialize = "screen16x10")]
    Screen16x10,
    /// Slide size should be optimized for 16x9 screen output
    #[strum(serialize = "screen16x9")]
    Screen16x9,
    /// Slide size should be optimized for 4x3 screen output
    #[strum(serialize = "screen4x3")]
    Screen4x3,
}

/// This simple type specifies the values for photo layouts within a photo album presentation.
/// See Fundamentals And Markup Language Reference for examples
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PhotoAlbumLayout {
    /// Fit Photos to Slide
    #[strum(serialize = "fitToSlide")]
    FitToSlide,
    /// 1 Photo per Slide
    #[strum(serialize = "pic1")]
    Pic1,
    /// 2 Photo per Slide
    #[strum(serialize = "pic2")]
    Pic2,
    /// 4 Photo per Slide
    #[strum(serialize = "pic4")]
    Pic4,
    /// 1 Photo per Slide with Titles
    #[strum(serialize = "picTitle1")]
    PicTitle1,
    /// 2 Photo per Slide with Titles
    #[strum(serialize = "picTitle2")]
    PicTitle2,
    /// 4 Photo per Slide with Titles
    #[strum(serialize = "picTitle4")]
    PicTitle4,
}

/// This simple type specifies the values for photo frame types within a photo album presentation.
/// See Fundamentals And Markup Language Reference for examples
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum PhotoAlbumFrameShape {
    /// Rectangle Photo Frame
    #[strum(serialize = "frameStyle1")]
    FrameStyle1,
    /// Rounded Rectangle Photo Frame
    #[strum(serialize = "frameStyle2")]
    FrameStyle2,
    /// Simple White Photo Frame
    #[strum(serialize = "frameStyle3")]
    FrameStyle3,
    /// Simple Black Photo Frame
    #[strum(serialize = "frameStyle4")]
    FrameStyle4,
    /// Compound Black Photo Frame
    #[strum(serialize = "frameStyle5")]
    FrameStyle5,
    /// Center Shadow Photo Frame
    #[strum(serialize = "frameStyle6")]
    FrameStyle6,
    /// Soft Edge Photo Frame
    #[strum(serialize = "frameStyle7")]
    FrameStyle7,
}

/// This simple type determines if the Embedded object is re-colored to reflect changes to the color schemes.
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum OleObjectFollowColorScheme {
    /// Setting this enumeration causes the Embedded object to not respond to changes in the color scheme in the
    /// presentation.
    #[strum(serialize = "none")]
    None,
    /// Setting this enumeration causes the Embedded object to respond to all changes in the color scheme in the
    /// presentation.
    #[strum(serialize = "full")]
    Full,
    /// Setting this enumeration causes the Embedded object to respond only to changes in the text and background
    /// colors of the color scheme in the presentation.
    #[strum(serialize = "textAndBackground")]
    TextAndBackground,
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct CustomerDataList {
    pub customer_data_list: Vec<RelationshipId>,
    /// This element specifies the existence of customer data in the form of tags. This allows for the storage of customer
    /// data within the PresentationML framework. While this is similar to the ext tag in that it can be used store
    /// information, this tag mainly focuses on referencing to other parts of the presentation document. This is
    /// accomplished via the relationship identification attribute that is required for all specified tags.
    pub tags: Option<RelationshipId>,
}

impl CustomerDataList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "custData" => {
                        let id = child_node
                            .attributes
                            .get("r:id")
                            .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                            .clone();
                        instance.customer_data_list.push(id);
                    }
                    "tags" => {
                        let id = child_node
                            .attributes
                            .get("r:id")
                            .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                            .clone();
                        instance.tags = Some(id);
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideSize {
    /// Specifies the length of the extents rectangle in EMUs. This rectangle shall dictate the size
    /// of the object as displayed (the result of any scaling to the original object).
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <... cx="1828800" cy="200000"/>
    /// ```xml
    pub width: SlideSizeCoordinate,
    /// Specifies the width of the extents rectangle in EMUs. This rectangle shall dictate the size
    /// of the object as displayed (the result of any scaling to the original object).
    ///
    /// # Xml example
    ///
    /// ```xml
    /// < ... cx="1828800" cy="200000"/>
    /// ```
    pub height: SlideSizeCoordinate,
    /// Specifies the kind of slide size that should be used. This identifies in particular the
    /// expected delivery platform for this presentation.
    pub size_type: Option<SlideSizeType>,
}

impl SlideSize {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut width = None;
        let mut height = None;
        let mut size_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "cx" => width = Some(value.parse()?),
                "cy" => height = Some(value.parse()?),
                "type" => size_type = Some(value.parse()?),
                _ => (),
            }
        }

        let width = width.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cx"))?;
        let height = height.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "cy"))?;

        Ok(Self {
            width,
            height,
            size_type,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideIdListEntry {
    /// Specifies the slide identifier that is to contain a value that is unique throughout the presentation.
    pub id: SlideId,
    /// Specifies the relationship identifier that is used in conjunction with a corresponding
    /// relationship file to resolve the location within a presentation of the sld element defining
    /// this slide.
    pub relationship_id: RelationshipId,
}

impl SlideIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;
        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideLayoutIdListEntry {
    /// Specifies the identification number that uniquely identifies this slide layout within the
    /// presentation file.
    pub id: Option<SlideLayoutId>,
    /// Specifies the relationship id value that the generating application can use to resolve
    /// which slide layout is used in the creation of the slide. This relationship id is used within
    /// the relationship file for the master slide to expose the location of the corresponding
    /// layout file within the presentation.
    pub relationship_id: RelationshipId,
}

impl SlideLayoutIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SlideLayoutIdList(pub Vec<SlideLayoutIdListEntry>);

impl SlideLayoutIdList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let id_list = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "sldLayoutId")
            .map(SlideLayoutIdListEntry::from_xml_element)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self(id_list))
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SlideMasterIdListEntry {
    /// Specifies the slide master identifier that is to contain a value that is unique throughout
    /// the presentation.
    pub id: Option<SlideMasterId>,
    /// Specifies the relationship identifier that is used in conjunction with a corresponding
    /// relationship file to resolve the location within a presentation of the sldMaster element
    /// defining this slide master.
    pub relationship_id: RelationshipId,
}

impl SlideMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut relationship_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "id" => id = Some(value.parse()?),
                "r:id" => relationship_id = Some(value.clone()),
                _ => (),
            }
        }

        let relationship_id =
            relationship_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { id, relationship_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct NotesMasterIdListEntry {
    /// Specifies the relationship identifier that is used in conjunction with a corresponding
    /// relationship file to resolve the location within a presentation of the notesMaster element
    /// defining this notes master.
    pub relationship_id: RelationshipId,
}

impl NotesMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let relationship_id = xml_node
            .attributes
            .get("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?
            .clone();

        Ok(Self { relationship_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct HandoutMasterIdListEntry {
    /// Specifies the relationship identifier that is used in conjunction with a corresponding
    /// relationship file to resolve the location within a presentation of the handoutMaster
    /// element defining this handout master.
    pub relationship_id: RelationshipId,
}

impl HandoutMasterIdListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let relationship_id = xml_node
            .attributes
            .get("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?
            .clone();

        Ok(Self { relationship_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct EmbeddedFontListEntry {
    /// This element specifies specific properties describing an embedded font. Once specified, this font is available
    /// for use within the presentation.
    /// Within a font specification there can be regular, bold, italic and boldItalic versions of the font specified.
    /// The actual font data for each of these is referenced using a relationships file that contains links to all
    /// available fonts.
    /// This font data contains font information for each of the characters to be made available in each version of
    /// the font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Font Substitution Logic
    ///
    /// If the specified font is not available on a system being used for rendering, then the attributes of this
    /// element are to be utilized in selecting an alternate font.
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub font: TextFont,
    /// This element specifies a regular embedded font that is linked to a parent typeface. Once specified, this regular
    /// version of the given typeface name is available for use within the presentation. The actual font data is
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font
    /// information for each of the characters to be made available.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub regular: Option<RelationshipId>,
    /// This element specifies a bold embedded font that is linked to a parent typeface. Once specified, this bold
    /// version of the given typeface name is available for use within the presentation. The actual font data is
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font
    /// information for each of the characters to be made available.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:bold r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
    pub bold: Option<RelationshipId>,
    /// This element specifies an italic embedded font that is linked to a parent typeface. Once specified, this italic
    /// version of the given typeface name is available for use within the presentation. The actual font data is
    /// referenced using a relationships file that contains links to all fonts available. This font data contains font
    /// information for each of the characters to be made available.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:italic r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine which
    /// characters are to be stored in the corresponding font data files.
    pub italic: Option<RelationshipId>,
    /// This element specifies a bold italic embedded font that is linked to a parent typeface. Once specified, this
    /// bold italic version of the given typeface name is available for use within the presentation. The actual font
    /// data is referenced using a relationships file that contains links to all fonts available. This font data
    /// contains font information for each of the characters to be made available.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:boldItalic r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
    pub bold_italic: Option<RelationshipId>,
}

impl EmbeddedFontListEntry {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut font = None;
        let mut regular = None;
        let mut bold = None;
        let mut italic = None;
        let mut bold_italic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "font" => font = Some(TextFont::from_xml_element(child_node)?),
                "regular" => {
                    let id = child_node
                        .attributes
                        .get("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                        .clone();
                    regular = Some(id);
                }
                "bold" => {
                    let id = child_node
                        .attributes
                        .get("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                        .clone();
                    bold = Some(id);
                }
                "italic" => {
                    let id = child_node
                        .attributes
                        .get("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                        .clone();
                    italic = Some(id);
                }
                "boldItalic" => {
                    let id = child_node
                        .attributes
                        .get("r:id")
                        .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                        .clone();
                    bold_italic = Some(id);
                }
                _ => (),
            }
        }

        let font = font.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "font"))?;

        Ok(Self {
            font,
            regular,
            bold,
            italic,
            bold_italic,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SlideRelationshipList(pub Vec<RelationshipId>);

impl SlideRelationshipList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let relationship_ids =
            xml_node
                .child_nodes
                .iter()
                .filter(|child_node| child_node.local_name() == "sld")
                .map(|child_node| {
                    child_node.attributes.get("r:id").cloned().ok_or_else(|| {
                        Box::<dyn Error>::from(MissingAttributeError::new(child_node.name.clone(), "r:id"))
                    })
                })
                .collect::<Result<Vec<_>>>()?;

        Ok(Self(relationship_ids))
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomShow {
    /// Specifies a name for the custom show.
    pub name: Name,
    /// Specifies the identification number for this custom show. This should be unique among
    /// all the custom shows within the corresponding presentation.
    pub id: u32,
    pub slides: SlideRelationshipList,
}

impl CustomShow {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<CustomShow> {
        let mut name = None;
        let mut id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "name" => name = Some(value.clone()),
                "id" => id = Some(value.parse::<u32>()?),
                _ => (),
            }
        }

        let name = name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;
        let id = id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        let slides = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "sldLst")
            .ok_or_else(|| Box::<dyn Error>::from(MissingChildNodeError::new(xml_node.name.clone(), "sldLst")))
            .and_then(SlideRelationshipList::from_xml_element)?;

        Ok(Self { name, id, slides })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct PhotoAlbum {
    /// Specifies whether all pictures in the photo album are to be displayed as black and white.
    ///
    /// Defaults to false
    pub black_and_white: Option<bool>,
    /// Specifies whether to show captions for pictures in the photo album. Captions are text
    /// boxes grouped with each image, with the group set to not allow ungrouping.
    ///
    /// Defaults to false
    pub show_captions: Option<bool>,
    /// Specifies the layout that is to be used to arrange the pictures in the photo album on
    /// individual slides.
    ///
    /// Defaults to PhotoAlbumLayout::FitToSlide
    pub layout: Option<PhotoAlbumLayout>,
    /// Specifies the frame type that is to be used on all the pictures in the photo album.
    ///
    /// Defaults to PhotoAlbumFrameShape::FrameStyle1
    pub frame: Option<PhotoAlbumFrameShape>,
}

impl PhotoAlbum {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "bw" => instance.black_and_white = Some(parse_xml_bool(value)?),
                    "showCaptions" => instance.show_captions = Some(parse_xml_bool(value)?),
                    "layout" => instance.layout = Some(value.parse()?),
                    "frame" => instance.frame = Some(value.parse()?),
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Kinsoku {
    /// Specifies the corresponding East Asian language that these settings apply to.
    pub language: Option<String>,
    /// Specifies the characters that cannot start a line of text.
    pub invalid_start_chars: String,
    /// Specifies the characters that cannot end a line of text.
    pub invalid_end_chars: String,
}

impl Kinsoku {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut language = None;
        let mut invalid_start_chars = None;
        let mut invalid_end_chars = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "lang" => language = Some(value.clone()),
                "invalStChars" => invalid_start_chars = Some(value.clone()),
                "invalEndChars" => invalid_end_chars = Some(value.clone()),
                _ => (),
            }
        }

        let invalid_start_chars =
            invalid_start_chars.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalStChars"))?;
        let invalid_end_chars =
            invalid_end_chars.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "invalEndChars"))?;

        Ok(Self {
            language,
            invalid_start_chars,
            invalid_end_chars,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct ModifyVerifier {
    /// Specifies the specific cryptographic hashing algorithm which shall be used along with the
    /// salt attribute and input password in order to compute the hash value.
    ///
    /// The following values are reserved:
    /// * MD2: Specifies that the MD2 algorithm, as defined by RFC 1319, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    ///
    /// * MD4: Specifies that the MD4 algorithm, as defined by RFC 1320, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    ///
    /// * MD5: Specifies that the MD5 algorithm, as defined by RFC 1321, shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    ///
    /// * RIPEMD-128: Specifies that the RIPEMD-128 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// __It is recommended that applications should avoid using this algorithm to store new hash values, due to
    /// publically known breaks.__
    ///
    /// * RIPEMD-160: Specifies that the RIPEMD-160 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-1: Specifies that the SHA-1 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-256: Specifies that the SHA-256 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-384: Specifies that the SHA-384 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * SHA-512: Specifies that the SHA-512 algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    /// * WHIRLPOOL: Specifies that the WHIRLPOOL algorithm, as defined by ISO/IEC 10118-3:2004 shall be used.
    ///
    /// # Xml example
    ///
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// < ... algorithmName="SHA-1" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The algorithm_name attribute value of “SHA-1” specifies that the SHA-1 hashing algorithm must be used to
    /// generate a hash from the user-defined password.
    pub algorithm_name: Option<String>,
    /// Specifies the hash value for the password required to edit this chartsheet. This value shall
    /// be compared with the resulting hash value after hashing the user-supplied password
    /// using the algorithm specified by the preceding attributes and parent XML element, and if
    /// the two values match, the protection shall no longer be enforced.
    ///
    /// If this value is omitted, then the reservationPassword attribute shall contain the
    /// password hash for the workbook.
    ///
    /// # Xml example
    ///
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... algorithmName="SHA-1" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The hashValue attribute value of 9oN7nWkCAyEZib1RomSJTjmPpCY= specifies that the
    /// user-supplied password must be hashed using the pre-processing defined by the parent
    /// element (if any) followed by the SHA-1 algorithm (specified via the algorithmName
    /// attribute value of SHA-1 ) and that the resulting has value must be
    /// 9oN7nWkCAyEZib1RomSJTjmPpCY= for the protection to be disabled.
    pub hash_value: Option<String>,
    /// Specifies the salt which was prepended to the user-supplied password before it was
    /// hashed using the hashing algorithm defined by the preceding attribute values to generate
    /// the hashValue attribute, and which shall also be prepended to the user-supplied
    /// password before attempting to generate a hash value for comparison. A salt is a random
    /// string which is added to a user-supplied password before it is hashed in order to prevent
    /// a malicious party from pre-calculating all possible password/hash combinations and
    /// simply using those pre-calculated values (often referred to as a "dictionary attack").
    ///
    /// If this attribute is omitted, then no salt shall be prepended to the user-supplied password
    /// before it is hashed for comparison with the stored hash value.
    ///
    /// # Xml example
    ///
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... saltValue="ZUdHa+D8F/OAKP3I7ssUnQ==" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    ///
    /// The saltValue attribute value of ZUdHa+D8F/OAKP3I7ssUnQ== specifies that the user-
    /// supplied password must have this value prepended before it is run through the specified
    /// hashing algorithm to generate a resulting hash value for comparison.
    pub salt_value: Option<String>,
    /// Specifies the number of times the hashing function shall be iteratively run (runs using
    /// each iteration's result plus a 4 byte value (0-based, little endian) containing the number
    /// of the iteration as the input for the next iteration) when attempting to compare a user-
    /// supplied password with the value stored in the hashValue attribute.
    ///
    /// # Rationale
    ///
    /// Running the algorithm many times increases the cost of exhaustive search
    /// attacks correspondingly. Storing this value allows for the number of iterations to be
    /// increased over time to accommodate faster hardware (and hence the ability to run more
    /// iterations in less time).
    ///
    /// # Xml example
    ///
    /// Consider an Office Open XML document with the following information stored in one of its protection elements:
    /// ```xml
    /// <... spinCount="100000" hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" />
    /// ```
    /// The spinCount attribute value of 100000 specifies that the hashing function must be run
    /// one hundred thousand times to generate a hash value for comparison with the
    /// hashValue attribute.
    pub spin_value: Option<u32>,
}

impl ModifyVerifier {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "algorithmName" => instance.algorithm_name = Some(value.clone()),
                "hashValue" => instance.hash_value = Some(value.clone()),
                "saltValue" => instance.salt_value = Some(value.clone()),
                "spinValue" => instance.spin_value = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

/// This element specifies within it fundamental presentation-wide properties.
///
/// # Xml example
///
/// Consider the following presentation with a single slide master and two slides. In addition to these
/// commonly used elements there can also be the specification of other properties such as slide size, notes size and
/// default text styles.
/// ```xml
/// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="...">
///  <p:sldMasterIdLst>
///    <p:sldMasterId id="2147483648" r:id="rId1"/>
///  </p:sldMasterIdLst>
///  <p:sldIdLst>
///    <p:sldId id="256" r:id="rId3"/>
///    <p:sldId id="257" r:id="rId4"/>
///  </p:sldIdLst>
///  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
///  <p:notesSz cx="6858000" cy="9144000"/>
///  <p:defaultTextStyle>
///  ...
///  </p:defaultTextStyle>
/// </p:presentation>
/// ```
#[derive(Default, Debug, Clone, PartialEq)]
pub struct Presentation {
    /// Specifies the scaling to be used when the presentation is embedded in another
    /// document. The embedded slides are to be scaled by this percentage.
    ///
    /// Defaults to 50_000
    pub server_zoom: Option<Percentage>,
    /// Specifies the first slide number in the presentation.
    ///
    /// Defaults to 1
    pub first_slide_num: Option<i32>,
    /// Specifies whether to show the header and footer placeholders on the title slides.
    ///
    /// Defaults to true
    pub show_special_placeholders_on_title_slide: Option<bool>,
    /// Specifies if the current view of the user interface is oriented right-to-left or left-to-right.
    /// The view is right-to-left is this value is set to true, and left-to-right otherwise.
    ///
    /// Defaults to false
    pub rtl: Option<bool>,
    /// Specifies whether to automatically remove personal information when the presentation
    /// document is saved.
    ///
    /// Defaults to false
    pub remove_personal_info_on_save: Option<bool>,
    /// Specifies whether the generating application is to be in a compatibility mode which
    /// serves to inform the user of any loss of content or functionality when working with older
    /// formats.
    ///
    /// Defaults to false
    pub compatibility_mode: Option<bool>,
    /// Specifies whether to use strict characters for starting and ending lines of Japanese text.
    ///
    /// Defaults to true
    pub strict_first_and_last_chars: Option<bool>,
    /// Specifies whether the generating application should automatically embed true type fonts or not.
    ///
    /// Defaults to false
    pub embed_true_type_fonts: Option<bool>,
    /// Specifies to save only the subset of characters used in the presentation when a font is embedded.
    ///
    /// Defaults to false
    pub save_subset_fonts: Option<bool>,
    /// Specifies whether the generating application should automatically compress all pictures
    /// for this presentation.
    ///
    /// Defaults to true
    pub auto_compress_pictures: Option<bool>,
    /// Specifies a seed for generating bookmark IDs to ensure IDs remain unique across the
    /// document. This value specifies the number to be used as the ID for the next new
    /// bookmark created.
    ///
    /// Defaults to 1
    pub bookmark_id_seed: Option<BookmarkIdSeed>,
    /// Specifies the conformance class (§2.1) to which the PresentationML document conforms.
    ///
    /// Defaults to ConformanceClass::Transitional
    pub conformance: Option<ConformanceClass>,
    /// This element specifies a list of identification information for the slide master slides that are available
    /// within the corresponding presentation.
    /// A slide master is a slide that is specifically designed to be a template for all related child layout slides.
    ///
    /// The SlideMasterIdListEntry specifies a slide master that is available within the corresponding presentation.
    /// A slide master is a slide that is specifically designed to be a template for all related child layout slides.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:sldMasterIdLst>
    ///     <p:sldMasterId id="2147483648" r:id="rId1"/>
    ///   </p:sldMasterIdLst>
    ///   ...
    /// </p:presentation>
    pub slide_master_id_list: Vec<SlideMasterIdListEntry>,
    /// The specifies a list of identification information for the notes master slides that are available within the
    /// corresponding presentation. A notes master is a slide that is specifically designed for the printing of the slide
    /// along with any attached notes.
    ///
    /// The NotesMasterIdListEntry specifies a notes master that is available within the corresponding presentation.
    /// A notes master is a slide that is specifically designed for the printing of the slide along with any
    /// attached notes.
    ///
    /// # Xml example
    ///
    /// Consider the following specification of a notes master within a presentation
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:notesMasterIdLst>
    ///     <p:notesMasterId r:id="rId8"/>
    ///   </p:notesMasterIdLst>
    ///   ...
    /// </p:presentation>
    /// ```xml
    ///
    /// # Note
    ///
    /// Even though the reference documentation states that this element is a list, the Xml schema states that it has
    /// only a single element.
    pub notes_master_id: Option<NotesMasterIdListEntry>,
    /// This element specifies a list of identification information for the handout master slides that are available
    /// within the corresponding presentation.
    /// A handout master is a slide that is specifically designed for printing as a handout.
    ///
    /// The HandoutMasterIdListEntry specifies a handout master that is available within the corresponding presentation. A handout
    /// master is a slide that is specifically designed for printing as a handout.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:handoutMasterIdLst>
    ///     <p:handoutMasterId r:id="rId8"/>
    ///   </p:handoutMasterIdLst>
    ///   ...
    /// </p:presentation>
    /// ```
    ///
    /// # Note
    ///
    /// Even though the reference documentation states that this element is a list, the Xml schema states that it has
    /// only a single element.
    pub handout_master_id: Option<HandoutMasterIdListEntry>,
    /// This element specifies a list of identification information for the slides that are available within the
    /// corresponding presentation. A slide contains the information that is specific to a single slide such as slide-
    /// specific shape and text information.
    ///
    /// The SlideIdListEntry specifies a list of presentation slides. A presentation slide contains the information
    /// that is specific to a single slide such as slide-specific shape and text information.
    pub slide_id_list: Vec<SlideIdListEntry>,
    /// This element specifies the size of the presentation slide surface. Objects within a presentation slide can be
    /// specified outside these extents, but this is the size of background surface that is shown when the slide is
    /// presented or printed.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
    ///   ...
    /// </p:presentation>
    pub slide_size: Option<SlideSize>,
    /// This element specifies the size of slide surface used for notes slides and handout slides. Objects within a
    /// notes slide can be specified outside these extents, but the notes slide has a background surface of the
    /// specified size when presented or printed.
    /// This element is intended to specify the region to which content is fitted in any special format of printout
    /// the application might choose to generate, such as an outline handout.
    ///
    /// # Xml example
    ///
    /// Consider the following specifying of the size of a notes slide.
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:notesSz cx="9144000" cy="6858000"/>
    ///   ...
    /// </p:presentation>
    /// ```
    pub notes_size: Option<PositiveSize2D>,
    /// This element specifies that references to smart tags exist within this document.
    /// To denote the location of smart tags on individual runs of text, there smart tag identifier attributes are
    /// specified for each run to which a smart tag applies.
    /// These are further specified in the run property attributes within DrawingML.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:presentation>
    ///   ...
    ///   <p:smartTags r:id="rId1"/>
    /// </p:presentation>
    /// ```
    ///
    /// The presence of the smartTags element specifies that there is smart tag information within the PresentationML
    /// package. Individual runs are then inspected for the value of the smtId attribute to determine where smart tags
    /// might apply, for example:
    ///
    /// ```xml
    /// <p:txBody>
    ///  <a:bodyPr/>
    ///  <a:lstStyle/>
    ///  <a:p>
    ///    <a:r>
    ///      <a:rPr lang="en-US" dirty="0" smtId="1"/>
    ///      <a:t>CNTS</a:t>
    ///    </a:r>
    ///    <a:endParaRPr lang="en-US" dirty="0"/>
    ///  </a:p>
    /// </p:txBody>
    /// ```
    ///
    /// In the sample above there is a smart tag identifier of 1 specified for this run of text to denote that the text
    /// should be inspected for smart tag information.
    ///
    /// # Note
    ///
    /// For a complete definition of smart tags, which are semantically identical throughout Office Open XML,
    /// see §17.5.1.
    pub smart_tags: Option<RelationshipId>,
    /// This element specifies a list of fonts that are embedded within the corresponding presentation. The font data
    /// for these fonts is stored alongside the other document parts within the document container. The actual font
    /// data is referenced within the EmbeddedFontListEntry element.
    ///
    /// The EmbeddedFontListEntry element specifies an embedded font. Once specified, this font is available for use within the
    /// presentation.
    /// Within a font specification there can be regular, bold, italic and boldItalic versions of the font specified.
    /// The actual font data for each of these is referenced using a relationships file that contains links to all
    /// available fonts.
    /// This font data contains font information for each of the characters to be made available in each version of the
    /// font.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:embeddedFont>
    ///   <p:font typeface="MyFont" pitchFamily="34" charset="0"/>
    ///   <p:regular r:id="rId2"/>
    /// </p:embeddedFont>
    /// ```
    ///
    /// # Note
    ///
    /// Not all characters for a typeface must be stored. It is up to the generating application to determine
    /// which characters are to be stored in the corresponding font data files.
    pub embedded_font_list: Vec<EmbeddedFontListEntry>,
    /// This element specifies a list of all custom shows that are available within the corresponding presentation.
    /// A custom show is a defined slide sequence that allows for the displaying of the slides with the presentation in
    /// any arbitrary order.
    pub custom_show_list: Vec<CustomShow>,
    /// This element specifies that the corresponding presentation contains a photo album. A photo album specifies a
    /// list of images within the presentation that spread across one or more slides, all of which share a consistent
    /// layout. Each image in the album is formatted with a consistent style. This functionality enables the application
    /// to manage all of the images together and modify their ordering, layout, and formatting as a set.
    ///
    /// This element does not enforce the specified properties on individual photo album images; rather, it specifies
    /// common settings that should be applied by default to all photo album images and their containing slides.
    /// Images that are part of the photo album are identified by the presence of the isPhoto element in the definition
    /// of the picture.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:presentation xmlns:a="..." xmlns:r="..." xmlns:p="..." embedTrueTypeFonts="1">
    ///   ...
    ///   <p:photoAlbum bw="1" layout="2pic"/>
    ///   ...
    /// </p:presentation>
    pub photo_album: Option<PhotoAlbum>,
    /// This element allows for the specifying of customer defined data within the PresentationML framework.
    /// References to custom data or tags can be defined within this list.
    ///
    /// The elements of this list specifies customer data which allows for the specifying and persistence of customer
    /// specific data within the presentation.
    pub customer_data_list: Option<CustomerDataList>,
    /// This element specifies the presentation-wide kinsoku settings that define the line breaking behaviour of East
    /// Asian text within the corresponding presentation.
    pub kinsoku: Option<Box<Kinsoku>>,
    /// This element specifies the default text styles that are to be used within the presentation.
    /// The text style defined here can be referenced when inserting a new slide if that slide is not associated with a
    /// master slide or if no styling information has been otherwise specified for the text within the
    /// presentation slide.
    pub default_text_style: Option<Box<TextListStyle>>,
    /// This element specifies the write protection settings which have been applied to a PresentationML document.
    /// Write protection refers to a mode in which the document's contents should not be modified, and the document
    /// should not be resaved using the same file name.
    ///
    /// When present, the application shall require a password to enable modifications to the document. If the
    /// supplied password does not match the hash value in this attribute, then write protection shall be enabled.
    /// If this element is omitted, then no write protection shall be applied to the current document.
    /// Since this protection does not encrypt the document, malicious applications might circumvent its use.
    ///
    /// The password supplied to the algorithm is to be a UTF-16LE encoded string; strings longer than 510 octets are
    /// truncated to 510 octets. If there is a leading BOM character (U+FEFF) in the encoded password it is removed
    /// before hash calculation. The attributes of this element specify the algorithm to be used to verify the password
    /// provided by the user.
    ///
    /// # Xml example
    ///
    /// Consider a PresentationML document that can only be opened in a write protected state unless a
    /// password is provided, in which case the file would be opened in an editable state. This requirement would be
    /// specified using the following PresentationML:
    /// ```xml
    /// <p:modifyVerifier p:algorithmName="SHA-512" ...
    /// p:hashValue="9oN7nWkCAyEZib1RomSJTjmPpCY=" ... />
    /// ```
    /// ...In order for the hosting application to enable edits to the document, the hosting application would have to
    /// be provided with a password that the hosting application would then hash using the algorithm specified by the
    /// algorithm attributes and compare to the value of the hashValue attribute ( 9oN7nWkCAyEZib1RomSJTjmPpCY= ).
    /// If the two values matched, the file would be opened in an editable state.
    pub modify_verifier: Option<Box<ModifyVerifier>>,
}

impl Presentation {
    pub fn from_zip<R>(zipper: &mut zip::ZipArchive<R>) -> Result<Self>
    where
        R: Read + Seek,
    {
        let mut presentation_file = zipper.by_name("ppt/presentation.xml")?;
        let mut xml_string = String::new();
        presentation_file.read_to_string(&mut xml_string)?;

        let root = XmlNode::from_str(xml_string.as_str())?;
        Self::from_xml_element(&root)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .attributes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, (attr, value)| {
                match attr.as_ref() {
                    "serverZoom" => instance.server_zoom = Some(value.parse()?),
                    "firstSlideNum" => instance.first_slide_num = Some(value.parse()?),
                    "showSpecialPlsOnTitleSld" => {
                        instance.show_special_placeholders_on_title_slide = Some(parse_xml_bool(value)?);
                    }
                    "rtl" => instance.rtl = Some(parse_xml_bool(value)?),
                    "removePersonalInfoOnSave" => instance.remove_personal_info_on_save = Some(parse_xml_bool(value)?),
                    "compatMode" => instance.compatibility_mode = Some(parse_xml_bool(value)?),
                    "strictFirstAndLastChars" => instance.strict_first_and_last_chars = Some(parse_xml_bool(value)?),
                    "embedTrueTypeFonts" => instance.embed_true_type_fonts = Some(parse_xml_bool(value)?),
                    "saveSubsetFonts" => instance.save_subset_fonts = Some(parse_xml_bool(value)?),
                    "autoCompressPictures" => instance.auto_compress_pictures = Some(parse_xml_bool(value)?),
                    "bookmarkIdSeed" => instance.bookmark_id_seed = Some(value.parse()?),
                    "conformance" => instance.conformance = Some(value.parse()?),
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
                            "sldMasterIdLst" => {
                                instance.slide_master_id_list = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|sld_master_id_node| sld_master_id_node.local_name() == "sldMasterId")
                                    .map(SlideMasterIdListEntry::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;
                            }
                            "notesMasterIdLst" => {
                                instance.notes_master_id = child_node
                                    .child_nodes
                                    .iter()
                                    .find(|notes_master_id_node| notes_master_id_node.local_name() == "notesMasterId")
                                    .map(NotesMasterIdListEntry::from_xml_element)
                                    .transpose()?;
                            }
                            "handoutMasterIdLst" => {
                                instance.handout_master_id = child_node
                                    .child_nodes
                                    .iter()
                                    .find(|handout_master_id_node| {
                                        handout_master_id_node.local_name() == "handoutMasterId"
                                    })
                                    .map(HandoutMasterIdListEntry::from_xml_element)
                                    .transpose()?;
                            }
                            "sldIdLst" => {
                                instance.slide_id_list = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|slide_id_node| slide_id_node.local_name() == "sldId")
                                    .map(SlideIdListEntry::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;
                            }
                            "sldSz" => instance.slide_size = Some(SlideSize::from_xml_element(child_node)?),
                            "notesSz" => instance.notes_size = Some(PositiveSize2D::from_xml_element(child_node)?),
                            "smartTags" => {
                                let r_id = child_node
                                    .attributes
                                    .get("r:id")
                                    .ok_or_else(|| MissingAttributeError::new(child_node.name.clone(), "r:id"))?
                                    .clone();

                                instance.smart_tags = Some(r_id);
                            }
                            "embeddedFontLst" => {
                                instance.embedded_font_list = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|embedded_font_node| embedded_font_node.local_name() == "embeddedFont")
                                    .map(EmbeddedFontListEntry::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;
                            }
                            "custShowLst" => {
                                instance.custom_show_list = child_node
                                    .child_nodes
                                    .iter()
                                    .filter(|cust_show_node| cust_show_node.local_name() == "custShow")
                                    .map(CustomShow::from_xml_element)
                                    .collect::<Result<Vec<_>>>()?;
                            }
                            "photoAlbum" => instance.photo_album = Some(PhotoAlbum::from_xml_element(child_node)?),
                            "custDataLst" => {
                                instance.customer_data_list = Some(CustomerDataList::from_xml_element(child_node)?)
                            }
                            "kinsoku" => instance.kinsoku = Some(Box::new(Kinsoku::from_xml_element(child_node)?)),
                            "defaultTextStyle" => {
                                instance.default_text_style =
                                    Some(Box::new(TextListStyle::from_xml_element(child_node)?))
                            }
                            "modifyVerifier" => {
                                instance.modify_verifier = Some(Box::new(ModifyVerifier::from_xml_element(child_node)?))
                            }
                            _ => (),
                        }

                        Ok(instance)
                    })
            })
    }
}
