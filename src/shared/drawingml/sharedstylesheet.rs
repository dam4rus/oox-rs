use crate::{
    error::{MissingAttributeError, MissingChildNodeError},
    xml::XmlNode,
    xsdtypes::XsdChoice,
};
use super::{
    colors::{Color, CustomColor},
    simpletypes::ColorSchemeIndex,
    styles::{DefaultShapeDefinition, FontScheme, StyleMatrix},
};
use log::trace;
use std::{io::Read, str::FromStr};
use zip::read::ZipFile;

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Debug, Clone, PartialEq)]
pub struct ColorMapping {
    /// A color defined which is associated as the first background color.
    pub background1: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the first text color.
    pub text1: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the second background color.
    pub background2: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the second text color.
    pub text2: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 1 color.
    pub accent1: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 2 color.
    pub accent2: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 3 color.
    pub accent3: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 4 color.
    pub accent4: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 5 color.
    pub accent5: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the accent 6 color.
    pub accent6: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the color for a hyperlink.
    pub hyperlink: ColorSchemeIndex,

    /// Specifies a color defined which is associated as the color for a followed hyperlink.
    pub followed_hyperlink: ColorSchemeIndex,
}

impl ColorMapping {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut background1 = None;
        let mut text1 = None;
        let mut background2 = None;
        let mut text2 = None;
        let mut accent1 = None;
        let mut accent2 = None;
        let mut accent3 = None;
        let mut accent4 = None;
        let mut accent5 = None;
        let mut accent6 = None;
        let mut hyperlink = None;
        let mut followed_hyperlink = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "bg1" => background1 = Some(value.parse()?),
                "tx1" => text1 = Some(value.parse()?),
                "bg2" => background2 = Some(value.parse()?),
                "tx2" => text2 = Some(value.parse()?),
                "accent1" => accent1 = Some(value.parse()?),
                "accent2" => accent2 = Some(value.parse()?),
                "accent3" => accent3 = Some(value.parse()?),
                "accent4" => accent4 = Some(value.parse()?),
                "accent5" => accent5 = Some(value.parse()?),
                "accent6" => accent6 = Some(value.parse()?),
                "hlink" => hyperlink = Some(value.parse()?),
                "folHlink" => followed_hyperlink = Some(value.parse()?),
                _ => (),
            }
        }

        let background1 = background1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bg1"))?;
        let text1 = text1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "tx1"))?;
        let background2 = background2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bg2"))?;
        let text2 = text2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "tx2"))?;
        let accent1 = accent1.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent1"))?;
        let accent2 = accent2.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent2"))?;
        let accent3 = accent3.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent3"))?;
        let accent4 = accent4.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent4"))?;
        let accent5 = accent5.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent5"))?;
        let accent6 = accent6.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "accent6"))?;
        let hyperlink = hyperlink.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "hlink"))?;
        let followed_hyperlink =
            followed_hyperlink.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "folHlink"))?;

        Ok(Self {
            background1,
            text1,
            background2,
            text2,
            accent1,
            accent2,
            accent3,
            accent4,
            accent5,
            accent6,
            hyperlink,
            followed_hyperlink,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ColorScheme {
    /// The common name for this color scheme. This name can show up in the user interface in
    /// a list of color schemes.
    pub name: String,

    /// This element defines a color that happens to be the dark 1 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub dark1: Color,

    /// This element defines a color that happens to be the accent 1 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub light1: Color,

    /// This element defines a color that happens to be the dark 2 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub dark2: Color,

    /// This element defines a color that happens to be the accent 1 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub light2: Color,

    /// This element defines a color that happens to be the accent 1 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent1: Color,

    /// This element defines a color that happens to be the accent 2 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent2: Color,

    /// This element defines a color that happens to be the accent 3 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent3: Color,

    /// This element defines a color that happens to be the accent 4 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent4: Color,

    /// This element defines a color that happens to be the accent 5 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent5: Color,

    /// This element defines a color that happens to be the accent 6 color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub accent6: Color,

    /// This element defines a color that happens to be the hyperlink color. The set of twelve colors come together to
    /// form the color scheme for a theme.
    pub hyperlink: Color,
    /// This element defines a color that happens to be the followed hyperlink color. The set of twelve colors come
    /// together to form the color scheme for a theme.
    pub followed_hyperlink: Color,
}

impl ColorScheme {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let name = xml_node
            .attributes
            .get("name")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?
            .clone();

        let mut dk1 = None;
        let mut lt1 = None;
        let mut dk2 = None;
        let mut lt2 = None;
        let mut accent1 = None;
        let mut accent2 = None;
        let mut accent3 = None;
        let mut accent4 = None;
        let mut accent5 = None;
        let mut accent6 = None;
        let mut hyperlink = None;
        let mut follow_hyperlink = None;

        for child_node in &xml_node.child_nodes {
            let color = child_node
                .child_nodes
                .iter()
                .find_map(Color::try_from_xml_element)
                .transpose()?
                .ok_or_else(|| MissingChildNodeError::new(child_node.name.clone(), "EG_Color"))?;

            match child_node.local_name() {
                "dk1" => dk1 = Some(color),
                "lt1" => lt1 = Some(color),
                "dk2" => dk2 = Some(color),
                "lt2" => lt2 = Some(color),
                "accent1" => accent1 = Some(color),
                "accent2" => accent2 = Some(color),
                "accent3" => accent3 = Some(color),
                "accent4" => accent4 = Some(color),
                "accent5" => accent5 = Some(color),
                "accent6" => accent6 = Some(color),
                "hlink" => hyperlink = Some(color),
                "folHlink" => follow_hyperlink = Some(color),
                _ => (),
            }
        }

        let dark1 = dk1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "dk1"))?;
        let light1 = lt1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lt1"))?;
        let dark2 = dk2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "dk2"))?;
        let light2 = lt2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lt2"))?;
        let accent1 = accent1.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent1"))?;
        let accent2 = accent2.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent2"))?;
        let accent3 = accent3.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent3"))?;
        let accent4 = accent4.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent4"))?;
        let accent5 = accent5.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent5"))?;
        let accent6 = accent6.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "accent6"))?;
        let hyperlink = hyperlink.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hlink"))?;
        let followed_hyperlink =
            follow_hyperlink.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "folHlink"))?;

        Ok(Self {
            name,
            dark1,
            light1,
            dark2,
            light2,
            accent1,
            accent2,
            accent3,
            accent4,
            accent5,
            accent6,
            hyperlink,
            followed_hyperlink,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ColorSchemeAndMapping {
    /// This element defines a set of colors which are referred to as a color scheme. The color scheme is responsible for
    /// defining a list of twelve colors. The twelve colors consist of six accent colors, two dark colors, two light colors
    /// and a color for each of a hyperlink and followed hyperlink.
    ///
    /// The Color Scheme Color elements appear in a sequence. The following listing shows the index value and
    /// corresponding Color Name.
    ///
    /// | Sequence Index        | Element (Color) Name              |
    /// |-----------------------|-----------------------------------|
    /// |0                      |dark1                              |
    /// |1                      |light1                             |
    /// |2                      |dark2                              |
    /// |3                      |light2                             |
    /// |4                      |accent1                            |
    /// |5                      |accent2                            |
    /// |6                      |accent3                            |
    /// |7                      |accent4                            |
    /// |8                      |accent5                            |
    /// |9                      |accent6                            |
    /// |10                     |hyperlink                          |
    /// |11                     |followedHyperlink                  |
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <clrScheme name="sample">
    ///   <dk1>
    ///     <sysClr val="windowText"/>
    ///   </dk1>
    ///   <lt1>
    ///     <sysClr val="window"/>
    ///   </lt1>
    ///   <dk2>
    ///     <srgbClr val="04617B"/>
    ///   </dk2>
    ///   <lt2>
    ///     <srgbClr val="DBF5F9"/>
    ///   </lt2>
    ///   <accent1>
    ///     <srgbClr val="0F6FC6"/>
    ///   </accent1>
    ///   <accent2>
    ///     <srgbClr val="009DD9"/>
    ///   </accent2>
    ///   <accent3>
    ///     <srgbClr val="0BD0D9"/>
    ///   </accent3>
    ///   <accent4>
    ///     <srgbClr val="10CF9B"/>
    ///   </accent4>
    ///   <accent5>
    ///     <srgbClr val="7CCA62"/>
    ///   </accent5>
    ///   <accent6>
    ///     <srgbClr val="A5C249"/>
    ///   </accent6>
    ///   <hlink>
    ///     <srgbClr val="FF9800"/>
    ///   </hlink>
    ///   <folHlink>
    ///     <srgbClr val="F45511"/>
    ///   </folHlink>
    /// </clrScheme>
    /// ```
    pub color_scheme: Box<ColorScheme>,

    /// This element specifics the color mapping layer which allows a user to define colors for background and text.
    /// This allows for swapping out of light/dark colors for backgrounds and the text on top of the background in order
    /// to maintain readability of the text On a deeper level, this specifies exactly which colors the first 12 values refer
    /// to in the color scheme.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
    /// accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
    /// accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    /// ```
    ///
    /// In this example, we see that bg1 is mapped to lt1, tx1 is mapped to dk1, and so on.
    pub color_mapping: Option<Box<ColorMapping>>,
}

impl ColorSchemeAndMapping {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut color_scheme = None;
        let mut color_mapping = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrScheme" => color_scheme = Some(Box::new(ColorScheme::from_xml_element(child_node)?)),
                "clrMap" => color_mapping = Some(Box::new(ColorMapping::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let color_scheme =
            color_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrScheme"))?;

        Ok(Self {
            color_scheme,
            color_mapping,
        })
    }
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct ObjectStyleDefaults {
    /// This element defines the formatting that is associated with the default shape. The default formatting can be
    /// applied to a shape when it is initially inserted into a document.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <spDef>
    ///   <spPr>
    ///     <solidFill>
    ///       <schemeClr val="accent2">
    ///         <shade val="75000"/>
    ///       </schemeClr>
    ///     </solidFill>
    ///   </spPr>
    ///   <bodyPr rtlCol="0" anchor="ctr"/>
    ///   <lstStyle>
    ///     <defPPr algn="ctr">
    ///       <defRPr/>
    ///     </defPPr>
    ///   </lstStyle>
    ///   <style>
    ///     <lnRef idx="1">
    ///       <schemeClr val="accent1"/>
    ///     </lnRef>
    ///     <fillRef idx="2">
    ///       <schemeClr val="accent1"/>
    ///     </fillRef>
    ///     <effectRef idx="1">
    ///       <schemeClr val="accent1"/>
    ///     </effectRef>
    ///     <fontRef idx="minor">
    ///       <schemeClr val="dk1"/>
    ///     </fontRef>
    ///   </style>
    /// </spDef>
    /// ```
    ///
    /// In this example, we see a default shape which references a certain themed fill, line, effect, and font along with
    /// an override fill to these.
    pub shape_definition: Option<Box<DefaultShapeDefinition>>,

    /// This element defines a default line that is used within a document.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <lnDef>
    ///   <spPr/>
    ///   <bodyPr/>
    ///   <lstStyle/>
    ///   <style>
    ///     <lnRef idx="1">
    ///       <schemeClr val="accent2"/>
    ///     </lnRef>
    ///     <fillRef idx="0">
    ///       <schemeClr val="accent2"/>
    ///     </fillRef>
    ///     <effectRef idx="0">
    ///       <schemeClr val="accent2"/>
    ///     </effectRef>
    ///     <fontRef idx="minor">
    ///       <schemeClr val="tx1"/>
    ///     </fontRef>
    ///   </style>
    /// </lnDef>
    /// ```
    ///
    /// In this example, we see that the default line for the document is being defined as a themed line which
    /// references the subtle line style with idx equal to 1.
    pub line_definition: Option<Box<DefaultShapeDefinition>>,

    /// This element defines the default formatting which is applied to text in a document by default. The default
    /// formatting can and should be applied to the shape when it is initially inserted into a document.
    ///
    /// ```xml
    /// <txDef>
    ///   <spPr>
    ///     <solidFill>
    ///       <schemeClr val="accent2">
    ///         <shade val="75000"/>
    ///       </schemeClr>
    ///     </solidFill>
    ///   </spPr>
    ///   <bodyPr rtlCol="0" anchor="ctr"/>
    ///   <lstStyle>
    ///     <defPPr algn="ctr">
    ///       <defRPr/>
    ///     </defPPr>
    ///   </lstStyle>
    ///   <style>
    ///     <lnRef idx="1">
    ///       <schemeClr val="accent1"/>
    ///     </lnRef>
    ///     <fillRef idx="2">
    ///       <schemeClr val="accent1"/>
    ///     </fillRef>
    ///     <effectRef idx="1">
    ///       <schemeClr val="accent1"/>
    ///     </effectRef>
    ///     <fontRef idx="minor">
    ///       <schemeClr val="dk1"/>
    ///     </fontRef>
    ///   </style>
    /// </txDef>
    /// ```
    pub text_definition: Option<Box<DefaultShapeDefinition>>,
}

impl ObjectStyleDefaults {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        xml_node
            .child_nodes
            .iter()
            .try_fold(Default::default(), |mut instance: Self, child_node| {
                match child_node.local_name() {
                    "spDef" => {
                        instance.shape_definition =
                            Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?))
                    }
                    "lnDef" => {
                        instance.line_definition = Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?))
                    }
                    "txDef" => {
                        instance.text_definition = Some(Box::new(DefaultShapeDefinition::from_xml_element(child_node)?))
                    }
                    _ => (),
                }

                Ok(instance)
            })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct OfficeStyleSheet {
    pub name: Option<String>,

    /// This element defines the theme formatting options for the theme and is the workhorse of the theme. This is
    /// where the bulk of the shared theme information is contained and used by a document. This element contains
    /// the color scheme, font scheme, and format scheme elements which define the different formatting aspects of
    /// what a theme defines.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <themeElements>
    ///   <clrScheme name="sample">
    ///     ...
    ///   </clrScheme>
    ///   <fontScheme name="sample">
    ///     ...
    ///   </fontScheme>
    ///   <fmtScheme name="sample">
    ///     <fillStyleLst>
    ///       ...
    ///     </fillStyleLst>
    ///     <lnStyleLst>
    ///       ...
    ///     </lnStyleLst>
    ///     <effectStyleLst>
    ///       ...
    ///     </effectStyleLst>
    ///     <bgFillStyleLst>
    ///       ...
    ///     </bgFillStyleLst>
    ///   </fmtScheme>
    /// </themeElements>
    /// ```
    ///
    /// In this example, we see the basic structure of how a theme elements is defined and have left out the true guts of
    /// each individual piece to save room. Each part (color scheme, font scheme, format scheme) is defined elsewhere
    /// within DrawingML.
    pub theme_elements: Box<BaseStyles>,

    /// This element allows for the definition of default shape, line, and textbox formatting properties. An application
    /// can use this information to format a shape (or text) initially on insertion into a document.
    pub object_defaults: Option<ObjectStyleDefaults>,

    /// This element is a container for the list of extra color schemes present in a document.
    ///
    /// An ColorSchemeAndMapping element defines an auxiliary color scheme, which includes both a color scheme and
    /// color mapping. This is mainly used for backward compatibility concerns and roundtrips information required by
    /// earlier versions.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <extraClrScheme>
    ///   <clrScheme name="extraColorSchemeSample">
    ///     <dk1>
    ///       <sysClr val="windowText"/>
    ///     </dk1>
    ///     <lt1>
    ///       <sysClr val="window"/>
    ///     </lt1>
    ///     <dk2>
    ///       <srgbClr val="04617B"/>
    ///     </dk2>
    ///     <lt2>
    ///       <srgbClr val="DBF5F9"/>
    ///     </lt2>
    ///     <accent1>
    ///       <srgbClr val="0F6FC6"/>
    ///     </accent1>
    ///     <accent2>
    ///       <srgbClr val="009DD9"/>
    ///     </accent2>
    ///     <accent3>
    ///       <srgbClr val="0BD0D9"/>
    ///     </accent3>
    ///     <accent4>
    ///       <srgbClr val="10CF9B"/>
    ///     </accent4>
    ///     <accent5>
    ///       <srgbClr val="7CCA62"/>
    ///     </accent5>
    ///     <accent6>
    ///       <srgbClr val="A5C249"/>
    ///     </accent6>
    ///     <hlink>
    ///       <srgbClr val="FF9800"/>
    ///     </hlink>
    ///     <folHlink>
    ///       <srgbClr val="F45511"/>
    ///     </folHlink>
    ///   </clrScheme>
    ///   <clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
    ///     accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"
    ///     accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    /// </extraClrScheme>
    /// ```
    pub extra_color_scheme_list: Option<Vec<ColorSchemeAndMapping>>,

    /// This element allows for a custom color palette to be created and which shows up alongside other color schemes.
    /// This can be very useful, for example, when someone would like to maintain a corporate color palette.
    pub custom_color_list: Option<Vec<CustomColor>>,
}

impl OfficeStyleSheet {
    pub fn from_zip_file(zip_file: &mut ZipFile<'_>) -> Result<Self> {
        let mut xml_string = String::new();
        zip_file.read_to_string(&mut xml_string)?;
        let xml_node = XmlNode::from_str(xml_string.as_str())?;

        Self::from_xml_element(&xml_node)
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        trace!("parsing OfficeStyleSheet '{}'", xml_node.name);
        let name = xml_node.attributes.get("name").cloned();

        let mut theme_elements = None;
        let mut object_defaults = None;
        let mut extra_color_scheme_list = None;
        let mut custom_color_list = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "themeElements" => theme_elements = Some(Box::new(BaseStyles::from_xml_element(child_node)?)),
                "objectDefaults" => object_defaults = Some(ObjectStyleDefaults::from_xml_element(child_node)?),
                "extraClrSchemeLst" => {
                    extra_color_scheme_list = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .filter(|child_node| child_node.local_name() == "extraClrScheme")
                            .map(ColorSchemeAndMapping::from_xml_element)
                            .collect::<Result<Vec<_>>>()?,
                    );
                }
                "custClrLst" => {
                    custom_color_list = Some(
                        child_node
                            .child_nodes
                            .iter()
                            .filter(|child_node| child_node.local_name() == "custClr")
                            .map(CustomColor::from_xml_element)
                            .collect::<Result<Vec<_>>>()?,
                    )
                }
                _ => (),
            }
        }

        let theme_elements =
            theme_elements.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "themeElements"))?;

        Ok(Self {
            name,
            theme_elements,
            object_defaults,
            extra_color_scheme_list,
            custom_color_list,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct BaseStyles {
    pub color_scheme: Box<ColorScheme>,

    /// This element defines the font scheme within the theme. The font scheme consists of a pair of major and minor
    /// fonts for which to use in a document. The major font corresponds well with the heading areas of a document,
    /// and the minor font corresponds well with the normal text or paragraph areas.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <fontScheme name="sample">
    ///   <majorFont>
    ///   ...
    ///   </majorFont>
    ///   <minorFont>
    ///   ...
    ///   </minorFont>
    /// </fontScheme>
    /// ```
    pub font_scheme: FontScheme,

    /// This element contains the background fill styles, effect styles, fill styles, and line styles which define the style
    /// matrix for a theme. The style matrix consists of subtle, moderate, and intense fills, lines, and effects. The
    /// background fills are not generally thought of to directly be associated with the matrix, but do play a role in the
    /// style of the overall document. Usually, a given object chooses a single line style, a single fill style, and a single
    /// effect style in order to define the overall final look of the object.
    pub format_scheme: Box<StyleMatrix>,
}

impl BaseStyles {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        trace!("parsing BaseStyles '{}'", xml_node.name);
        let mut color_scheme = None;
        let mut font_scheme = None;
        let mut format_scheme = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "clrScheme" => color_scheme = Some(Box::new(ColorScheme::from_xml_element(child_node)?)),
                "fontScheme" => font_scheme = Some(FontScheme::from_xml_element(child_node)?),
                "fmtScheme" => format_scheme = Some(Box::new(StyleMatrix::from_xml_element(child_node)?)),
                _ => (),
            }
        }

        let color_scheme =
            color_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "clrScheme"))?;
        let font_scheme = font_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fontScheme"))?;
        let format_scheme =
            format_scheme.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "fmtScheme"))?;

        Ok(Self {
            color_scheme,
            font_scheme,
            format_scheme,
        })
    }
}
