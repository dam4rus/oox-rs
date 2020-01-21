use crate::error::{AdjustParseError, ParseHexColorRGBError, StringLengthMismatch};
use std::str::FromStr;

/// This simple type specifies that its values shall be a 128-bit globally unique identifier (GUID) value.
///
/// This simple type's contents shall match the following regular expression pattern:
/// \{[0-9A-F]{8}-[0-9AF]{/// 4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}.
pub type Guid = String;

/// This simple type specifies that its contents will contain a percentage value. See the union's member types for
/// details.
pub type Percentage = f32;

/// This simple type specifies that its contents will contain a positive percentage value. See the union's member
/// types for details.
pub type PositivePercentage = f32;

/// This simple type specifies that its contents will contain a positive percentage value from zero through one
/// hundred percent.
///
/// Values represented by this type are restricted to: 0 <= n <= 100000
pub type PositiveFixedPercentage = f32;

/// This simple type represents a fixed percentage from negative one hundred to positive one hundred percent. See
/// the union's member types for details.
///
/// Values represented by this type are restricted to: -100000 <= n <= 100000
pub type FixedPercentage = f32;

/// This simple type specifies that its contents shall contain a color value in RRGGBB hexadecimal format, specified
/// using six hexadecimal digits. Each of the red, green, and blue color values, from 0-255, is encoded as two
/// hexadecimal digits.
///
/// # Example
/// Consider a color defined as follows:
///
/// Red:   122
/// Green:  23
/// Blue:  209
///
/// The resulting RRGGBB value would be 7A17D1, as each color is transformed into its hexadecimal equivalent.
pub type HexColorRGB = [u8; 3];

pub fn parse_hex_color_rgb(s: &str) -> Result<HexColorRGB, ParseHexColorRGBError> {
    match s.len() {
        6 => Ok([
            u8::from_str_radix(&s[0..2], 16)?,
            u8::from_str_radix(&s[2..4], 16)?,
            u8::from_str_radix(&s[4..6], 16)?,
        ]),
        len => Err(ParseHexColorRGBError::InvalidLength(StringLengthMismatch {
            required: 6,
            provided: len,
        })),
    }
}

/// This simple type represents a one dimensional position or length as either:
///
/// * EMUs.
/// * A number followed immediately by a unit identifier.
pub type Coordinate = i64;

/// This simple type represents a positive position or length in EMUs.
pub type PositiveCoordinate = u64;

/// This simple type specifies a coordinate within the document. This can be used for measurements or spacing; its
/// maximum size is 2147483647 EMUs.
///
/// Its contents can contain either:
///
/// * A whole number, whose contents consist of a measurement in EMUs (English Metric Units)
/// * A number immediately followed by a unit identifier
pub type Coordinate32 = i32;

/// This simple type specifies the a positive coordinate point that has a maximum size of 32 bits.
///
/// The units of measurement used here are EMUs (English Metric Units).
pub type PositiveCoordinate32 = u32;

/// This simple type specifies the width of a line in EMUs. 1 pt = 12700 EMUs
///
/// Values represented by this type are restricted to: 0 <= n <= 20116800
pub type LineWidth = Coordinate32;

/// This simple type specifies a unique integer identifier for each drawing element.
pub type DrawingElementId = u32;

/// This simple type represents an angle in 60,000ths of a degree. Positive angles are clockwise (i.e., towards the
/// positive y axis); negative angles are counter-clockwise (i.e., towards the negative y axis).
pub type Angle = i32;

/// This simple type represents a fixed range angle in 60000ths of a degree. Range from (-90, 90 degrees).
///
/// Values represented by this type are restricted to: -5400000 <= n <= 5400000
pub type FixedAngle = Angle;

/// This simple type represents a positive angle in 60000ths of a degree. Range from [0, 360 degrees).
///
/// Values represented by this type are restricted to: 0 <= n <= 21600000
pub type PositiveFixedAngle = Angle;

/// This simple type specifies a geometry guide name.
pub type GeomGuideName = String;

/// This simple type specifies a geometry guide formula.
pub type GeomGuideFormula = String;

/// This simple type specifies an index into one of the lists in the style matrix specified by the
/// BaseStyles::format_scheme element (StyleMatrix::bg_fill_style_list, StyleMatrix::effect_style_list,
/// StyleMatrix::fill_style_list, or StyleMatrix::line_style_list).
pub type StyleMatrixColumnIndex = u32;

/// This simple type specifies the number of columns.
///
/// Values represented by this type are restricted to: 1 <= n <= 16
pub type TextColumnCount = i32;

/// Values represented by this type are restricted to: 1000 <= n <= 100000
pub type TextFontScalePercent = Percentage;

/// Values represented by this type are restricted to: 0 <= n <= 13200000
pub type TextSpacingPercent = Percentage;

/// This simple type specifies the Text Spacing that is used in terms of font point size.
///
/// Values represented by this type are restricted to: 0 <= n <= 158400
pub type TextSpacingPoint = i32;

/// This simple type specifies the margin that is used and its corresponding size.
///
/// Values represented by this type are restricted to: 0 <= n <= 51206400
pub type TextMargin = Coordinate32;

/// This simple type specifies the text indentation amount to be used.
///
/// Values represented by this type are restricted to: -51206400 <= n <= 51206400
pub type TextIndent = Coordinate32;

/// This simple type specifies the indent level type. We support list level 0 to 8, and we use -1 and -2 for outline
/// mode levels that should only exist in memory.
///
/// Values represented by this type are restricted to: 0 <= n <= 8
pub type TextIndentLevelType = i32;

/// This simple type specifies the range that the bullet percent can be. A bullet percent is the size of the bullet with
/// respect to the text that should follow it.
///
/// Values represented by this type are restricted to: 25000 <= n <= 400000
pub type TextBulletSizePercent = Percentage;

/// This simple type specifies the size of any text in hundredths of a point. Shall be at least 1 point.
///
/// Values represented by this type are restricted to: 100 <= n <= 400000
pub type TextFontSize = i32;

/// This simple type specifies the way we represent a font typeface.
pub type TextTypeFace = String;

/// Specifies a language tag as defined by RFC 3066. See simple type for additional information.
pub type TextLanguageID = String;

/// This simple type specifies a number consisting of 20 hexadecimal digits which defines the Panose-1 font
/// classification.
///
/// This simple type's contents have a length of exactly 20 hexadecimal digit(s).
///
/// # Xml example
///
/// ```xml
/// <w:font w:name="Times New Roman">
///   <w:panose1 w:val="02020603050405020304" />
///   …
/// </w:font>
/// ```
pub type Panose = String; // TODO: hex, length=10

/// This simple type specifies the range that the start at number for a bullet's auto-numbering sequence can begin
/// at. When the numbering is alphabetical, then the numbers map to the appropriate letter. 1->a, 2->b, etc. If the
/// numbers go above 26, then the numbers begin to double up. For example, 27->aa and 53->aaa.
///
/// Values represented by this type are restricted to: 1 <= n <= 32767
pub type TextBulletStartAtNum = i32;

/// This simple type specifies that its contents contains a language identifier as defined by RFC 4646/BCP 47.
///
/// The contents of this language are interpreted based on the context of the parent XML element.
///
/// # Xml example
///
/// ```xml
/// <w:lang w:val="en-CA" />
/// ```
///
/// This language is therefore specified as English (en) and Canada (CA), resulting in use of the English (Canada)
/// language setting.
pub type Lang = String;

/// This simple type specifies a non-negative font size in hundredths of a point.
///
/// Values represented by this type are restricted to: 0 <= n <= 400000
pub type TextNonNegativePoint = i32;

/// This simple type specifies a coordinate within the document. This can be used for measurements or spacing
///
/// Values represented by this type are restricted to: -400000 <= n <= 400000
pub type TextPoint = i32;

/// Specifies the shape ID for legacy shape identification purposes.
pub type ShapeId = String;

/// This simple type is an adjustable coordinate is either an absolute coordinate position or a reference to a
/// geometry guide.
#[derive(Debug, Clone, PartialEq)]
pub enum AdjCoordinate {
    Coordinate(Coordinate),
    GeomGuideName(GeomGuideName),
}

impl FromStr for AdjCoordinate {
    type Err = AdjustParseError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s.parse::<Coordinate>() {
            Ok(coord) => Ok(AdjCoordinate::Coordinate(coord)),
            Err(_) => Ok(AdjCoordinate::GeomGuideName(GeomGuideName::from(s))),
        }
    }
}

/// This simple type is an adjustable angle, either an absolute angle or a reference to a geometry guide. The units
/// for an adjustable angle are 60,000ths of a degree.
#[derive(Debug, Clone, PartialEq)]
pub enum AdjAngle {
    Angle(Angle),
    GeomGuideName(GeomGuideName),
}

impl FromStr for AdjAngle {
    type Err = AdjustParseError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s.parse::<Angle>() {
            Ok(angle) => Ok(AdjAngle::Angle(angle)),
            Err(_) => Ok(AdjAngle::GeomGuideName(GeomGuideName::from(s))),
        }
    }
}

/// This simple type indicates whether/how to flip the contents of a tile region when using it to fill a larger fill
/// region.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TileFlipMode {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "x")]
    X,
    #[strum(serialize = "y")]
    Y,
    #[strum(serialize = "xy")]
    XY,
}

/// This simple type describes how to position two rectangles relative to each other.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum RectAlignment {
    #[strum(serialize = "l")]
    Left,
    #[strum(serialize = "t")]
    Top,
    #[strum(serialize = "r")]
    Right,
    #[strum(serialize = "b")]
    Bottom,
    #[strum(serialize = "tl")]
    TopLeft,
    #[strum(serialize = "tr")]
    TopRight,
    #[strum(serialize = "bl")]
    BottomLeft,
    #[strum(serialize = "br")]
    BottomRight,
    #[strum(serialize = "ctr")]
    Center,
}

/// This simple type specifies the manner in which a path should be filled. The lightening and darkening of a path
/// allow for certain parts of the shape to be colored lighter of darker depending on user preference.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PathFillMode {
    /// This specifies that the corresponding path should have no fill.
    #[strum(serialize = "none")]
    None,
    /// This specifies that the corresponding path should have a normally shaded color applied to it’s fill.
    #[strum(serialize = "norm")]
    Norm,
    /// This specifies that the corresponding path should have a lightly shaded color applied to it’s fill.
    #[strum(serialize = "lighten")]
    Lighten,
    /// This specifies that the corresponding path should have a slightly lighter shaded color applied to it’s fill.
    #[strum(serialize = "lightenLess")]
    LightenLess,
    /// This specifies that the corresponding path should have a darker shaded color applied to it’s fill.
    #[strum(serialize = "darken")]
    Darken,
    /// This specifies that the corresponding path should have a slightly darker shaded color applied to it’s fill.
    #[strum(serialize = "darkenLess")]
    DarkenLess,
}

/// This simple type specifies the preset shape geometry that is to be used for a shape. An enumeration of this
/// simple type is used so that a custom geometry does not have to be specified but instead can be constructed
/// automatically by the generating application. For each enumeration listed there is also the corresponding
/// DrawingML code that would be used to construct this shape were it a custom geometry. Within the construction
/// code for each of these preset shapes there are predefined guides that the generating application shall maintain
/// for calculation purposes at all times. The necessary guides should have the following values:
///
/// * **3/4 of a Circle ('3cd4') - Constant value of "16200000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 270 degrees.
///
/// * **3/8 of a Circle ('3cd8') - Constant value of "8100000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 135 degrees.
///
/// * **5/8 of a Circle ('5cd8') - Constant value of "13500000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 225 degrees.
///
/// * **7/8 of a Circle ('7cd8') - Constant value of "18900000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 315 degrees.
///
/// * **Shape Bottom Edge ('b') - Constant value of "h"**
///
///     This is the bottom edge of the shape and since the top edge of the shape is considered the 0 point, the
///     bottom edge is thus the shape height.
///
/// * **1/2 of a Circle ('cd2') - Constant value of "10800000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 180 degrees.
///
/// * **1/4 of a Circle ('cd4') - Constant value of "5400000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 90 degrees.
///
/// * **1/8 of a Circle ('cd8') - Constant value of "2700000.0"**
///
///     The units here are in 60,000ths of a degree. This is equivalent to 45 degrees.
///
/// * **Shape Height ('h')**
///
///     This is the variable height of the shape defined in the shape properties. This value is received from the shape
///     transform listed within the <spPr> element.
///
/// * **Horizontal Center ('hc') - Calculated value of "\*/ w 1.0 2.0"**
///
///     This is the horizontal center of the shape which is just the width divided by 2.
///
/// * **1/2 of Shape Height ('hd2') - Calculated value of "\*/ h 1.0 2.0"**
///
///     This is 1/2 the shape height.
///
/// * **1/4 of Shape Height ('hd4') - Calculated value of "\*/ h 1.0 4.0"**
///
///     This is 1/4 the shape height.
///
/// * **1/5 of Shape Height ('hd5') - Calculated value of "\*/ h 1.0 5.0"**
///
///     This is 1/5 the shape height.
///
/// * **1/6 of Shape Height ('hd6') - Calculated value of "\*/ h 1.0 6.0"**
///
///     This is 1/6 the shape height.
///
/// * **1/8 of Shape Height ('hd8') - Calculated value of "\*/ h 1.0 8.0"**
///
///     This is 1/8 the shape height.
///
/// * **Shape Left Edge ('l') - Constant value of "0"**
///
///     This is the left edge of the shape and the left edge of the shape is considered the horizontal 0 point.
///
/// * **Longest Side of Shape ('ls') - Calculated value of "max w h"**
///
///     This is the longest side of the shape. This value is either the width or the height depending on which is greater.
///
/// * **Shape Right Edge ('r') - Constant value of "w"**
///
///     This is the right edge of the shape and since the left edge of the shape is considered the 0 point, the right edge
///     is thus the shape width.
///
/// * **Shortest Side of Shape ('ss') - Calculated value of "min w h"**
///
///     This is the shortest side of the shape. This value is either the width or the height depending on which is
///     smaller.
///
/// * **1/2 Shortest Side of Shape ('ssd2') - Calculated value of "\*/ ss 1.0 2.0"**
///
///     This is 1/2 the shortest side of the shape.
///
/// * **1/4 Shortest Side of Shape ('ssd4') - Calculated value of "\*/ ss 1.0 4.0"**
///
///     This is 1/4 the shortest side of the shape.
///
/// * **1/6 Shortest Side of Shape ('ssd6') - Calculated value of "\*/ ss 1.0 6.0"**
///
///     This is 1/6 the shortest side of the shape.
///
/// * **1/8 Shortest Side of Shape ('ssd8') - Calculated value of "\*/ ss 1.0 8.0"**
///
///     This is 1/8 the shortest side of the shape.
///
/// * **Shape Top Edge ('t') - Constant value of "0"**
///
///     This is the top edge of the shape and the top edge of the shape is considered the vertical 0 point.
///
/// * **Vertical Center of Shape ('vc') - Calculated value of "\*/ h 1.0 2.0"**
///
///     This is the vertical center of the shape which is just the height divided by 2.
///
/// * **Shape Width ('w')**
///
///     This is the variable width of the shape defined in the shape properties. This value is received from the shape
///     transform listed within the <spPr> element.
///
/// * **1/2 of Shape Width ('wd2') - Calculated value of "\*/ w 1.0 2.0"**
///
///     This is 1/2 the shape width.
///
/// * **1/4 of Shape Width ('wd4') - Calculated value of "\*/ w 1.0 4.0"**
///
///     This is 1/4 the shape width.
///
/// * **1/5 of Shape Width ('wd5') - Calculated value of "\*/ w 1.0 5.0"**
///
///     This is 1/5 the shape width.
///
/// * **1/6 of Shape Width ('wd6') - Calculated value of "\*/ w 1.0 6.0"**
///
///     This is 1/6 the shape width.
///
/// * **1/8 of Shape Width ('wd8') - Calculated value of "\*/ w 1.0 8.0"**
///
///     This is 1/8 the shape width.
///
/// * **1/10 of Shape Width ('wd10') - Calculated value of "\*/ w 1.0 10.0"**
///
///     This is 1/10 the shape width.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum ShapeType {
    #[strum(serialize = "line")]
    Line,
    #[strum(serialize = "lineInv")]
    LineInverse,
    #[strum(serialize = "triangle")]
    Triangle,
    #[strum(serialize = "rtTriangle")]
    RightTriangle,
    #[strum(serialize = "rect")]
    Rect,
    #[strum(serialize = "diamond")]
    Diamond,
    #[strum(serialize = "parallelogram")]
    Parallelogram,
    #[strum(serialize = "trapezoid")]
    Trapezoid,
    #[strum(serialize = "nonIsoscelesTrapezoid")]
    NonIsoscelesTrapezoid,
    #[strum(serialize = "pentagon")]
    Pentagon,
    #[strum(serialize = "hexagon")]
    Hexagon,
    #[strum(serialize = "heptagon")]
    Heptagon,
    #[strum(serialize = "octagon")]
    Octagon,
    #[strum(serialize = "decagon")]
    Decagon,
    #[strum(serialize = "dodecagon")]
    Dodecagon,
    #[strum(serialize = "star4")]
    Star4,
    #[strum(serialize = "star5")]
    Star5,
    #[strum(serialize = "star6")]
    Star6,
    #[strum(serialize = "star7")]
    Star7,
    #[strum(serialize = "star8")]
    Star8,
    #[strum(serialize = "star10")]
    Star10,
    #[strum(serialize = "star12")]
    Star12,
    #[strum(serialize = "star16")]
    Star16,
    #[strum(serialize = "star24")]
    Star24,
    #[strum(serialize = "star32")]
    Star32,
    #[strum(serialize = "roundRect")]
    RoundRect,
    #[strum(serialize = "round1Rect")]
    Round1Rect,
    #[strum(serialize = "round2SameRect")]
    Round2SameRect,
    #[strum(serialize = "round2DiagRect")]
    Round2DiagRect,
    #[strum(serialize = "snipRoundRect")]
    SnipRoundRect,
    #[strum(serialize = "snip1Rect")]
    Snip1Rect,
    #[strum(serialize = "snip2SameRect")]
    Snip2SameRect,
    #[strum(serialize = "snip2DiagRect")]
    Snip2DiagRect,
    #[strum(serialize = "plaque")]
    Plaque,
    #[strum(serialize = "ellipse")]
    Ellipse,
    #[strum(serialize = "teardrop")]
    Teardrop,
    #[strum(serialize = "homePlate")]
    HomePlate,
    #[strum(serialize = "chevron")]
    Chevron,
    #[strum(serialize = "pieWedge")]
    PieWedge,
    #[strum(serialize = "pie")]
    Pie,
    #[strum(serialize = "blockArc")]
    BlockArc,
    #[strum(serialize = "donut")]
    Donut,
    #[strum(serialize = "noSmoking")]
    NoSmoking,
    #[strum(serialize = "rightArrow")]
    RightArrow,
    #[strum(serialize = "leftArrow")]
    LeftArrow,
    #[strum(serialize = "upArrow")]
    UpArrow,
    #[strum(serialize = "downArrow")]
    DownArrow,
    #[strum(serialize = "stripedRightArrow")]
    StripedRightArrow,
    #[strum(serialize = "notchedRightArrow")]
    NotchedRightArrow,
    #[strum(serialize = "bentUpArrow")]
    BentUpArrow,
    #[strum(serialize = "leftRightArrow")]
    LeftRightArrow,
    #[strum(serialize = "upDownArrow")]
    UpDownArrow,
    #[strum(serialize = "leftUpArrow")]
    LeftUpArrow,
    #[strum(serialize = "leftRightUpArrow")]
    LeftRightUpArrow,
    #[strum(serialize = "quadArrow")]
    QuadArrow,
    #[strum(serialize = "leftArrowCallout")]
    LeftArrowCallout,
    #[strum(serialize = "rightArrowCallout")]
    RightArrowCallout,
    #[strum(serialize = "upArrowCallout")]
    UpArrowCallout,
    #[strum(serialize = "downArrowCallout")]
    DownArrowCallout,
    #[strum(serialize = "leftRightArrowCallout")]
    LeftRightArrowCallout,
    #[strum(serialize = "upDownArrowCallout")]
    UpDownArrowCallout,
    #[strum(serialize = "quadArrowCallout")]
    QuadArrowCallout,
    #[strum(serialize = "bentArrow")]
    BentArrow,
    #[strum(serialize = "uturnArrow")]
    UturnArrow,
    #[strum(serialize = "circularArrow")]
    CircularArrow,
    #[strum(serialize = "leftCircularArrow")]
    LeftCircularArrow,
    #[strum(serialize = "leftRightCircularArrow")]
    LeftRightCircularArrow,
    #[strum(serialize = "curvedRightArrow")]
    CurvedRightArrow,
    #[strum(serialize = "curvedLeftArrow")]
    CurvedLeftArrow,
    #[strum(serialize = "curvedUpArrow")]
    CurvedUpArrow,
    #[strum(serialize = "curvedDownArrow")]
    CurvedDownArrow,
    #[strum(serialize = "swooshArrow")]
    SwooshArrow,
    #[strum(serialize = "cube")]
    Cube,
    #[strum(serialize = "can")]
    Can,
    #[strum(serialize = "lightningBolt")]
    LightningBolt,
    #[strum(serialize = "heart")]
    Heart,
    #[strum(serialize = "sun")]
    Sun,
    #[strum(serialize = "moon")]
    Moon,
    #[strum(serialize = "smileyFace")]
    SmileyFace,
    #[strum(serialize = "irregularSeal1")]
    IrregularSeal1,
    #[strum(serialize = "irregularSeal2")]
    IrregularSeal2,
    #[strum(serialize = "foldedCorner")]
    FoldedCorner,
    #[strum(serialize = "bevel")]
    Bevel,
    #[strum(serialize = "frame")]
    Frame,
    #[strum(serialize = "halfFrame")]
    HalfFrame,
    #[strum(serialize = "corner")]
    Corner,
    #[strum(serialize = "diagStripe")]
    DiagStripe,
    #[strum(serialize = "chord")]
    Chord,
    #[strum(serialize = "arc")]
    Arc,
    #[strum(serialize = "leftBracket")]
    LeftBracket,
    #[strum(serialize = "rightBracket")]
    RightBracket,
    #[strum(serialize = "leftBrace")]
    LeftBrace,
    #[strum(serialize = "rightBrace")]
    RightBrace,
    #[strum(serialize = "bracketPair")]
    BracketPair,
    #[strum(serialize = "bracePair")]
    BracePair,
    #[strum(serialize = "straightConnector1")]
    StraightConnector1,
    #[strum(serialize = "bentConnector2")]
    BentConnector2,
    #[strum(serialize = "bentConnector3")]
    BentConnector3,
    #[strum(serialize = "bentConnector4")]
    BentConnector4,
    #[strum(serialize = "bentConnector5")]
    BentConnector5,
    #[strum(serialize = "curvedConnector2")]
    CurvedConnector2,
    #[strum(serialize = "curvedConnector3")]
    CurvedConnector3,
    #[strum(serialize = "curvedConnector4")]
    CurvedConnector4,
    #[strum(serialize = "curvedConnector5")]
    CurvedConnector5,
    #[strum(serialize = "callout1")]
    Callout1,
    #[strum(serialize = "callout2")]
    Callout2,
    #[strum(serialize = "callout3")]
    Callout3,
    #[strum(serialize = "accentCallout1")]
    AccentCallout1,
    #[strum(serialize = "accentCallout2")]
    AccentCallout2,
    #[strum(serialize = "accentCallout3")]
    AccentCallout3,
    #[strum(serialize = "borderCallout1")]
    BorderCallout1,
    #[strum(serialize = "borderCallout2")]
    BorderCallout2,
    #[strum(serialize = "borderCallout3")]
    BorderCallout3,
    #[strum(serialize = "accentBorderCallout1")]
    AccentBorderCallout1,
    #[strum(serialize = "accentBorderCallout2")]
    AccentBorderCallout2,
    #[strum(serialize = "accentBorderCallout3")]
    AccentBorderCallout3,
    #[strum(serialize = "wedgeRectCallout")]
    WedgeRectCallout,
    #[strum(serialize = "wedgeRoundRectCallout")]
    WedgeRoundRectCallout,
    #[strum(serialize = "wedgeEllipseCallout")]
    WedgeEllipseCallout,
    #[strum(serialize = "cloudCallout")]
    CloudCallout,
    #[strum(serialize = "cloud")]
    Cloud,
    #[strum(serialize = "ribbon")]
    Ribbon,
    #[strum(serialize = "ribbon2")]
    Ribbon2,
    #[strum(serialize = "ellipseRibbon")]
    EllipseRibbon,
    #[strum(serialize = "ellipseRibbon2")]
    EllipseRibbon2,
    #[strum(serialize = "leftRightRibbon")]
    LeftRightRibbon,
    #[strum(serialize = "verticalScroll")]
    VerticalScroll,
    #[strum(serialize = "horizontalScroll")]
    HorizontalScroll,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "doubleWave")]
    DoubleWave,
    #[strum(serialize = "plus")]
    Plus,
    #[strum(serialize = "flowChartProcess")]
    FlowChartProcess,
    #[strum(serialize = "flowChartDecision")]
    FlowChartDecision,
    #[strum(serialize = "flowChartInputOutput")]
    FlowChartInputOutput,
    #[strum(serialize = "flowChartPredefinedProcess")]
    FlowChartPredefinedProcess,
    #[strum(serialize = "flowChartInternalStorage")]
    FlowChartInternalStorage,
    #[strum(serialize = "flowChartDocument")]
    FlowChartDocument,
    #[strum(serialize = "flowChartMultidocument")]
    FlowChartMultidocument,
    #[strum(serialize = "flowChartTerminator")]
    FlowChartTerminator,
    #[strum(serialize = "flowChartPreparation")]
    FlowChartPreparation,
    #[strum(serialize = "flowChartManualInput")]
    FlowChartManualInput,
    #[strum(serialize = "flowChartOperation")]
    FlowChartManualOperation,
    #[strum(serialize = "flowChartConnector")]
    FlowChartConnector,
    #[strum(serialize = "flowChartPunchedCard")]
    FlowChartPunchedCard,
    #[strum(serialize = "flowChartPunchedTape")]
    FlowChartPunchedTape,
    #[strum(serialize = "flowChartSummingJunction")]
    FlowChartSummingJunction,
    #[strum(serialize = "flowChartOr")]
    FlowChartOr,
    #[strum(serialize = "flowChartCollate")]
    FlowChartCollate,
    #[strum(serialize = "flowChartSort")]
    FlowChartSort,
    #[strum(serialize = "flowChartExtract")]
    FlowChartExtract,
    #[strum(serialize = "flowChartMerge")]
    FlowChartMerge,
    #[strum(serialize = "flowChartOfflineStorage")]
    FlowChartOfflineStorage,
    #[strum(serialize = "flowChartOnlineStorage")]
    FlowChartOnlineStorage,
    #[strum(serialize = "flowChartMagneticTape")]
    FlowChartMagneticTape,
    #[strum(serialize = "flowChartMagneticDisk")]
    FlowChartMagneticDisk,
    #[strum(serialize = "flowChartMagneticDrum")]
    FlowChartMagneticDrum,
    #[strum(serialize = "flowChartDisplay")]
    FlowChartDisplay,
    #[strum(serialize = "flowChartDelay")]
    FlowChartDelay,
    #[strum(serialize = "flowChartAlternateProcess")]
    FlowChartAlternateProcess,
    #[strum(serialize = "flowChartOffpageConnector")]
    FlowChartOffpageConnector,
    #[strum(serialize = "actionButtonBlank")]
    ActionButtonBlank,
    #[strum(serialize = "actionButtonHome")]
    ActionButtonHome,
    #[strum(serialize = "actionButtonHelp")]
    ActionButtonHelp,
    #[strum(serialize = "actionButtonInformation")]
    ActionButtonInformation,
    #[strum(serialize = "actionButtonForwardNext")]
    ActionButtonForwardNext,
    #[strum(serialize = "actionButtonBackPrevious")]
    ActionButtonBackPrevious,
    #[strum(serialize = "actionButtonEnd")]
    ActionButtonEnd,
    #[strum(serialize = "actionButtonBeginning")]
    ActionButtonBeginning,
    #[strum(serialize = "actionButtonReturn")]
    ActionButtonReturn,
    #[strum(serialize = "actionButtonDocument")]
    ActionButtonDocument,
    #[strum(serialize = "actionButtonSound")]
    ActionButtonSound,
    #[strum(serialize = "actionButtonMovie")]
    ActionButtonMovie,
    #[strum(serialize = "gear6")]
    Gear6,
    #[strum(serialize = "gear9")]
    Gear9,
    #[strum(serialize = "funnel")]
    Funnel,
    #[strum(serialize = "mathPlus")]
    MathPlus,
    #[strum(serialize = "mathMinus")]
    MathMinus,
    #[strum(serialize = "mathMultiply")]
    MathMultiply,
    #[strum(serialize = "mathDivide")]
    MathDivide,
    #[strum(serialize = "mathEqual")]
    MathEqual,
    #[strum(serialize = "mathNotEqual")]
    MathNotEqual,
    #[strum(serialize = "cornerTabs")]
    CornerTabs,
    #[strum(serialize = "squareTabs")]
    SquareTabs,
    #[strum(serialize = "plaqueTabs")]
    PlaqueTabs,
    #[strum(serialize = "chartX")]
    ChartX,
    #[strum(serialize = "chartStar")]
    ChartStar,
    #[strum(serialize = "chartPlus")]
    ChartPlus,
}

/// This simple type specifies how to cap the ends of lines. This also affects the ends of line segments for dashed
/// lines.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum LineCap {
    /// Rounded ends. Semi-circle protrudes by half line width.
    #[strum(serialize = "rnd")]
    Round,
    /// Square protrudes by half line width.
    #[strum(serialize = "sq")]
    Square,
    /// Line ends at end point.
    #[strum(serialize = "flat")]
    Flat,
}

/// This simple type specifies the compound line type that is to be used for lines with text such as underlines.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum CompoundLine {
    /// Single line: one normal width
    #[strum(serialize = "sng")]
    Single,
    /// Double lines of equal width
    #[strum(serialize = "dbl")]
    Double,
    /// Double lines: one thick, one thin
    #[strum(serialize = "thickThin")]
    ThickThin,
    /// Double lines: one thin, one thick
    #[strum(serialize = "thinThick")]
    ThinThick,
    /// Three lines: thin, thick, thin
    #[strum(serialize = "tri")]
    Triple,
}

/// This simple type specifies the Pen Alignment type for use within a text body.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PenAlignment {
    /// Center pen (line drawn at center of path stroke).
    #[strum(serialize = "ctr")]
    Center,
    /// Inset pen (the pen is aligned on the inside of the edge of the path).
    #[strum(serialize = "in")]
    Inset,
}

/// This simple type represents preset line dash values. The description for each style shows an illustration of the
/// line style. Each style also contains a precise binary representation of the repeating dash style. Each 1
/// corresponds to a line segment of the same length as the line width, and each 0 corresponds to a space of the
/// same length as the line width.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PresetLineDashVal {
    /// 1
    #[strum(serialize = "solid")]
    Solid,
    /// 1000
    #[strum(serialize = "dot")]
    Dot,
    /// 1111000
    #[strum(serialize = "dash")]
    Dash,
    /// 11111111000
    #[strum(serialize = "lgDash")]
    LargeDash,
    /// 11110001000
    #[strum(serialize = "dashDot")]
    DashDot,
    /// 111111110001000
    #[strum(serialize = "lgDashDot")]
    LargeDashDot,
    /// 1111111100010001000
    #[strum(serialize = "ldDashDotDot")]
    LargeDashDotDot,
    /// 1110
    #[strum(serialize = "sysDash")]
    SystemDash,
    /// 10
    #[strum(serialize = "sysDot")]
    SystemDot,
    /// 111010
    #[strum(serialize = "sysDashDot")]
    SystemDashDot,
    /// 11101010
    #[strum(serialize = "sysDashDotDot")]
    SystemDashDotDot,
}

/// This simple type represents the shape decoration that appears at the ends of lines. For example, one choice is an
/// arrow head.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum LineEndType {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "triangle")]
    Triangle,
    #[strum(serialize = "stealth")]
    Stealth,
    #[strum(serialize = "diamond")]
    Diamond,
    #[strum(serialize = "oval")]
    Oval,
    #[strum(serialize = "arrow")]
    Arrow,
}

/// This simple type represents the width of the line end decoration (e.g., arrowhead) relative to the width of the
/// line itself.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum LineEndWidth {
    #[strum(serialize = "sm")]
    Small,
    #[strum(serialize = "med")]
    Medium,
    #[strum(serialize = "lg")]
    Large,
}

/// This simple type represents the length of the line end decoration (e.g., arrowhead) relative to the width of the
/// line itself.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum LineEndLength {
    #[strum(serialize = "sm")]
    Small,
    #[strum(serialize = "med")]
    Medium,
    #[strum(serialize = "lg")]
    Large,
}

/// This simple type indicates one of 20 preset shadow types. Each enumeration value description illustrates the
/// type of shadow represented by the value. Each description contains the parameters to the outer shadow effect
/// represented by the preset, in addition to those attributes common to all prstShdw effects.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PresetShadowVal {
    /// No additional attributes specified.
    #[strum(serialize = "shdw1")]
    TopLeftDropShadow,
    /// No additional attributes specified.
    #[strum(serialize = "shdw2")]
    TopRightDropShadow,
    /// align = "b"
    /// ky = 40.89°
    /// sy = 50%
    #[strum(serialize = "shdw3")]
    BackLeftPerspectiveShadow,
    /// align = "b"
    /// kx = -40.89°
    /// sy = 50%
    #[strum(serialize = "shdw4")]
    BackRightPerspectiveShadow,
    /// No additional attributes specified.
    #[strum(serialize = "shdw5")]
    BottomLeftDropShadow,
    /// No additional attributes specified.
    #[strum(serialize = "shdw6")]
    BottomRightDropShadow,
    /// align = "b"
    /// kx = 40.89°
    /// sy = -50%
    #[strum(serialize = "shdw7")]
    FrontLeftPerspectiveShadow,
    /// align = "b"
    /// kx = -40.89°
    /// sy = -50%
    #[strum(serialize = "shdw8")]
    FrontRightPerspectiveShadow,
    /// align = "tl"
    /// sx = 75%
    /// sy = 75%
    #[strum(serialize = "shdw9")]
    TopLeftSmallDropShadow,
    /// align = "br"
    /// sx = 125%
    /// sy = 125%
    #[strum(serialize = "shdw10")]
    TopLeftLargeDropShadow,
    /// align = "b"
    /// kx = 40.89°
    /// sy = 50%
    #[strum(serialize = "shdw11")]
    BackLeftLongPerspectiveShadow,
    /// align = "b"
    /// kx = -40.89°
    /// sy = 50%
    #[strum(serialize = "shdw12")]
    BackRightLongPerspectiveShadow,
    /// Equivalent to two outer shadow effects.
    ///
    /// Shadow 1:
    /// No additional attributes specified.
    ///
    /// Shadow 2:
    /// color = min(1, shadow 1's color (0 <= r, g, b <= 1) +
    /// 102/255), per r, g, b component
    /// dist = 2 * shadow 1's distance
    #[strum(serialize = "shdw13")]
    TopLeftDoubleDropShadow,
    /// No additional attributes specified.
    #[strum(serialize = "shdw14")]
    BottomRightSmallDropShadow,
    /// align = "b"
    /// kx = 40.89°
    /// sy = -50%
    #[strum(serialize = "shdw15")]
    FrontLeftLongPerspectiveShadow,
    /// align = "b"
    /// kx = -40.89°
    /// sy = -50%
    #[strum(serialize = "shdw16")]
    FrontRightLongPerspectiveShadow,
    /// Equivalent to two outer shadow effects.
    ///
    /// Shadow 1:
    /// No additional attributes specified.
    ///
    /// Shadow 2:
    /// color = min(1, shadow 1's color (0 <= r, g, b <= 1) +
    /// 102/255), per r, g, b component
    /// dir = shadow 1's direction + 180°
    #[strum(serialize = "shdw17")]
    ThreeDOuterBoxShadow,
    /// Equivalent to two outer shadow effects.
    ///
    /// Shadow 1:
    /// No additional attributes specified.
    ///
    /// Shadow 2:
    /// color = min(1, shadow 1's color (0 <= r, g, b <= 1) +
    /// 102/255), per r, g, b component
    /// dir = shadow 1's direction + 180°
    #[strum(serialize = "shdw18")]
    ThreeDInnerBoxShadow,
    /// align = "b"
    /// sy = 50°
    #[strum(serialize = "shdw19")]
    BackCenterPerspectiveShadow,
    /// align = "b"
    /// sy = -100°
    #[strum(serialize = "shdw20")]
    FrontBottomShadow,
}

/// This simple type determines the relationship between effects in a container, either sibling or tree.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum EffectContainerType {
    /// Each effect is separately applied to the parent object.
    ///
    /// # Example
    ///
    /// If the parent element contains an outer shadow and a reflection, the resulting effect is a
    /// shadow around the parent object and a reflection of the object. The reflection does not have a shadow.
    #[strum(serialize = "sib")]
    Sib,
    /// Each effect is applied to the result of the previous effect.
    ///
    /// # Example
    ///
    /// If the parent element contains an outer shadow followed by a glow, the shadow is first applied
    /// to the parent object. Then, the glow is applied to the shadow (rather than the original object). The resulting
    /// effect would be a glowing shadow.
    #[strum(serialize = "tree")]
    Tree,
}

/// This simple type represents one of the fonts associated with the style.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum FontCollectionIndex {
    /// The major font of the style's font scheme.
    #[strum(serialize = "major")]
    Major,
    /// The minor font of the style's font scheme.
    #[strum(serialize = "minor")]
    Minor,
    /// No font reference.
    #[strum(serialize = "none")]
    None,
}

/// This simple type specifies an animation build step within a diagram animation.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum DgmBuildStep {
    /// Animate a diagram shape for this animation build step
    #[strum(serialize = "sp")]
    Shape,
    /// Animate the diagram background for this animation build step
    #[strum(serialize = "bg")]
    Background,
}

/// This simple type specifies an animation build step within a chart animation.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum ChartBuildStep {
    /// Animate a chart category for this animation build step
    #[strum(serialize = "category")]
    Category,
    /// Animate a point in a chart category for this animation build step
    #[strum(serialize = "ptInCategory")]
    PtInCategory,
    /// Animate a chart series for this animation build step
    #[strum(serialize = "series")]
    Series,
    /// Animate a point in a chart series for this animation build step
    #[strum(serialize = "ptInSeries")]
    PtInSeries,
    /// Animate all points within the chart for this animation build step
    #[strum(serialize = "allPts")]
    AllPts,
    /// Animate the chart grid and legend for this animation build step
    #[strum(serialize = "gridLegend")]
    GridLegend,
}

/// This simple type represents whether a style property should be applied.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum OnOffStyleType {
    /// Property is on.
    #[strum(serialize = "on")]
    On,
    /// Property is off.
    #[strum(serialize = "off")]
    Off,
    /// Follow parent settings. For a themed property, follow the theme settings. For an unthemed property, follow
    /// the parent setting in the property inheritance chain.
    #[strum(serialize = "def")]
    Default,
}

/// This simple type specifies a system color value. This color is based upon the value that this color currently has
/// within the system on which the document is being viewed.
///
/// Applications shall use the lastClr attribute to determine the absolute value of the last color used if system colors
/// are not supported.
#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum SystemColorVal {
    /// Specifies the scroll bar gray area color.
    #[strum(serialize = "scrollBar")]
    ScrollBar,
    ///Specifies the desktop background color.
    #[strum(serialize = "background")]
    Background,
    /// Specifies the active window title bar color. In particular the left side color in the color gradient of an
    /// active window's title bar if the gradient effect is enabled.
    #[strum(serialize = "activeCaption")]
    ActiveCaption,
    /// Specifies the color of the Inactive window caption.
    /// Specifies the left side color in the color gradient of an inactive window's title bar if the gradient effect is
    /// enabled.
    #[strum(serialize = "inactiveCaption")]
    InactiveCaption,
    /// Specifies the menu background color.
    #[strum(serialize = "menu")]
    Menu,
    /// Specifies window background color.
    #[strum(serialize = "window")]
    Window,
    /// Specifies the window frame color.
    #[strum(serialize = "windowFrame")]
    WindowFrame,
    /// Specifies the color of Text in menus.
    #[strum(serialize = "menuText")]
    MenuText,
    /// Specifies the color of text in windows.
    #[strum(serialize = "windowText")]
    WindowText,
    /// Specifies the color of text in the caption, size box, and scroll bar arrow box.
    #[strum(serialize = "captionText")]
    CaptionText,
    /// Specifies an Active Window Border Color.
    #[strum(serialize = "activeBorder")]
    ActiveBorder,
    /// Specifies the color of the Inactive window border.
    #[strum(serialize = "inactiveBorder")]
    InactiveBorder,
    /// Specifies the Background color of multiple document interface (MDI) applications
    #[strum(serialize = "appWorkspace")]
    AppWorkspace,
    /// Specifies the color of Item(s) selected in a control.
    #[strum(serialize = "highlight")]
    Highlight,
    /// Specifies the text color of item(s) selected in a control.
    #[strum(serialize = "highlightText")]
    HighlightText,
    /// Specifies the face color for three-dimensional display elements and for dialog box backgrounds.
    #[strum(serialize = "btnFace")]
    ButtonFace,
    /// Specifies the shadow color for three-dimensional display elements (for edges facing away from the light source).
    #[strum(serialize = "btnShadow")]
    ButtonShadow,
    /// Specifies a grayed (disabled) text. This color is set to 0 if the current display driver does not support a
    /// solid gray color.
    #[strum(serialize = "grayText")]
    GrayText,
    /// Specifies the color of text on push buttons.
    #[strum(serialize = "btnText")]
    ButtonText,
    /// Specifies the color of text in an inactive caption.
    #[strum(serialize = "inactiveCaptionText")]
    InactiveCaptionText,
    /// Specifies the highlight color for three-dimensional display elements (for edges facing the light source).
    #[strum(serialize = "btnHighlight")]
    ButtonHighlight,
    /// Specifies a Dark shadow color for three-dimensional display elements.
    #[strum(serialize = "3dDkShadow")]
    DarkShadow3d,
    /// Specifies a Light color for three-dimensional display elements (for edges facing the light source).
    #[strum(serialize = "3dLight")]
    Light3d,
    /// Specifies the text color for tooltip controls.
    #[strum(serialize = "infoText")]
    InfoText,
    /// Specifies the background color for tooltip controls.
    #[strum(serialize = "infoBk")]
    InfoBack,
    #[strum(serialize = "hotLight")]
    /// Specifies the color for a hyperlink or hot-tracked item.
    HotLight,
    #[strum(serialize = "gradientActiveCaption")]
    /// Specifies the right side color in the color gradient of an active window's title bar.
    GradientActiveCaption,
    /// Specifies the right side color in the color gradient of an inactive window's title bar.
    #[strum(serialize = "gradientInactiveCaption")]
    GradientInactiveCaption,
    /// Specifies the color used to highlight menu items when the menu appears as a flat menu.
    #[strum(serialize = "menuHighlight")]
    MenuHighlight,
    /// Specifies the background color for the menu bar when menus appear as flat menus.
    #[strum(serialize = "menubar")]
    MenuBar,
}

/// This simple type represents a preset color value.
#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PresetColorVal {
    /// Specifies a color with RGB value (240,248,255)
    #[strum(serialize = "aliceBlue")]
    AliceBlue,
    /// Specifies a color with RGB value (250,235,215)
    #[strum(serialize = "antiqueWhite")]
    AntiqueWhite,
    /// Specifies a color with RGB value (0,255,255)
    #[strum(serialize = "aqua")]
    Aqua,
    /// Specifies a color with RGB value (127,255,212)
    #[strum(serialize = "aquamarine")]
    Aquamarine,
    /// Specifies a color with RGB value (240,255,255)
    #[strum(serialize = "azure")]
    Azure,
    ///Specifies a color with RGB value (245,245,220)
    #[strum(serialize = "beige")]
    Beige,
    /// Specifies a color with RGB value (255,228,196)
    #[strum(serialize = "bisque")]
    Bisque,
    /// Specifies a color with RGB value (0,0,0)
    #[strum(serialize = "black")]
    Black,
    /// Specifies a color with RGB value (255,235,205)
    #[strum(serialize = "blanchedAlmond")]
    BlanchedAlmond,
    /// Specifies a color with RGB value (0,0,255)
    #[strum(serialize = "blue")]
    Blue,
    /// Specifies a color with RGB value (138,43,226)
    #[strum(serialize = "blueViolet")]
    BlueViolet,
    /// Specifies a color with RGB value (165,42,42)
    #[strum(serialize = "brown")]
    Brown,
    /// Specifies a color with RGB value (222,184,135)
    #[strum(serialize = "burlyWood")]
    BurlyWood,
    /// Specifies a color with RGB value (95,158,160)
    #[strum(serialize = "cadetBlue")]
    CadetBlue,
    /// Specifies a color with RGB value (127,255,0)
    #[strum(serialize = "chartreuse")]
    Chartreuse,
    /// Specifies a color with RGB value (210,105,30)
    #[strum(serialize = "chocolate")]
    Chocolate,
    /// Specifies a color with RGB value (255,127,80)
    #[strum(serialize = "coral")]
    Coral,
    /// Specifies a color with RGB value (100,149,237)
    #[strum(serialize = "cornflowerBlue")]
    CornflowerBlue,
    /// Specifies a color with RGB value (255,248,220)
    #[strum(serialize = "cornsilk")]
    Cornsilk,
    /// Specifies a color with RGB value (220,20,60)
    #[strum(serialize = "crimson")]
    Crimson,
    /// Specifies a color with RGB value (0,255,255)
    #[strum(serialize = "cyan")]
    Cyan,
    /// Specifies a color with RGB value (0,0,139)
    #[strum(serialize = "darkBlue")]
    DarkBlue,
    /// Specifies a color with RGB value (0,139,139)
    #[strum(serialize = "darkCyan")]
    DarkCyan,
    /// Specifies a color with RGB value (184,134,11)
    #[strum(serialize = "darkGoldenrod")]
    DarkGoldenrod,
    /// Specifies a color with RGB value (169,169,169)
    #[strum(serialize = "darkGray")]
    DarkGray,
    /// Specifies a color with RGB value (169,169,169)
    #[strum(serialize = "darkGrey")]
    DarkGrey,
    /// Specifies a color with RGB value (0,100,0)
    #[strum(serialize = "darkGreen")]
    DarkGreen,
    /// Specifies a color with RGB value (189,183,107)
    #[strum(serialize = "darkKhaki")]
    DarkKhaki,
    /// Specifies a color with RGB value (139,0,139)
    #[strum(serialize = "darkMagenta")]
    DarkMagenta,
    /// Specifies a color with RGB value (85,107,47)
    #[strum(serialize = "darkOliveGreen")]
    DarkOliveGreen,
    /// Specifies a color with RGB value (255,140,0)
    #[strum(serialize = "darkOrange")]
    DarkOrange,
    /// Specifies a color with RGB value (153,50,204)
    #[strum(serialize = "darkOrchid")]
    DarkOrchid,
    /// Specifies a color with RGB value (139,0,0)
    #[strum(serialize = "darkRed")]
    DarkRed,
    /// Specifies a color with RGB value (233,150,122)
    #[strum(serialize = "darkSalmon")]
    DarkSalmon,
    /// Specifies a color with RGB value (143,188,143)
    #[strum(serialize = "darkSeaGreen")]
    DarkSeaGreen,
    /// Specifies a color with RGB value (72,61,139)
    #[strum(serialize = "darkSlateBlue")]
    DarkSlateBlue,
    /// Specifies a color with RGB value (47,79,79)
    #[strum(serialize = "darkSlateGray")]
    DarkSlateGray,
    /// Specifies a color with RGB value (47,79,79)
    #[strum(serialize = "darkSlateGrey")]
    DarkSlateGrey,
    /// Specifies a color with RGB value (0,206,209)
    #[strum(serialize = "darkTurquoise")]
    DarkTurqoise,
    /// Specifies a color with RGB value (148,0,211)
    #[strum(serialize = "darkViolet")]
    DarkViolet,
    /// Specifies a color with RGB value (0,0,139)
    #[strum(serialize = "dkBlue")]
    DkBlue,
    /// Specifies a color with RGB value (0,139,139)
    #[strum(serialize = "dkCyan")]
    DkCyan,
    /// Specifies a color with RGB value (184,134,11)
    #[strum(serialize = "dkGoldenrod")]
    DkGoldenrod,
    /// Specifies a color with RGB value (169,169,169)
    #[strum(serialize = "dkGray")]
    DkGray,
    /// Specifies a color with RGB value (169,169,169)
    #[strum(serialize = "dkGrey")]
    DkGrey,
    /// Specifies a color with RGB value (0,100,0)
    #[strum(serialize = "dkGreen")]
    DkGreen,
    /// Specifies a color with RGB value (189,183,107)
    #[strum(serialize = "dkKhaki")]
    DkKhaki,
    /// Specifies a color with RGB value (139,0,139)
    #[strum(serialize = "dkMagenta")]
    DkMagenta,
    /// Specifies a color with RGB value (85,107,47)
    #[strum(serialize = "dkOliveGreen")]
    DkOliveGreen,
    /// Specifies a color with RGB value (255,140,0)
    #[strum(serialize = "dkOrange")]
    DkOrange,
    /// Specifies a color with RGB value (153,50,204)
    #[strum(serialize = "dkOrchid")]
    DkOrchid,
    /// Specifies a color with RGB value (139,0,0)
    #[strum(serialize = "dkRed")]
    DkRed,
    /// Specifies a color with RGB value (233,150,122)
    #[strum(serialize = "dkSalmon")]
    DkSalmon,
    /// Specifies a color with RGB value (143,188,139)
    #[strum(serialize = "dkSeaGreen")]
    DkSeaGreen,
    /// Specifies a color with RGB value (72,61,139)
    #[strum(serialize = "dkSlateBlue")]
    DkSlateBlue,
    /// Specifies a color with RGB value (47,79,79)
    #[strum(serialize = "dkSlateGray")]
    DkSlateGray,
    /// Specifies a color with RGB value (47,79,79)
    #[strum(serialize = "dkSlateGrey")]
    DkSlateGrey,
    /// Specifies a color with RGB value (0,206,209)
    #[strum(serialize = "dkTurquoise")]
    DkTurquoise,
    /// Specifies a color with RGB value (148,0,211)
    #[strum(serialize = "dkViolet")]
    DkViolet,
    /// Specifies a color with RGB value (255,20,147)
    #[strum(serialize = "deepPink")]
    DeepPink,
    /// Specifies a color with RGB value (0,191,255)
    #[strum(serialize = "deepSkyBlue")]
    DeepSkyBlue,
    /// Specifies a color with RGB value (105,105,105)
    #[strum(serialize = "dimGray")]
    DimGray,
    /// Specifies a color with RGB value (105,105,105)
    #[strum(serialize = "dimGrey")]
    DimGrey,
    /// Specifies a color with RGB value (30,144,255)
    #[strum(serialize = "dodgerBlue")]
    DodgerBluet,
    /// Specifies a color with RGB value (178,34,34)
    #[strum(serialize = "firebrick")]
    Firebrick,
    /// Specifies a color with RGB value (255,250,240)
    #[strum(serialize = "floralWhite")]
    FloralWhite,
    /// Specifies a color with RGB value (34,139,34)
    #[strum(serialize = "forestGreen")]
    ForestGreen,
    /// Specifies a color with RGB value (255,0,255)
    #[strum(serialize = "fuchsia")]
    Fuchsia,
    /// Specifies a color with RGB value (220,220,220)
    #[strum(serialize = "gainsboro")]
    Gainsboro,
    /// Specifies a color with RGB value (248,248,255)
    #[strum(serialize = "ghostWhite")]
    GhostWhite,
    /// Specifies a color with RGB value (255,215,0)
    #[strum(serialize = "gold")]
    Gold,
    /// Specifies a color with RGB value (218,165,32)
    #[strum(serialize = "goldenrod")]
    Goldenrod,
    /// Specifies a color with RGB value (128,128,128)
    #[strum(serialize = "gray")]
    Gray,
    /// Specifies a color with RGB value (128,128,128)
    #[strum(serialize = "grey")]
    Grey,
    /// Specifies a color with RGB value (0,128,0)
    #[strum(serialize = "green")]
    Green,
    /// Specifies a color with RGB value (173,255,47)
    #[strum(serialize = "greenYellow")]
    GreenYellow,
    /// Specifies a color with RGB value (240,255,240)
    #[strum(serialize = "honeydew")]
    Honeydew,
    /// Specifies a color with RGB value (255,105,180)
    #[strum(serialize = "hotPink")]
    HotPink,
    /// Specifies a color with RGB value (205,92,92)
    #[strum(serialize = "indianRed")]
    IndianRed,
    /// Specifies a color with RGB value (75,0,130)
    #[strum(serialize = "indigo")]
    Indigo,
    /// Specifies a color with RGB value (255,255,240)
    #[strum(serialize = "ivory")]
    Ivory,
    /// Specifies a color with RGB value (240,230,140)
    #[strum(serialize = "khaki")]
    Khaki,
    /// Specifies a color with RGB value (230,230,250)
    #[strum(serialize = "lavender")]
    Lavender,
    /// Specifies a color with RGB value (255,240,245)
    #[strum(serialize = "lavenderBlush")]
    LavenderBlush,
    /// Specifies a color with RGB value (124,252,0)
    #[strum(serialize = "lawnGreen")]
    LawnGreen,
    /// Specifies a color with RGB value (255,250,205)
    #[strum(serialize = "lemonChiffon")]
    LemonChiffon,
    /// Specifies a color with RGB value (173,216,230)
    #[strum(serialize = "lightBlue")]
    LightBlue,
    /// Specifies a color with RGB value (240,128,128)
    #[strum(serialize = "lightCoral")]
    LightCoral,
    /// Specifies a color with RGB value (224,255,255)
    #[strum(serialize = "lightCyan")]
    LightCyan,
    /// Specifies a color with RGB value (250,250,210)
    #[strum(serialize = "lightGoldenrodYellow")]
    LightGoldenrodYellow,
    /// Specifies a color with RGB value (211,211,211)
    #[strum(serialize = "lightGray")]
    LightGray,
    /// Specifies a color with RGB value (211,211,211)
    #[strum(serialize = "lightGrey")]
    LightGrey,
    /// Specifies a color with RGB value (144,238,144)
    #[strum(serialize = "lightGreen")]
    LightGreen,
    /// Specifies a color with RGB value (255,182,193)
    #[strum(serialize = "lightPink")]
    LightPink,
    /// Specifies a color with RGB value (255,160,122)
    #[strum(serialize = "lightSalmon")]
    LightSalmon,
    /// Specifies a color with RGB value (32,178,170)
    #[strum(serialize = "lightSeaGreen")]
    LightSeaGreen,
    /// Specifies a color with RGB value (135,206,250)
    #[strum(serialize = "lightSkyBlue")]
    LightSkyBlue,
    /// Specifies a color with RGB value (119,136,153)
    #[strum(serialize = "lightSlateGray")]
    LightSlateGray,
    /// Specifies a color with RGB value (119,136,153)
    #[strum(serialize = "lightSlateGrey")]
    LightSlateGrey,
    /// Specifies a color with RGB value (176,196,222)
    #[strum(serialize = "lightSteelBlue")]
    LightSteelBlue,
    /// Specifies a color with RGB value (255,255,224)
    #[strum(serialize = "lightYellow")]
    LightYellow,
    /// Specifies a color with RGB value (173,216,230)
    #[strum(serialize = "ltBlue")]
    LtBlue,
    /// Specifies a color with RGB value (240,128,128)
    #[strum(serialize = "ltCoral")]
    LtCoral,
    /// Specifies a color with RGB value (224,255,255)
    #[strum(serialize = "ltCyan")]
    LtCyan,
    /// Specifies a color with RGB value (250,250,120)
    #[strum(serialize = "ltGoldenrodYellow")]
    LtGoldenrodYellow,
    /// Specifies a color with RGB value (211,211,211)
    #[strum(serialize = "ltGray")]
    LtGray,
    /// Specifies a color with RGB value (211,211,211)
    #[strum(serialize = "ltGrey")]
    LtGrey,
    /// Specifies a color with RGB value (144,238,144)
    #[strum(serialize = "ltGreen")]
    LtGreen,
    /// Specifies a color with RGB value (255,182,193)
    #[strum(serialize = "ltPink")]
    LtPink,
    /// Specifies a color with RGB value (255,160,122)
    #[strum(serialize = "ltSalmon")]
    LtSalmon,
    /// Specifies a color with RGB value (32,178,170)
    #[strum(serialize = "ltSeaGreen")]
    LtSeaGreen,
    /// Specifies a color with RGB value (135,206,250)
    #[strum(serialize = "ltSkyBlue")]
    LtSkyBlue,
    /// Specifies a color with RGB value (119,136,153)
    #[strum(serialize = "ltSlateGray")]
    LtSlateGray,
    /// Specifies a color with RGB value (119,136,153)
    #[strum(serialize = "ltSlateGrey")]
    LtSlateGrey,
    /// Specifies a color with RGB value (176,196,222)
    #[strum(serialize = "ltSteelBlue")]
    LtSteelBlue,
    /// Specifies a color with RGB value (255,255,224)
    #[strum(serialize = "ltYellow")]
    LtYellow,
    /// Specifies a color with RGB value (0,255,0)
    #[strum(serialize = "lime")]
    Lime,
    /// Specifies a color with RGB value (50,205,50)
    #[strum(serialize = "limeGreen")]
    LimeGreen,
    /// Specifies a color with RGB value (250,240,230)
    #[strum(serialize = "linen")]
    Linen,
    /// Specifies a color with RGB value (255,0,255)
    #[strum(serialize = "magenta")]
    Magenta,
    /// Specifies a color with RGB value (128,0,0)
    #[strum(serialize = "maroon")]
    Maroon,
    /// Specifies a color with RGB value (102,205,170)
    #[strum(serialize = "medAquamarine")]
    MedAquamarine,
    /// Specifies a color with RGB value (0,0,205)
    #[strum(serialize = "medBlue")]
    MedBlue,
    /// Specifies a color with RGB value (186,85,211)
    #[strum(serialize = "medOrchid")]
    MedOrchid,
    /// Specifies a color with RGB value (147,112,219)
    #[strum(serialize = "medPurple")]
    MedPurple,
    /// Specifies a color with RGB value (60,179,113)
    #[strum(serialize = "medSeaGreen")]
    MedSeaGreen,
    /// Specifies a color with RGB value (123,104,238)
    #[strum(serialize = "medSlateBlue")]
    MedSlateBlue,
    /// Specifies a color with RGB value (0,250,154)
    #[strum(serialize = "medSpringGreen")]
    MedSpringGreen,
    /// Specifies a color with RGB value (72,209,204)
    #[strum(serialize = "medTurquoise")]
    MedTurquoise,
    /// Specifies a color with RGB value (199,21,133)
    #[strum(serialize = "medVioletRed")]
    MedVioletRed,
    /// Specifies a color with RGB value (102,205,170)
    #[strum(serialize = "mediumAquamarine")]
    MediumAquamarine,
    /// Specifies a color with RGB value (0,0,205)
    #[strum(serialize = "mediumBlue")]
    MediumBlue,
    /// Specifies a color with RGB value (186,85,211)
    #[strum(serialize = "mediumOrchid")]
    MediumOrchid,
    /// Specifies a color with RGB value (147,112,219)
    #[strum(serialize = "mediumPurple")]
    MediumPurple,
    /// Specifies a color with RGB value (60,179,113)
    #[strum(serialize = "mediumSeaGreen")]
    MediumSeaGreen,
    /// Specifies a color with RGB value (123,104,238)
    #[strum(serialize = "mediumSlateBlue")]
    MediumSlateBlue,
    /// Specifies a color with RGB value (0,250,154)
    #[strum(serialize = "mediumSpringGreen")]
    MediumSpringGreen,
    /// Specifies a color with RGB value (72,209,204)
    #[strum(serialize = "mediumTurquoise")]
    MediumTurquoise,
    /// Specifies a color with RGB value (199,21,133)
    #[strum(serialize = "mediumVioletRed")]
    MediumVioletRed,
    /// Specifies a color with RGB value (25,25,112)
    #[strum(serialize = "midnightBlue")]
    MidnightBlue,
    /// Specifies a color with RGB value (245,255,250)
    #[strum(serialize = "mintCream")]
    MintCream,
    /// Specifies a color with RGB value (255,228,225)
    #[strum(serialize = "mistyRose")]
    MistyRose,
    /// Specifies a color with RGB value (255,228,181)
    #[strum(serialize = "moccasin")]
    Moccasin,
    /// Specifies a color with RGB value (255,222,173)
    #[strum(serialize = "navajoWhite")]
    NavajoWhite,
    /// Specifies a color with RGB value (0,0,128)
    #[strum(serialize = "navy")]
    Navy,
    /// Specifies a color with RGB value (253,245,230)
    #[strum(serialize = "oldLace")]
    OldLace,
    /// Specifies a color with RGB value (128,128,0)
    #[strum(serialize = "olive")]
    Olive,
    /// Specifies a color with RGB value (107,142,35)
    #[strum(serialize = "oliveDrab")]
    OliveDrab,
    /// Specifies a color with RGB value (255,165,0)
    #[strum(serialize = "orange")]
    Orange,
    /// Specifies a color with RGB value (255,69,0)
    #[strum(serialize = "orangeRed")]
    OrangeRed,
    /// Specifies a color with RGB value (218,112,214)
    #[strum(serialize = "orchid")]
    Orchid,
    /// Specifies a color with RGB value (238,232,170)
    #[strum(serialize = "paleGoldenrod")]
    PaleGoldenrod,
    /// Specifies a color with RGB value (152,251,152)
    #[strum(serialize = "paleGreen")]
    PaleGreen,
    /// Specifies a color with RGB value (175,238,238)
    #[strum(serialize = "paleTurquoise")]
    PaleTurquoise,
    /// Specifies a color with RGB value (219,112,147)
    #[strum(serialize = "paleVioletRed")]
    PaleVioletRed,
    /// Specifies a color with RGB value (255,239,213)
    #[strum(serialize = "papayaWhip")]
    PapayaWhip,
    /// Specifies a color with RGB value (255,218,185)
    #[strum(serialize = "peachPuff")]
    PeachPuff,
    /// Specifies a color with RGB value (205,133,63)
    #[strum(serialize = "peru")]
    Peru,
    /// Specifies a color with RGB value (255,192,203)
    #[strum(serialize = "pink")]
    Pink,
    /// Specifies a color with RGB value (221,160,221)
    #[strum(serialize = "plum")]
    Plum,
    /// Specifies a color with RGB value (176,224,230)
    #[strum(serialize = "powderBlue")]
    PowderBlue,
    /// Specifies a color with RGB value (128,0,128)
    #[strum(serialize = "purple")]
    Purple,
    /// Specifies a color with RGB value (255,0,0)
    #[strum(serialize = "red")]
    Red,
    /// Specifies a color with RGB value (188,143,143)
    #[strum(serialize = "rosyBrown")]
    RosyBrown,
    /// Specifies a color with RGB value (65,105,225)
    #[strum(serialize = "royalBlue")]
    RoyalBlue,
    /// Specifies a color with RGB value (139,69,19)
    #[strum(serialize = "saddleBrown")]
    SaddleBrown,
    /// Specifies a color with RGB value (250,128,114)
    #[strum(serialize = "salmon")]
    Salmon,
    /// Specifies a color with RGB value (244,164,96)
    #[strum(serialize = "sandyBrown")]
    SandyBrown,
    /// Specifies a color with RGB value (46,139,87)
    #[strum(serialize = "seaGreen")]
    SeaGreen,
    /// Specifies a color with RGB value (255,245,238)
    #[strum(serialize = "seaShell")]
    SeaShell,
    /// Specifies a color with RGB value (160,82,45)
    #[strum(serialize = "sienna")]
    Sienna,
    /// Specifies a color with RGB value (192,192,192)
    #[strum(serialize = "silver")]
    Silver,
    /// Specifies a color with RGB value (135,206,235)
    #[strum(serialize = "skyBlue")]
    SkyBlue,
    /// Specifies a color with RGB value (106,90,205)
    #[strum(serialize = "slateBlue")]
    SlateBlue,
    /// Specifies a color with RGB value (112,128,144)
    #[strum(serialize = "slateGray")]
    SlateGray,
    /// Specifies a color with RGB value (112,128,144)
    #[strum(serialize = "slateGrey")]
    SlateGrey,
    /// Specifies a color with RGB value (255,250,250)
    #[strum(serialize = "snow")]
    Snow,
    /// Specifies a color with RGB value (0,255,127)
    #[strum(serialize = "springGreen")]
    SpringGreen,
    /// Specifies a color with RGB value (70,130,180)
    #[strum(serialize = "steelBlue")]
    SteelBlue,
    /// Specifies a color with RGB value (210,180,140)
    #[strum(serialize = "tan")]
    Tan,
    /// Specifies a color with RGB value (0,128,128)
    #[strum(serialize = "teal")]
    Teal,
    /// Specifies a color with RGB value (216,191,216)
    #[strum(serialize = "thistle")]
    Thistle,
    /// Specifies a color with RGB value (255,99,71)
    #[strum(serialize = "tomato")]
    Tomato,
    /// Specifies a color with RGB value (64,224,208)
    #[strum(serialize = "turquoise")]
    Turquoise,
    /// Specifies a color with RGB value (238,130,238)
    #[strum(serialize = "violet")]
    Violet,
    /// Specifies a color with RGB value (245,222,179)
    #[strum(serialize = "wheat")]
    Wheat,
    /// Specifies a color with RGB value (255,255,255)
    #[strum(serialize = "white")]
    White,
    /// Specifies a color with RGB value (245,245,245)
    #[strum(serialize = "whiteSmoke")]
    WhiteSmoke,
    /// Specifies a color with RGB value (255,255,0)
    #[strum(serialize = "yellow")]
    Yellow,
    /// Specifies a color with RGB value (154,205,50)
    #[strum(serialize = "yellowGreen")]
    YellowGreen,
}

/// This simple type represents a scheme color value.
#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum SchemeColorVal {
    #[strum(serialize = "bg1")]
    Background1,
    #[strum(serialize = "tx1")]
    Text1,
    #[strum(serialize = "bg2")]
    Background2,
    #[strum(serialize = "tx2")]
    Text2,
    #[strum(serialize = "accent1")]
    Accent1,
    #[strum(serialize = "accent2")]
    Accent2,
    #[strum(serialize = "accent3")]
    Accent3,
    #[strum(serialize = "accent4")]
    Accent4,
    #[strum(serialize = "accent5")]
    Accent5,
    #[strum(serialize = "accent6")]
    Accent6,
    #[strum(serialize = "hlink")]
    Hyperlink,
    #[strum(serialize = "folHlink")]
    FollowedHyperlink,
    /// A color used in theme definitions which means to use the color of the style.
    #[strum(serialize = "phClr")]
    PlaceholderColor,
    #[strum(serialize = "dk1")]
    Dark1,
    #[strum(serialize = "lt1")]
    Light1,
    #[strum(serialize = "dk2")]
    Dark2,
    #[strum(serialize = "lt2")]
    Light2,
}

/// A reference to a color in the color scheme.
#[repr(C)]
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum ColorSchemeIndex {
    #[strum(serialize = "dk1")]
    Dark1,
    #[strum(serialize = "lt1")]
    Light1,
    #[strum(serialize = "dk2")]
    Dark2,
    #[strum(serialize = "lt2")]
    Light2,
    #[strum(serialize = "accent1")]
    Accent1,
    #[strum(serialize = "accent2")]
    Accent2,
    #[strum(serialize = "accent3")]
    Accent3,
    #[strum(serialize = "accent4")]
    Accent4,
    #[strum(serialize = "accent5")]
    Accent5,
    #[strum(serialize = "accent6")]
    Accent6,
    #[strum(serialize = "hlink")]
    Hyperlink,
    #[strum(serialize = "folHlink")]
    FollowedHyperlink,
}

/// This simple type specifies the text alignment types
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextAlignType {
    /// Align text to the left margin.
    #[strum(serialize = "l")]
    Left,
    /// Align text in the center.
    #[strum(serialize = "ctr")]
    Center,
    /// Align text to the right margin.
    #[strum(serialize = "r")]
    Right,
    /// Align text so that it is justified across the whole line. It is smart in the sense that it does not justify
    /// sentences which are short.
    #[strum(serialize = "just")]
    Justified,
    /// Aligns the text with an adjusted kashida length for Arabic text.
    #[strum(serialize = "justLow")]
    JustifiedLow,
    /// Distributes the text words across an entire text line.
    #[strum(serialize = "dist")]
    Distributed,
    /// Distributes Thai text specially, because each character is treated as a word.
    #[strum(serialize = "thaiDist")]
    ThaiDistributed,
}

/// This simple type specifies the different kinds of font alignment.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextFontAlignType {
    /// When the text flow is horizontal or simple vertical same as fontBaseline but for other vertical modes
    /// same as fontCenter.
    #[strum(serialize = "auto")]
    Auto,
    /// The letters are anchored to the top baseline of a single line.
    #[strum(serialize = "t")]
    Top,
    /// The letters are anchored between the two baselines of a single line.
    #[strum(serialize = "ctr")]
    Center,
    /// The letters are anchored to the bottom baseline of a single line.
    #[strum(serialize = "base")]
    Baseline,
    /// The letters are anchored to the very bottom of a single line. This is different than the bottom baseline because
    /// of letters such as "g," "q," "y," etc.
    #[strum(serialize = "b")]
    Bottom,
}

/// This simple type specifies a list of automatic numbering schemes.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextAutonumberScheme {
    /// (a), (b), (c), …
    #[strum(serialize = "alphaLcParenBoth")]
    AlphaLcParenBoth,
    /// (A), (B), (C), …
    #[strum(serialize = "alphaUcParenBoth")]
    AlphaUcParenBoth,
    /// a), b), c), …
    #[strum(serialize = "alphaLcParenR")]
    AlphaLcParenR,
    /// A), B), C), …
    #[strum(serialize = "alphaUcParenR")]
    AlphaUcParenR,
    /// a., b., c., …
    #[strum(serialize = "alphaLcPeriod")]
    AlphaLcPeriod,
    /// A., B., C., …
    #[strum(serialize = "alphaUcPeriod")]
    AlphaUcPeriod,
    /// (1), (2), (3), …
    #[strum(serialize = "arabicParenBoth")]
    ArabicParenBoth,
    /// 1), 2), 3), …
    #[strum(serialize = "arabicParenR")]
    ArabicParenR,
    /// 1., 2., 3., …
    #[strum(serialize = "arabicPeriod")]
    ArabicPeriod,
    /// 1, 2, 3, …
    #[strum(serialize = "arabicPlain")]
    ArabicPlain,
    /// (i), (ii), (iii), …
    #[strum(serialize = "romanLcParenBoth")]
    RomanLcParenBoth,
    /// (I), (II), (III), …
    #[strum(serialize = "romanUcParenBoth")]
    RomanUcParenBoth,
    /// i), ii), iii), …
    #[strum(serialize = "romanLcParenR")]
    RomanLcParenR,
    /// I), II), III), …
    #[strum(serialize = "romanUcParenR")]
    RomanUcParenR,
    /// i., ii., iii., …
    #[strum(serialize = "romanLcPeriod")]
    RomanLcPeriod,
    /// I., II., III., …
    #[strum(serialize = "romanUcPeriod")]
    RomanUcPeriod,
    /// Dbl-byte circle numbers (1-10 circle[0x2460-], 11-arabic numbers)
    #[strum(serialize = "circleNumDbPlain")]
    CircleNumDbPlain,
    /// Wingdings black circle numbers
    #[strum(serialize = "circleNumWdBlackPlain")]
    CircleNumWdBlackPlain,
    /// Wingdings white circle numbers (0-10 circle[0x0080-], 11- arabic numbers)
    #[strum(serialize = "circleNumWdWhitePlain")]
    CircleNumWdWhitePlain,
    /// Dbl-byte Arabic numbers w/ double-byte period
    #[strum(serialize = "arabicDbPeriod")]
    ArabicDbPeriod,
    /// Dbl-byte Arabic numbers
    #[strum(serialize = "arabicDbPlain")]
    ArabicDbPlain,
    /// EA: Simplified Chinese w/ single-byte period
    #[strum(serialize = "ea1ChsPeriod")]
    Ea1ChsPeriod,
    /// EA: Simplified Chinese (TypeA 1-99, TypeC 100-)
    #[strum(serialize = "ea1ChsPlain")]
    Ea1ChsPlain,
    /// EA: Traditional Chinese w/ single-byte period
    #[strum(serialize = "ea1ChtPeriod")]
    Ea1ChtPeriod,
    /// EA: Traditional Chinese (TypeA 1-19, TypeC 20-)
    #[strum(serialize = "ea1ChtPlain")]
    Ea1ChtPlain,
    /// EA: Japanese w/ double-byte period
    #[strum(serialize = "ea1JpnChsDbPeriod")]
    Ea1JpnChsDbPeriod,
    /// EA: Japanese/Korean (TypeC 1-)
    #[strum(serialize = "ea1JpnKorPlain")]
    Ea1JpnKorPlain,
    /// EA: Japanese/Korean w/ single-byte period
    #[strum(serialize = "ea1JpnKorPeriod")]
    Ea1JpnKorPeriod,
    /// Bidi Arabic 1 (AraAlpha) with ANSI minus symbol
    #[strum(serialize = "arabic1Minus")]
    Arabic1Minus,
    /// Bidi Arabic 2 (AraAbjad) with ANSI minus symbol
    #[strum(serialize = "arabic2Minus")]
    Arabic2Minus,
    /// Bidi Hebrew 2 with ANSI minus symbol
    #[strum(serialize = "hebrew2Minus")]
    Hebrew2Minus,
    /// Thai alphabet period
    #[strum(serialize = "thaiAlphaPeriod")]
    ThaiAlphaPeriod,
    /// Thai alphabet parentheses - right
    #[strum(serialize = "thaiAlphaParenR")]
    ThaiAlphaParenR,
    /// Thai alphabet parentheses - both
    #[strum(serialize = "thaiAlphaParenBoth")]
    ThaiAlphaParenBoth,
    /// Thai numerical period
    #[strum(serialize = "thaiNumPeriod")]
    ThaiNumPeriod,
    /// Thai numerical parentheses - right
    #[strum(serialize = "thaiNumParenR")]
    ThaiNumParenR,
    /// Thai numerical period
    #[strum(serialize = "thaiNumParenBoth")]
    ThaiNumParenBoth,
    /// Hindi alphabet period - consonants
    #[strum(serialize = "hindiAlphaPeriod")]
    HindiAlphaPeriod,
    /// Hindi numerical period
    #[strum(serialize = "hindiNumPeriod")]
    HindiNumPeriod,
    /// Hindi numerical parentheses - right
    #[strum(serialize = "hindiNumParenR")]
    HindiNumParenR,
    /// Hindi alphabet period - consonants
    #[strum(serialize = "hindiAlpha1Period")]
    HindiAlpha1Period,
}

/// This simple type describes the shape of path to follow for a path gradient shade.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PathShadeType {
    /// Gradient follows the shape
    #[strum(serialize = "shape")]
    Shape,
    /// Gradient follows a circular path
    #[strum(serialize = "circle")]
    Circle,
    /// Gradient follows a rectangular pat
    #[strum(serialize = "rect")]
    Rect,
}

/// This simple type indicates a preset type of pattern fill. The description of each value contains an illustration of
/// the fill type.
///
/// # Note
///
/// These presets correspond to members of the HatchStyle enumeration in the Microsoft .NET Framework.
/// A reference for this type can be found at http://msdn2.microsoft.com/enus/library/system.drawing.drawing2d.hatchstyle.aspx
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum PresetPatternVal {
    #[strum(serialize = "pct5")]
    Percent5,
    #[strum(serialize = "pct10")]
    Percent10,
    #[strum(serialize = "pct20")]
    Percent20,
    #[strum(serialize = "pct25")]
    Percent25,
    #[strum(serialize = "pct30")]
    Percent30,
    #[strum(serialize = "pct40")]
    Percent40,
    #[strum(serialize = "pct50")]
    Percent50,
    #[strum(serialize = "pct60")]
    Percent60,
    #[strum(serialize = "pct70")]
    Percent70,
    #[strum(serialize = "pct75")]
    Percent75,
    #[strum(serialize = "pct80")]
    Percent80,
    #[strum(serialize = "pct90")]
    Percent90,
    #[strum(serialize = "horz")]
    Horizontal,
    #[strum(serialize = "vert")]
    Vertical,
    #[strum(serialize = "ltHorz")]
    LightHorizontal,
    #[strum(serialize = "ltVert")]
    LightVertical,
    #[strum(serialize = "dkHorz")]
    DarkHorizontal,
    #[strum(serialize = "dkVert")]
    DarkVertical,
    #[strum(serialize = "narHorz")]
    NarrowHorizontal,
    #[strum(serialize = "narVert")]
    NarrowVertical,
    #[strum(serialize = "dashHorz")]
    DashedHorizontal,
    #[strum(serialize = "dashVert")]
    DashedVertical,
    #[strum(serialize = "cross")]
    Cross,
    #[strum(serialize = "dnDiag")]
    DownwardDiagonal,
    #[strum(serialize = "upDiag")]
    UpwardDiagonal,
    #[strum(serialize = "ltDnDiag")]
    LightDownwardDiagonal,
    #[strum(serialize = "ltUpDiag")]
    LightUpwardDiagonal,
    #[strum(serialize = "dkDnDiag")]
    DarkDownwardDiagonal,
    #[strum(serialize = "dkUpDiag")]
    DarkUpwardDiagonal,
    #[strum(serialize = "wdDnDiag")]
    WideDownwardDiagonal,
    #[strum(serialize = "wdUpDiag")]
    WideUpwardDiagonal,
    #[strum(serialize = "dashDnDiag")]
    DashedDownwardDiagonal,
    #[strum(serialize = "dashUpDiag")]
    DashedUpwardDiagonal,
    #[strum(serialize = "diagCross")]
    DiagonalCross,
    #[strum(serialize = "smCheck")]
    SmallCheckerBoard,
    #[strum(serialize = "lgCheck")]
    LargeCheckerBoard,
    #[strum(serialize = "smGrid")]
    SmallGrid,
    #[strum(serialize = "lgGrid")]
    LargeGrid,
    #[strum(serialize = "dotGrid")]
    DottedGrid,
    #[strum(serialize = "smConfetti")]
    SmallConfetti,
    #[strum(serialize = "lgConfetti")]
    LargeConfetti,
    #[strum(serialize = "horzBrick")]
    HorizontalBrick,
    #[strum(serialize = "diagBrick")]
    DiagonalBrick,
    #[strum(serialize = "solidDmnd")]
    SolidDiamond,
    #[strum(serialize = "openDmnd")]
    OpenDiamond,
    #[strum(serialize = "dotDmnd")]
    DottedDiamond,
    #[strum(serialize = "plaid")]
    Plaid,
    #[strum(serialize = "sphere")]
    Sphere,
    #[strum(serialize = "weave")]
    Weave,
    #[strum(serialize = "divot")]
    Divot,
    #[strum(serialize = "shingle")]
    Shingle,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "trellis")]
    Trellis,
    #[strum(serialize = "zigzag")]
    ZigZag,
}

/// This simple type describes how to render effects one on top of another.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum BlendMode {
    #[strum(serialize = "over")]
    Overlay,
    #[strum(serialize = "mult")]
    Multiply,
    #[strum(serialize = "screen")]
    Screen,
    #[strum(serialize = "lighten")]
    Lighten,
    #[strum(serialize = "darken")]
    Darken,
}

/// This simple type specifies the text tab alignment types.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextTabAlignType {
    /// The text at this tab stop is left aligned.
    #[strum(serialize = "l")]
    Left,
    /// The text at this tab stop is center aligned.
    #[strum(serialize = "ctr")]
    Center,
    /// The text at this tab stop is right aligned.
    #[strum(serialize = "r")]
    Right,
    /// At this tab stop, the decimals are lined up. From a user's point of view, the text here behaves as right
    /// aligned until the decimal, and then as left aligned after the decimal.
    #[strum(serialize = "dec")]
    Decimal,
}

/// This simple type specifies the text underline types that is used.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextUnderlineType {
    /// The reason we cannot implicitly have noUnderline be the scenario where underline is not specified is
    /// because not being specified implies deriving from a particular style and the user might want to override
    /// that and make some text not be underlined even though the style says otherwise.
    #[strum(serialize = "none")]
    None,
    /// Underline just the words and not the spaces between them.
    #[strum(serialize = "words")]
    Words,
    /// Underline the text with a single line of normal thickness.
    #[strum(serialize = "sng")]
    Single,
    /// Underline the text with two lines of normal thickness.
    #[strum(serialize = "dbl")]
    Double,
    /// Underline the text with a single, thick line.
    #[strum(serialize = "heavy")]
    Heavy,
    /// Underline the text with a single, dotted line of normal thickness.
    #[strum(serialize = "dotted")]
    Dotted,
    /// Underline the text with a single, thick, dotted line.
    #[strum(serialize = "dottedHeavy")]
    DottedHeavy,
    /// Underline the text with a single, dashed line of normal thickness.
    #[strum(serialize = "dash")]
    Dash,
    /// Underline the text with a single, dashed, thick line.
    #[strum(serialize = "dashHeavy")]
    DashHeavy,
    /// Underline the text with a single line consisting of long dashes of normal thickness.
    #[strum(serialize = "dashLong")]
    DashLong,
    /// Underline the text with a single line consisting of long, thick dashes.
    #[strum(serialize = "dashLongHeavy")]
    DashLongHeavy,
    /// Underline the text with a single line of normal thickness consisting of repeating dots and dashes.
    #[strum(serialize = "dotDash")]
    DotDash,
    /// Underline the text with a single, thick line consisting of repeating dots and dashes.
    #[strum(serialize = "dotDashHeavy")]
    DotDashHeavy,
    /// Underline the text with a single line of normal thickness consisting of repeating two dots and dashes.
    #[strum(serialize = "dotDotDash")]
    DotDotDash,
    /// Underline the text with a single, thick line consisting of repeating two dots and dashes.
    #[strum(serialize = "dotDotDashHeavy")]
    DotDotDashHeavy,
    /// Underline the text with a single wavy line of normal thickness.
    #[strum(serialize = "wavy")]
    Wavy,
    /// Underline the text with a single, thick wavy line.
    #[strum(serialize = "wavyHeavy")]
    WavyHeavy,
    /// Underline just the words and not the spaces between them.
    #[strum(serialize = "wavyDbl")]
    WavyDouble,
}

/// This simple type specifies the strike type.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextStrikeType {
    #[strum(serialize = "noStrike")]
    NoStrike,
    #[strum(serialize = "sngStrike")]
    SingleStrike,
    #[strum(serialize = "dblStrike")]
    DoubleStrike,
}

/// This simple type specifies the cap types of the text.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextCapsType {
    /// The reason we cannot implicitly have noCaps be the scenario where capitalization is not specified is
    /// because not being specified implies deriving from a particular style and the user might want to override
    /// that and make some text not have a capitalization scheme even though the style says otherwise.
    #[strum(serialize = "none")]
    None,
    /// Apply small caps to the text. All letters are converted to lower case.
    #[strum(serialize = "small")]
    Small,
    /// Apply all caps on the text. All lower case letters are converted to upper case even though they are stored
    /// differently in the backing store.
    #[strum(serialize = "all")]
    All,
}

/// This simple type specifies the preset text shape geometry that is to be used for a shape. An enumeration of this
/// simple type is used so that a custom geometry does not have to be specified but instead can be constructed
/// automatically by the generating application. For each enumeration listed there is also the corresponding
/// DrawingML code that would be used to construct this shape were it a custom geometry. Within the construction
/// code for each of these preset text shapes there are predefined guides that the generating application shall
/// maintain for calculation purposes at all times. See [ShapeType](enum.ShapeType.html) to see the necessary guide values.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextShapeType {
    #[strum(serialize = "textNoShape")]
    NoShape,
    #[strum(serialize = "textPlain")]
    Plain,
    #[strum(serialize = "textStop")]
    Stop,
    #[strum(serialize = "textTriangle")]
    Triangle,
    #[strum(serialize = "textTriangleInverted")]
    TriangleInverted,
    #[strum(serialize = "textChevron")]
    Chevron,
    #[strum(serialize = "textChevronInverted")]
    ChevronInverted,
    #[strum(serialize = "textRingInside")]
    RingInside,
    #[strum(serialize = "textRingOutside")]
    RingOutside,
    #[strum(serialize = "textArchUp")]
    ArchUp,
    #[strum(serialize = "textArchDown")]
    ArchDown,
    #[strum(serialize = "textCircle")]
    Circle,
    #[strum(serialize = "textButton")]
    Button,
    #[strum(serialize = "textArchUpPour")]
    ArchUpPour,
    #[strum(serialize = "textArchDownPour")]
    ArchDownPour,
    #[strum(serialize = "textCirclePour")]
    CirclePour,
    #[strum(serialize = "textButtonPour")]
    ButtonPour,
    #[strum(serialize = "textCurveUp")]
    CurveUp,
    #[strum(serialize = "textCurveDown")]
    CurveDown,
    #[strum(serialize = "textCanUp")]
    CanUp,
    #[strum(serialize = "textCanDown")]
    CanDown,
    #[strum(serialize = "textWave1")]
    Wave1,
    #[strum(serialize = "textWave2")]
    Wave2,
    #[strum(serialize = "textWave4")]
    Wave4,
    #[strum(serialize = "textDoubleWave1")]
    DoubleWave1,
    #[strum(serialize = "textInflate")]
    Inflate,
    #[strum(serialize = "textDeflate")]
    Deflate,
    #[strum(serialize = "textInflateBottom")]
    InflateBottom,
    #[strum(serialize = "textDeflateBottom")]
    DeflateBottom,
    #[strum(serialize = "textInflateTop")]
    InflateTop,
    #[strum(serialize = "textDeflateTop")]
    DeflateTop,
    #[strum(serialize = "textDeflateInflate")]
    DeflateInflate,
    #[strum(serialize = "textDeflateInflateDeflate")]
    DeflateInflateDeflate,
    #[strum(serialize = "textFadeLeft")]
    FadeLeft,
    #[strum(serialize = "textFadeUp")]
    FadeUp,
    #[strum(serialize = "textFadeRight")]
    FadeRight,
    #[strum(serialize = "textFadeDown")]
    FadeDown,
    #[strum(serialize = "textSlantUp")]
    SlantUp,
    #[strum(serialize = "textSlantDown")]
    SlantDown,
    #[strum(serialize = "textCascadeUp")]
    CascadeUp,
    #[strum(serialize = "textCascadeDown")]
    CascadeDown,
}

/// This simple type specifies the text vertical overflow.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextVertOverflowType {
    /// Overflow the text and pay no attention to top and bottom barriers.
    #[strum(serialize = "overflow")]
    Overflow,
    /// Pay attention to top and bottom barriers. Use an ellipsis to denote that there is text which is not visible.
    #[strum(serialize = "ellipsis")]
    Ellipsis,
    /// Pay attention to top and bottom barriers. Provide no indication that there is text which is not visible.
    #[strum(serialize = "clip")]
    Clip,
}

/// This simple type specifies the text horizontal overflow types
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextHorizontalOverflowType {
    /// When a big character does not fit into a line, allow a horizontal overflow.
    #[strum(serialize = "overflow")]
    Overflow,
    /// When a big character does not fit into a line, clip it at the proper horizontal overflow.
    #[strum(serialize = "clip")]
    Clip,
}

/// If there is vertical text, determines what kind of vertical text is going to be used.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextVerticalType {
    /// Horizontal text. This should be default.
    #[strum(serialize = "horz")]
    Horizontal,
    /// Determines if all of the text is vertical orientation (each line is 90 degrees rotated clockwise, so it goes
    /// from top to bottom; each next line is to the left from the previous one).
    #[strum(serialize = "vert")]
    Vertical,
    /// Determines if all of the text is vertical orientation (each line is 270 degrees rotated clockwise, so it goes
    /// from bottom to top; each next line is to the right from the previous one).
    #[strum(serialize = "vert270")]
    Vertical270,
    /// Determines if all of the text is vertical ("one letter on top of another").
    #[strum(serialize = "wordArtVert")]
    WordArtVertical,
    /// A special version of vertical text, where some fonts are displayed as if rotated by 90 degrees while some fonts
    /// (mostly East Asian) are displayed vertical.
    #[strum(serialize = "eaVert")]
    EastAsianVertical,
    /// A special version of vertical text, where some fonts are displayed as if rotated by 90 degrees while some fonts
    /// (mostly East Asian) are displayed vertical. The difference between this and the eastAsianVertical is
    /// the text flows top down then LEFT RIGHT, instead of RIGHT LEFT
    #[strum(serialize = "mongolianVert")]
    MongolianVertical,
    /// Specifies that vertical WordArt should be shown from right to left rather than left to right.
    #[strum(serialize = "wordArtVertRtl")]
    WordArtVerticalRtl,
}

#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextWrappingType {
    /// No wrapping occurs on this text body. Words spill out without paying attention to the bounding rectangle
    /// boundaries.
    #[strum(serialize = "none")]
    None,
    /// Determines whether we wrap words within the bounding rectangle.
    #[strum(serialize = "square")]
    Square,
}

/// This simple type specifies a list of available anchoring types for text.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum TextAnchoringType {
    /// Anchor the text at the top of the bounding rectangle.
    #[strum(serialize = "t")]
    Top,
    /// Anchor the text at the middle of the bounding rectangle.
    #[strum(serialize = "ctr")]
    Center,
    /// Anchor the text at the bottom of the bounding rectangle.
    #[strum(serialize = "b")]
    Bottom,
    /// Anchor the text so that it is justified vertically. When text is horizontal, this spaces out the actual lines of
    /// text and is almost always identical in behavior to 'distrib' (special case: if only 1 line, then anchored at
    /// top). When text is vertical, then it justifies the letters vertically. This is different than anchorDistributed,
    /// because in some cases such as very little text in a line, it does not justify.
    #[strum(serialize = "just")]
    Justified,
    /// Anchor the text so that it is distributed vertically.
    /// When text is horizontal, this spaces out the actual lines of text and is almost always identical in behavior to
    /// anchorJustified (special case: if only 1 line, then anchored in middle). When text is vertical, then it
    /// distributes the letters vertically. This is different than anchorJustified, because it always forces
    /// distribution of the words, even if there are only one or two words in a line.
    #[strum(serialize = "dist")]
    Distributed,
}

/// This simple type specifies how an object should be rendered when specified to be in black and white mode.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum BlackWhiteMode {
    /// Object rendered with normal coloring
    #[strum(serialize = "clr")]
    Color,
    /// Object rendered with automatic coloring
    #[strum(serialize = "auto")]
    Auto,
    /// Object rendered with gray coloring
    #[strum(serialize = "gray")]
    Gray,
    /// Object rendered with light gray coloring
    #[strum(serialize = "ltGray")]
    LightGray,
    /// Object rendered with inverse gray coloring
    #[strum(serialize = "invGray")]
    InverseGray,
    /// Object rendered within gray and white coloring
    #[strum(serialize = "grayWhite")]
    GrayWhite,
    /// Object rendered with black and gray coloring
    #[strum(serialize = "blackGray")]
    BlackGray,
    /// Object rendered within black and white coloring
    #[strum(serialize = "blackWhite")]
    BlackWhite,
    /// Object rendered with black-only coloring
    #[strum(serialize = "black")]
    Black,
    /// Object rendered within white coloirng
    #[strum(serialize = "white")]
    White,
    /// Object rendered with hidden coloring
    #[strum(serialize = "hidden")]
    Hidden,
}

/// This simple type specifies the ways that an animation can be built, or animated.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum AnimationBuildType {
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
}

/// This simple type specifies the build options available only for animating a diagram. These options specify the
/// manner in which the objects within the chart should be grouped and animated.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum AnimationDgmOnlyBuildType {
    /// Animate the diagram by elements. For a tree diagram the animation occurs by branch within the diagram tree.
    #[strum(serialize = "one")]
    One,
    /// Animate the diagram by the elements within a level, animating them one level element at a time.
    #[strum(serialize = "lvlOne")]
    LvlOne,
    /// Animate the diagram one level at a time, animating the whole level as one object
    #[strum(serialize = "lvlAtOnce")]
    LvlAtOnce,
}

/// This simple type specifies the ways that a diagram animation can be built. That is, it specifies the way in which
/// the objects within the diagram graphical object should be animated.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum AnimationDgmBuildType {
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
    #[strum(serialize = "one")]
    One,
    #[strum(serialize = "lvlOne")]
    LvlOne,
    #[strum(serialize = "lvlAtOnce")]
    LvlAtOnce,
}

/// This simple type specifies the build options available only for animating a chart. These options specify the
/// manner in which the objects within the chart should be grouped and animated.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum AnimationChartOnlyBuildType {
    /// Animate by each series
    #[strum(serialize = "series")]
    Series,
    /// Animate by each category
    #[strum(serialize = "category")]
    Category,
    /// Animate by each element within the series
    #[strum(serialize = "seriesElement")]
    SeriesElement,
    /// Animate by each element within the category
    #[strum(serialize = "categoryElement")]
    CategoryElement,
}

/// This simple type specifies the ways that a chart animation can be built. That is, it specifies the way in which the
/// objects within the chart should be animated.
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum AnimationChartBuildType {
    #[strum(serialize = "allAtOnce")]
    AllAtOnce,
    #[strum(serialize = "series")]
    Series,
    #[strum(serialize = "category")]
    Category,
    #[strum(serialize = "seriesElement")]
    SeriesElement,
    #[strum(serialize = "categoryElement")]
    CategoryElement,
}

/// This type specifies the amount of compression that has been used for a particular binary large image or picture
/// (blip).
#[derive(Debug, Clone, Copy, EnumString, PartialEq)]
pub enum BlipCompression {
    /// Compression size suitable for inclusion with email
    #[strum(serialize = "email")]
    Email,
    /// Compression size suitable for viewing on screen
    #[strum(serialize = "screen")]
    Screen,
    /// Compression size suitable for printing
    #[strum(serialize = "print")]
    Print,
    /// Compression size suitable for high quality printing
    #[strum(serialize = "hqprint")]
    HqPrint,
    /// No compression was used
    #[strum(serialize = "none")]
    None,
}
