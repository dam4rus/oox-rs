use crate::error::PatternRestrictionError;
use regex::Regex;
use std::{marker::PhantomData, str::FromStr};

pub type OnOff = bool;
pub type Lang = String;
pub type XmlName = String; // 1 <= length <= 255
pub type PositiveUniversalMeasure = UniversalMeasure<Unsigned>;

/// Trait indicating that a data type is restricted by a string pattern. A pattern is basically a regular expression.
pub trait PatternRestricted {
    fn restriction_pattern() -> &'static str;
}

/// Empty struct used to tag a data type implying that the stored value is signed.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Signed;

/// Empty struct used to tag a data type implying that the stored value is unsigned.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Unsigned;

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum CalendarType {
    #[strum(serialize = "gregorian")]
    Gregorian,
    #[strum(serialize = "gregorianUs")]
    GregorianUs,
    #[strum(serialize = "gregorianMeFrench")]
    GregorianMeFrench,
    #[strum(serialize = "gregorianArabic")]
    GregorianArabic,
    #[strum(serialize = "hijri")]
    Hijri,
    #[strum(serialize = "hebrew")]
    Hebrew,
    #[strum(serialize = "taiwan")]
    Taiwan,
    #[strum(serialize = "japan")]
    Japan,
    #[strum(serialize = "thai")]
    Thai,
    #[strum(serialize = "korea")]
    Korea,
    #[strum(serialize = "saka")]
    Saka,
    #[strum(serialize = "gregorianXlitEnglish")]
    GregorianXlitEnglish,
    #[strum(serialize = "gregorianXlitFrench")]
    GregorianXlitFrench,
    #[strum(serialize = "none")]
    None,
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum VerticalAlignRun {
    #[strum(serialize = "baseline")]
    Baseline,
    #[strum(serialize = "superscript")]
    Superscript,
    #[strum(serialize = "subscript")]
    Subscript,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum XAlign {
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum YAlign {
    #[strum(serialize = "inline")]
    Inline,
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "bottom")]
    Bottom,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum UniversalMeasureUnit {
    #[strum(serialize = "mm")]
    Millimeter,
    #[strum(serialize = "cm")]
    Centimeter,
    #[strum(serialize = "in")]
    Inch,
    #[strum(serialize = "pt")]
    Point,
    #[strum(serialize = "pc")]
    Pica,
    #[strum(serialize = "pi")]
    Pitch,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct UniversalMeasure<T = Signed> {
    pub value: f64,
    pub unit: UniversalMeasureUnit,
    pub _phantom: PhantomData<T>,
}

impl<T> UniversalMeasure<T> {
    pub fn new(value: f64, unit: UniversalMeasureUnit) -> Self {
        Self {
            value,
            unit,
            _phantom: PhantomData,
        }
    }
}

impl PatternRestricted for UniversalMeasure<Signed> {
    fn restriction_pattern() -> &'static str {
        r#"^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$"#
    }
}

impl PatternRestricted for UniversalMeasure<Unsigned> {
    fn restriction_pattern() -> &'static str {
        r#"^([0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$"#
    }
}

impl<T> FromStr for UniversalMeasure<T>
where
    UniversalMeasure<T>: PatternRestricted,
{
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let re = Regex::new(Self::restriction_pattern()).expect("valid regexp should be provided");
        let captures = re
            .captures(s)
            .ok_or_else(|| Box::new(PatternRestrictionError::NoMatch))?;
        // Group 1 and 2 can't be empty if the match succeeds
        let value_slice = captures.get(1).unwrap();
        let unit_slice = captures.get(2).unwrap();
        Ok(Self::new(value_slice.as_str().parse()?, unit_slice.as_str().parse()?))
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TwipsMeasure {
    Decimal(u64),
    UniversalMeasure(PositiveUniversalMeasure),
}

impl FromStr for TwipsMeasure {
    // TODO custom error type
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        if let Ok(value) = s.parse::<u64>() {
            Ok(TwipsMeasure::Decimal(value))
        } else {
            Ok(TwipsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Percentage(pub f64);

impl PatternRestricted for Percentage {
    fn restriction_pattern() -> &'static str {
        r#"^(-?[0-9]+(?:\.[0-9]+)?)%$"#
    }
}

impl FromStr for Percentage {
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let re = Regex::new(Self::restriction_pattern()).expect("valid regexp should be provided");
        let captures = re
            .captures(s)
            .ok_or_else(|| Box::new(PatternRestrictionError::NoMatch))?;

        Ok(Self(captures.get(1).unwrap().as_str().parse()?))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, EnumString)]
pub enum ConformanceClass {
    #[strum(serialize = "strict")]
    Strict,
    #[strum(serialize = "transitional")]
    Transitional,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    pub fn test_universal_measure_from_str() {
        assert_eq!(
            "123.4567mm".parse::<UniversalMeasure>().unwrap(),
            UniversalMeasure::new(123.4567, UniversalMeasureUnit::Millimeter),
        );
        assert_eq!(
            "123cm".parse::<UniversalMeasure>().unwrap(),
            UniversalMeasure::new(123.0, UniversalMeasureUnit::Centimeter),
        );
        assert_eq!(
            "123cm".parse::<PositiveUniversalMeasure>().unwrap(),
            PositiveUniversalMeasure::new(123.0, UniversalMeasureUnit::Centimeter),
        );
        assert_eq!(
            "-123in".parse::<UniversalMeasure>().unwrap(),
            UniversalMeasure::new(-123.0, UniversalMeasureUnit::Inch),
        );
    }

    #[test]
    pub fn test_twips_measure_from_str() {
        assert_eq!("123".parse::<TwipsMeasure>().unwrap(), TwipsMeasure::Decimal(123));
        assert_eq!(
            "123.456mm".parse::<TwipsMeasure>().unwrap(),
            TwipsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter)),
        );
    }

    #[test]
    pub fn test_percentage_from_str() {
        assert_eq!("100%".parse::<Percentage>().unwrap(), Percentage(100.0));
        assert_eq!("-100%".parse::<Percentage>().unwrap(), Percentage(-100.0));
        assert_eq!("123.456%".parse::<Percentage>().unwrap(), Percentage(123.456));
        assert_eq!("-123.456%".parse::<Percentage>().unwrap(), Percentage(-123.456));
    }
}
