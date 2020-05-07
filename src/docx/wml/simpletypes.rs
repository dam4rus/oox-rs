use crate::{
    error::{ParseBoolError, PatternRestrictionError},
    shared::sharedtypes::OnOff,
    xml::{parse_xml_bool, XmlNode},
};
use regex::Regex;

pub type UcharHexNumber = u8;
pub type ShortHexNumber = u16;
pub type LongHexNumber = u32;
pub type UnqualifiedPercentage = i32;
pub type DecimalNumber = i64;
pub type UnsignedDecimalNumber = u64;
pub type DateTime = String;
pub type MacroName = String; // maxLength=33
pub type FFName = String; // maxLength=65
pub type FFHelpTextVal = String; // maxLength=256
pub type FFStatusTextVal = String; // maxLength=140
pub type EightPointMeasure = u64;
pub type PointMeasure = u64;
pub type TextScalePercent = f64; // pattern=0*(600|([0-5]?[0-9]?[0-9]))%
pub type TextScaleDecimal = i32; // 0 <= n <= 600
pub type TextScale = TextScalePercent;

pub(crate) fn parse_text_scale_percent(s: &str) -> Result<f64, Box<dyn std::error::Error>> {
    let re = Regex::new("^0*(600|([0-5]?[0-9]?[0-9]))%$").expect("valid regexp should be provided");
    let captures = re.captures(s).ok_or_else(|| PatternRestrictionError::NoMatch)?;
    Ok(f64::from(captures[1].parse::<i32>()?))
}

pub(crate) fn parse_on_off_xml_element(xml_node: &XmlNode) -> Result<OnOff, ParseBoolError> {
    Ok(xml_node
        .attributes
        .get("w:val")
        .map(parse_xml_bool)
        .transpose()?
        .unwrap_or(true))
}
