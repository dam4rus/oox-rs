use super::wml::{
    document::{
        Border, Color, EastAsianLayout, Em, FitText, Fonts, HighlightColor, HpsMeasure, Language, PPrBase, RPrBase,
        Shd, SignedHpsMeasure, SignedTwipsMeasure, TextEffect, Underline,
    },
    simpletypes::TextScale,
    styles::Style,
};
use crate::{
    shared::sharedtypes::{OnOff, VerticalAlignRun},
    update::{update_options, Update},
};

pub type ParagraphProperties = PPrBase;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RunProperties {
    pub style: Option<String>,
    pub fonts: Option<Fonts>,
    pub bold: Option<OnOff>,
    pub complex_script_bold: Option<OnOff>,
    pub italic: Option<OnOff>,
    pub complex_script_italic: Option<OnOff>,
    pub all_capitals: Option<OnOff>,
    pub all_small_capitals: Option<OnOff>,
    pub strikethrough: Option<OnOff>,
    pub double_strikethrough: Option<OnOff>,
    pub outline: Option<OnOff>,
    pub shadow: Option<OnOff>,
    pub emboss: Option<OnOff>,
    pub imprint: Option<OnOff>,
    pub no_proofing: Option<OnOff>,
    pub snap_to_grid: Option<OnOff>,
    pub vanish: Option<OnOff>,
    pub web_hidden: Option<OnOff>,
    pub color: Option<Color>,
    pub spacing: Option<SignedTwipsMeasure>,
    pub width: Option<TextScale>,
    pub kerning: Option<HpsMeasure>,
    pub position: Option<SignedHpsMeasure>,
    pub font_size: Option<HpsMeasure>,
    pub complex_script_font_size: Option<HpsMeasure>,
    pub highlight: Option<HighlightColor>,
    pub underline: Option<Underline>,
    pub effect: Option<TextEffect>,
    pub border: Option<Border>,
    pub shading: Option<Shd>,
    pub fit_text: Option<FitText>,
    pub vertical_alignment: Option<VerticalAlignRun>,
    pub rtl: Option<OnOff>,
    pub complex_script: Option<OnOff>,
    pub emphasis_mark: Option<Em>,
    pub language: Option<Language>,
    pub east_asian_layout: Option<EastAsianLayout>,
    pub special_vanish: Option<OnOff>,
    pub o_math: Option<OnOff>,
}

impl RunProperties {
    pub fn from_vec(properties_vec: &[RPrBase]) -> Self {
        properties_vec
            .iter()
            .fold(Default::default(), |mut instance: Self, property| {
                match property {
                    RPrBase::RunStyle(style) => instance.style = Some(style.clone()),
                    RPrBase::RunFonts(fonts) => instance.fonts = Some(fonts.clone()),
                    RPrBase::Bold(b) => instance.bold = Some(*b),
                    RPrBase::ComplexScriptBold(b) => instance.complex_script_bold = Some(*b),
                    RPrBase::Italic(i) => instance.italic = Some(*i),
                    RPrBase::ComplexScriptItalic(i) => instance.complex_script_italic = Some(*i),
                    RPrBase::Capitals(caps) => instance.all_capitals = Some(*caps),
                    RPrBase::SmallCapitals(small_caps) => instance.all_small_capitals = Some(*small_caps),
                    RPrBase::Strikethrough(strike) => {
                        instance.strikethrough = Some(*strike);
                        instance.double_strikethrough = None;
                    }
                    RPrBase::DoubleStrikethrough(dbl_strike) => {
                        instance.double_strikethrough = Some(*dbl_strike);
                        instance.strikethrough = None;
                    }
                    RPrBase::Outline(outline) => instance.outline = Some(*outline),
                    RPrBase::Shadow(shadow) => instance.shadow = Some(*shadow),
                    RPrBase::Emboss(emboss) => instance.emboss = Some(*emboss),
                    RPrBase::Imprint(imprint) => instance.imprint = Some(*imprint),
                    RPrBase::NoProofing(no_proof) => instance.no_proofing = Some(*no_proof),
                    RPrBase::SnapToGrid(snap_to_grid) => instance.snap_to_grid = Some(*snap_to_grid),
                    RPrBase::Vanish(vanish) => instance.vanish = Some(*vanish),
                    RPrBase::WebHidden(web_hidden) => instance.web_hidden = Some(*web_hidden),
                    RPrBase::Color(color) => instance.color = Some(*color),
                    RPrBase::Spacing(spacing) => instance.spacing = Some(*spacing),
                    RPrBase::Width(width) => instance.width = Some(*width),
                    RPrBase::Kerning(kerning) => instance.kerning = Some(*kerning),
                    RPrBase::Position(pos) => instance.position = Some(*pos),
                    RPrBase::FontSize(size) => instance.font_size = Some(*size),
                    RPrBase::ComplexScriptFontSize(cs_size) => instance.complex_script_font_size = Some(*cs_size),
                    RPrBase::Highlight(color) => instance.highlight = Some(*color),
                    RPrBase::Underline(u) => instance.underline = Some(*u),
                    RPrBase::Effect(effect) => instance.effect = Some(*effect),
                    RPrBase::Border(border) => instance.border = Some(*border),
                    RPrBase::Shading(shd) => instance.shading = Some(*shd),
                    RPrBase::FitText(fit_text) => instance.fit_text = Some(*fit_text),
                    RPrBase::VerticalAlignment(align) => instance.vertical_alignment = Some(*align),
                    RPrBase::Rtl(rtl) => instance.rtl = Some(*rtl),
                    RPrBase::ComplexScript(cs) => instance.complex_script = Some(*cs),
                    RPrBase::EmphasisMark(em) => instance.emphasis_mark = Some(*em),
                    RPrBase::Language(lang) => instance.language = Some(lang.clone()),
                    RPrBase::EastAsianLayout(ea_layout) => instance.east_asian_layout = Some(*ea_layout),
                    RPrBase::SpecialVanish(vanish) => instance.special_vanish = Some(*vanish),
                    RPrBase::OMath(o_math) => instance.o_math = Some(*o_math),
                }

                instance
            })
    }

    pub fn update_with(self, other: Self) -> Self {
        Self {
            style: other.style.or(self.style),
            fonts: update_options(self.fonts, other.fonts),
            bold: other.bold.or(self.bold),
            complex_script_bold: other.complex_script_bold.or(self.complex_script_bold),
            italic: other.italic.or(self.italic),
            complex_script_italic: other.complex_script_italic.or(self.complex_script_italic),
            all_capitals: other.all_capitals.or(self.all_capitals),
            all_small_capitals: other.all_small_capitals.or(self.all_small_capitals),
            strikethrough: other.strikethrough.or(self.strikethrough),
            double_strikethrough: other.double_strikethrough.or(self.double_strikethrough),
            outline: other.outline.or(self.outline),
            shadow: other.shadow.or(self.shadow),
            emboss: other.emboss.or(self.emboss),
            imprint: other.imprint.or(self.imprint),
            no_proofing: other.no_proofing.or(self.no_proofing),
            snap_to_grid: other.snap_to_grid.or(self.snap_to_grid),
            vanish: other.vanish.or(self.vanish),
            web_hidden: other.web_hidden.or(self.web_hidden),
            color: update_options(self.color, other.color),
            spacing: other.spacing.or(self.spacing),
            width: other.width.or(self.width),
            kerning: other.kerning.or(self.kerning),
            position: other.position.or(self.position),
            font_size: other.font_size.or(self.font_size),
            complex_script_font_size: other.complex_script_font_size.or(self.complex_script_font_size),
            highlight: other.highlight.or(self.highlight),
            underline: update_options(self.underline, other.underline),
            effect: other.effect.or(self.effect),
            border: update_options(self.border, other.border),
            shading: update_options(self.shading, other.shading),
            fit_text: other.fit_text.or(self.fit_text),
            vertical_alignment: other.vertical_alignment.or(self.vertical_alignment),
            rtl: other.rtl.or(self.rtl),
            complex_script: other.complex_script.or(self.complex_script),
            emphasis_mark: other.emphasis_mark.or(self.emphasis_mark),
            language: update_options(self.language, other.language),
            east_asian_layout: update_options(self.east_asian_layout, other.east_asian_layout),
            special_vanish: other.special_vanish.or(self.special_vanish),
            o_math: other.o_math.or(self.o_math),
        }
    }

    pub fn update_with_style_on_another_level(self, other: Self) -> Self {
        Self {
            bold: update_or_toggle_on_off(self.bold, other.bold),
            complex_script_bold: update_or_toggle_on_off(self.complex_script_bold, other.complex_script_bold),
            italic: update_or_toggle_on_off(self.italic, other.italic),
            complex_script_italic: update_or_toggle_on_off(self.complex_script_italic, other.complex_script_italic),
            all_capitals: update_or_toggle_on_off(self.all_capitals, other.all_capitals),
            all_small_capitals: update_or_toggle_on_off(self.all_small_capitals, other.all_small_capitals),
            strikethrough: update_or_toggle_on_off(self.strikethrough, other.strikethrough),
            double_strikethrough: update_or_toggle_on_off(self.double_strikethrough, other.double_strikethrough),
            outline: update_or_toggle_on_off(self.outline, other.outline),
            shadow: update_or_toggle_on_off(self.shadow, other.shadow),
            emboss: update_or_toggle_on_off(self.emboss, other.emboss),
            imprint: update_or_toggle_on_off(self.imprint, other.imprint),
            no_proofing: update_or_toggle_on_off(self.no_proofing, other.no_proofing),
            snap_to_grid: update_or_toggle_on_off(self.snap_to_grid, other.snap_to_grid),
            vanish: update_or_toggle_on_off(self.vanish, other.vanish),
            web_hidden: update_or_toggle_on_off(self.web_hidden, other.web_hidden),
            rtl: update_or_toggle_on_off(self.rtl, other.rtl),
            complex_script: update_or_toggle_on_off(self.complex_script, self.complex_script),
            special_vanish: update_or_toggle_on_off(self.special_vanish, other.special_vanish),
            o_math: update_or_toggle_on_off(self.o_math, other.o_math),
            ..self.update_with(other)
        }
    }
}
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ResolvedStyle {
    pub paragraph_properties: Box<ParagraphProperties>,
    pub run_properties: Box<RunProperties>,
}

impl ResolvedStyle {
    pub fn from_paragraph_properties(paragraph_properties: Box<ParagraphProperties>) -> Self {
        Self {
            paragraph_properties,
            ..Default::default()
        }
    }

    pub fn from_run_properties(run_properties: Box<RunProperties>) -> Self {
        Self {
            run_properties,
            ..Default::default()
        }
    }

    pub fn from_wml_style(style: &Style) -> Self {
        let paragraph_properties = Box::new(
            style
                .paragraph_properties
                .as_ref()
                .map(|p_pr| p_pr.base.clone())
                .unwrap_or_default(),
        );

        let run_properties = Box::new(
            style
                .run_properties
                .as_ref()
                .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
                .unwrap_or_default(),
        );

        Self {
            paragraph_properties,
            run_properties,
        }
    }

    pub fn update_with(mut self, other: Self) -> Self {
        *self.paragraph_properties = self.paragraph_properties.update_with(*other.paragraph_properties);
        *self.run_properties = self.run_properties.update_with(*other.run_properties);
        self
    }

    pub fn update_with_style_on_another_level(mut self, other: Self) -> Self {
        *self.paragraph_properties = self.paragraph_properties.update_with(*other.paragraph_properties);
        *self.run_properties = self
            .run_properties
            .update_with_style_on_another_level(*other.run_properties);
        self
    }

    pub fn update_paragraph_with(mut self, other: ParagraphProperties) -> Self {
        *self.paragraph_properties = self.paragraph_properties.update_with(other);
        self
    }

    pub fn update_run_with(mut self, other: RunProperties) -> Self {
        *self.run_properties = self.run_properties.update_with(other);
        self
    }
}

fn update_or_toggle_on_off(lhs: Option<OnOff>, rhs: Option<OnOff>) -> Option<OnOff> {
    match (lhs, rhs) {
        (Some(lhs), Some(rhs)) => Some(lhs ^ rhs),
        (lhs, rhs) => rhs.or(lhs),
    }
}
