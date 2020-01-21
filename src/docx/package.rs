use super::{
    resolvedstyle::{ResolvedStyle, RunProperties},
    wml::{
        document::{
            BlockLevelElts, ContentBlockContent, ContentRunContent, Document, PContent, PPr, RPr, RPrBase, SectPrContents,
            P, R,
        },
        footnotes::{Footnotes, FtnEdn, FtnEdnType},
        numbering::{Lvl, Numbering},
        settings::Settings,
        styles::{Style, StyleType, Styles},
    },
};
use log::error;
use crate::{
    shared::{
        docprops::{AppInfo, Core},
        drawingml::sharedstylesheet::OfficeStyleSheet,
        relationship::{Relationship, THEME_RELATION_TYPE},
    },
    update::Update,
    xml::zip_file_to_xml_node,
};
use std::{
    collections::HashMap,
    error::Error,
    ffi::OsStr,
    fs::File,
    path::{Path, PathBuf},
};
use zip::ZipArchive;

#[derive(Debug, Default)]
pub struct Package {
    pub app_info: Option<AppInfo>,
    pub core: Option<Core>,
    pub main_document: Option<Box<Document>>,
    pub main_document_relationships: Vec<Relationship>,
    pub styles: Option<Box<Styles>>,
    pub footnotes: Option<Footnotes>,
    pub numbering: Option<Numbering>,
    pub settings: Option<Box<Settings>>,
    pub medias: Vec<PathBuf>,
    pub themes: HashMap<String, OfficeStyleSheet>,
}

impl Package {
    pub fn from_file(file_path: &Path) -> Result<Self, Box<dyn Error>> {
        let file = File::open(file_path)?;
        let mut zipper = ZipArchive::new(&file)?;

        let mut instance: Self = Default::default();
        for idx in 0..zipper.len() {
            let mut zip_file = zipper.by_index(idx)?;

            match zip_file.name() {
                "docProps/app.xml" => instance.app_info = Some(AppInfo::from_zip_file(&mut zip_file)?),
                "docProps/core.xml" => instance.core = Some(Core::from_zip_file(&mut zip_file)?),
                "word/document.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.main_document = Some(Box::new(Document::from_xml_element(&xml_node)?));
                }
                "word/_rels/document.xml.rels" => {
                    instance.main_document_relationships = zip_file_to_xml_node(&mut zip_file)?
                        .child_nodes
                        .iter()
                        .map(Relationship::from_xml_element)
                        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
                }
                "word/styles.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.styles = Some(Box::new(Styles::from_xml_element(&xml_node)?));
                }
                "word/settings.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.settings = Some(Box::new(Settings::from_xml_element(&xml_node)?));
                }
                "word/footnotes.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.footnotes = Some(Footnotes::from_xml_element(&xml_node)?);
                }
                "word/numbering.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.numbering = Some(Numbering::from_xml_element(&xml_node)?);
                }
                path if path.starts_with("word/media/") => instance.medias.push(PathBuf::from(file_path)),
                path if path.starts_with("word/theme/") => {
                    let file_stem = match Path::new(path).file_stem().and_then(OsStr::to_str).map(String::from) {
                        Some(name) => name,
                        None => {
                            error!("Couldn't get file name of theme");
                            continue;
                        }
                    };
                    let style_sheet = OfficeStyleSheet::from_xml_element(&zip_file_to_xml_node(&mut zip_file)?)?;
                    instance.themes.insert(file_stem, style_sheet);
                }
                _ => (),
            }
        }

        Ok(instance)
    }

    pub fn resolve_document_default_style(&self) -> Option<ResolvedStyle> {
        self.styles.as_ref()?.document_defaults.as_ref().map(|doc_defaults| {
            let run_properties = Box::new(
                doc_defaults
                    .run_properties_default
                    .as_ref()
                    .and_then(|r_pr_default| r_pr_default.0.as_ref())
                    .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
                    .unwrap_or_default(),
            );

            let paragraph_properties = Box::new(
                doc_defaults
                    .paragraph_properties_default
                    .as_ref()
                    .and_then(|p_pr_default| p_pr_default.0.as_ref())
                    .map(|p_pr| p_pr.base.clone())
                    .unwrap_or_default(),
            );

            ResolvedStyle {
                run_properties,
                paragraph_properties,
            }
        })
    }

    pub fn resolve_default_style(&self, style_type: StyleType) -> Option<ResolvedStyle> {
        let default_style =
            self.styles
                .as_ref()?
                .styles
                .iter()
                .find(|style| match (style.style_type, style.is_default) {
                    (Some(iter_style_type), Some(true)) => iter_style_type == style_type,
                    _ => false,
                })?;

        Some(ResolvedStyle::from_wml_style(default_style))
    }

    pub fn resolve_paragraph_style(&self, paragraph_properties: &PPr) -> Option<ResolvedStyle> {
        paragraph_properties
            .base
            .style
            .as_ref()
            .and_then(|style_name| self.resolve_style_with_id(style_name))
    }

    pub fn resolve_run_style(&self, run_properties: &RPr) -> Option<ResolvedStyle> {
        run_properties.r_pr_bases.iter().find_map(|r_pr_base| {
            if let RPrBase::RunStyle(style_name) = r_pr_base {
                self.resolve_style_with_id(style_name)
            } else {
                None
            }
        })
    }

    fn resolve_style_with_id<T: AsRef<str>>(&self, style_id: T) -> Option<ResolvedStyle> {
        // TODO(kalmar.robert) Use caching
        let styles = &self.styles.as_ref()?.styles;

        let top_most_style = styles.iter().find(|style| {
            style
                .style_id
                .as_ref()
                .filter(|s_id| (*s_id).as_str() == style_id.as_ref())
                .is_some()
        })?;

        let style_hierarchy: Vec<&Style> = std::iter::successors(Some(top_most_style), |child_style| {
            styles.iter().find(|style| style.style_id == child_style.based_on)
        })
        .collect();

        Some(
            style_hierarchy
                .iter()
                .rev()
                .fold(Default::default(), |mut resolved_style: ResolvedStyle, style| {
                    if let Some(style_p_pr) = &style.paragraph_properties {
                        *resolved_style.paragraph_properties =
                            resolved_style.paragraph_properties.update_with(style_p_pr.base.clone());
                    }

                    if let Some(style_r_pr) = &style.run_properties {
                        let folded_style_r_pr = RunProperties::from_vec(&style_r_pr.r_pr_bases);
                        *resolved_style.run_properties = resolved_style.run_properties.update_with(folded_style_r_pr);
                    }

                    resolved_style
                }),
        )
    }

    pub fn resolve_style_inheritance(&self, paragraph: &P, run: &R) -> Option<ResolvedStyle> {
        let paragraph_style = paragraph
            .properties
            .as_ref()
            .and_then(|p_pr| self.resolve_paragraph_style(p_pr))
            .or_else(|| self.resolve_default_style(StyleType::Paragraph));

        let run_style = run
            .run_properties
            .as_ref()
            .and_then(|r_pr| self.resolve_run_style(r_pr))
            .or_else(|| self.resolve_default_style(StyleType::Character));

        let calced_style = match (paragraph_style, run_style) {
            (Some(p_style), Some(r_style)) => Some(p_style.update_with_style_on_another_level(r_style)),
            (p_style, r_style) => p_style.or(r_style),
        };

        let default_style = self.resolve_document_default_style();
        let calced_style = match (default_style, calced_style) {
            (Some(def_style), Some(calced_style)) => Some(def_style.update_with(calced_style)),
            (def_style, calced_style) => def_style.or(calced_style),
        };

        calced_style.map(|resolved_style| {
            let run_style = run
                .run_properties
                .as_ref()
                .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases));

            match (paragraph.properties.as_ref(), run_style) {
                (Some(p_style), Some(r_style)) => resolved_style
                    .update_paragraph_with(p_style.base.clone())
                    .update_run_with(r_style),
                (Some(p_style), None) => resolved_style.update_paragraph_with(p_style.base.clone()),
                (None, Some(r_style)) => resolved_style.update_run_with(r_style),
                _ => resolved_style,
            }
        })
    }

    pub fn get_main_document_theme(&self) -> Option<&OfficeStyleSheet> {
        let theme_relation = self
            .main_document_relationships
            .iter()
            .find(|rel| rel.rel_type == THEME_RELATION_TYPE)?;

        let rel_target_file = Path::new(theme_relation.target.as_str())
            .file_stem()
            .and_then(OsStr::to_str)?;

        self.themes.get(rel_target_file)
    }

    pub fn get_main_document_section_properties(&self) -> Option<&SectPrContents> {
        self.main_document
            .as_ref()?
            .body
            .as_ref()?
            .section_properties
            .as_ref()?
            .contents
            .as_ref()
    }

    pub fn find_footnote_with_id(&self, id: i32) -> Option<&FtnEdn> {
        self.footnotes.as_ref()?.0.iter().find(|ftn_edn| ftn_edn.id == id)
    }

    pub fn resolve_footnote_style(&self, footnote_type: FtnEdnType) -> Option<ResolvedStyle> {
        self.footnotes
            .as_ref()?
            .0
            .iter()
            .find(|ftn_edn| ftn_edn.ftn_edn_type == Some(footnote_type))
            .and_then(|ftn_edn| {
                ftn_edn.block_level_elements.iter().find_map(|block_level_elt| {
                    if let BlockLevelElts::Chunk(ContentBlockContent::Paragraph(par)) = &block_level_elt {
                        Some(par)
                    } else {
                        None
                    }
                })
            })
            .map(|paragraph| {
                let run_properties = paragraph
                    .contents
                    .iter()
                    .find_map(|content| {
                        if let PContent::ContentRunContent(crc) = content {
                            if let ContentRunContent::Run(run) = &(**crc) {
                                return Some(
                                    run.run_properties
                                        .as_ref()
                                        .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
                                        .unwrap_or_default(),
                                );
                            }
                        }

                        None
                    })
                    .unwrap_or_default();

                if let Some(p_pr) = &paragraph.properties {
                    self.resolve_paragraph_style(p_pr)
                        .unwrap_or_default()
                        .update_run_with(run_properties)
                        .update_paragraph_with(p_pr.base.clone())
                } else {
                    ResolvedStyle::from_run_properties(Box::new(run_properties))
                }
            })
    }

    pub fn find_numbering_level(&self, numbering_id: i32, level: i32) -> Option<&Lvl> {
        if !(0..=8).contains(&level) {
            return None;
        }

        let numbering = self.numbering.as_ref()?;
        let num = numbering
            .numberings
            .iter()
            .find(|num| num.numbering_id == numbering_id)?;
        let abstract_num = numbering
            .abstract_numberings
            .iter()
            .find(|abstract_num| abstract_num.abstract_num_id == num.abstract_num_id)?;

        abstract_num.levels.iter().find(|lvl| lvl.level == level)
    }

    pub fn resolve_numbering_level_style(numbering_level: &Lvl) -> ResolvedStyle {
        let paragraph_properties = Box::new(
            numbering_level
                .paragraph_properties
                .as_ref()
                .map(|p_pr| p_pr.base.clone())
                .unwrap_or_default(),
        );

        let run_properties = Box::new(
            numbering_level
                .run_properties
                .as_ref()
                .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
                .unwrap_or_default(),
        );

        ResolvedStyle {
            paragraph_properties,
            run_properties,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        Package, RunProperties,
        super::{
            resolvedstyle::ParagraphProperties,
            wml::{
                document::{
                    BlockLevelElts, ContentBlockContent, ContentRunContent, Document, LineSpacingRule, PContent, PPr,
                    PPrBase, PPrGeneral, ParaRPr, RPr, RPrBase, RunInnerContent, SignedTwipsMeasure, Spacing,
                    TextAlignment, Underline, UnderlineType, P, R,
                },
                footnotes::{Footnotes, FtnEdn, FtnEdnType},
                settings::Settings,
                styles::{DocDefaults, PPrDefault, RPrDefault, Style, StyleType, Styles},
            }
        },
    };
    use crate::shared::docprops::{AppInfo, Core};

    #[test]
    #[ignore]
    fn test_size_of() {
        use std::mem::size_of;

        println!("sizeof Package: {}", size_of::<Package>());
        println!("sizeof AppInfo: {}", size_of::<AppInfo>());
        println!("sizeof Core: {}", size_of::<Core>());
        println!("sizeof Document: {}", size_of::<Document>());
        println!("sizeof Styles: {}", size_of::<Styles>());
        println!("sizeof Settings: {}", size_of::<Settings>());
    }

    fn doc_defaults_for_test() -> DocDefaults {
        let default_p_pr = PPr {
            base: PPrBase {
                start_on_next_page: Some(false),
                ..Default::default()
            },
            ..Default::default()
        };

        let default_r_pr = RPr {
            r_pr_bases: vec![RPrBase::Bold(true), RPrBase::Italic(false)],
            ..Default::default()
        };

        DocDefaults {
            paragraph_properties_default: Some(PPrDefault(Some(default_p_pr))),
            run_properties_default: Some(RPrDefault(Some(default_r_pr))),
        }
    }

    fn styles_for_test() -> Vec<Style> {
        let normal_style = Style {
            name: Some(String::from("Normal")),
            style_id: Some(String::from("Normal")),
            style_type: Some(StyleType::Paragraph),
            is_default: Some(true),
            paragraph_properties: Some(PPrGeneral {
                base: PPrBase {
                    start_on_next_page: Some(true),
                    ..Default::default()
                },
                ..Default::default()
            }),
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Italic(true)],
                ..Default::default()
            }),
            ..Default::default()
        };

        let child_style = Style {
            name: Some(String::from("Child")),
            style_id: Some(String::from("Child")),
            style_type: Some(StyleType::Paragraph),
            based_on: Some(String::from("Normal")),
            paragraph_properties: Some(PPrGeneral {
                base: PPrBase {
                    text_alignment: Some(TextAlignment::Center),
                    ..Default::default()
                },
                ..Default::default()
            }),
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Underline(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                })],
                ..Default::default()
            }),
            ..Default::default()
        };

        let default_par_style = Style {
            name: Some(String::from("DefaultParagraphFont")),
            style_id: Some(String::from("DefaultParagraphFont")),
            style_type: Some(StyleType::Character),
            ui_priority: Some(1),
            ..Default::default()
        };

        let emphasis_style = Style {
            name: Some(String::from("Emphasis")),
            style_id: Some(String::from("Emphasis")),
            style_type: Some(StyleType::Character),
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Italic(true)],
                ..Default::default()
            }),
            ..Default::default()
        };

        vec![normal_style, child_style, default_par_style, emphasis_style]
    }

    fn paragraph_with_style_for_test() -> P {
        P {
            properties: Some(PPr {
                base: PPrBase {
                    style: Some(String::from("Child")),
                    keep_lines_on_one_page: Some(true),
                    ..Default::default()
                },
                run_properties: Some(ParaRPr {
                    bases: vec![RPrBase::Bold(true), RPrBase::Italic(true)],
                    ..Default::default()
                }),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn run_with_style_for_test() -> R {
        R {
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::RunStyle(String::from("Emphasis"))],
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn package_for_test() -> Package {
        Package {
            styles: Some(Box::new(Styles {
                document_defaults: Some(doc_defaults_for_test()),
                latent_styles: None,
                styles: styles_for_test(),
            })),
            footnotes: Some(Footnotes(vec![FtnEdn {
                ftn_edn_type: Some(FtnEdnType::Separator),
                id: 0,
                block_level_elements: vec![BlockLevelElts::Chunk(ContentBlockContent::Paragraph(Box::new(P {
                    properties: Some(PPr {
                        base: PPrBase {
                            spacing: Some(Spacing {
                                line: Some(SignedTwipsMeasure::Decimal(240)),
                                line_rule: Some(LineSpacingRule::Auto),
                                ..Default::default()
                            }),
                            ..Default::default()
                        },
                        ..Default::default()
                    }),
                    contents: vec![PContent::ContentRunContent(Box::new(ContentRunContent::Run(R {
                        run_inner_contents: vec![RunInnerContent::Separator],
                        ..Default::default()
                    })))],
                    ..Default::default()
                })))],
            }])),
            ..Default::default()
        }
    }

    #[test]
    pub fn test_resolve_default_style() {
        let package = package_for_test();

        let default_style = package.resolve_document_default_style().unwrap();
        assert_eq!(
            *default_style.paragraph_properties,
            ParagraphProperties {
                start_on_next_page: Some(false),
                ..Default::default()
            },
        );
        assert_eq!(
            *default_style.run_properties,
            RunProperties {
                bold: Some(true),
                italic: Some(false),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_paragraph_style() {
        let package = package_for_test();

        let paragraph_style = package
            .resolve_paragraph_style(&paragraph_with_style_for_test().properties.unwrap())
            .unwrap();
        assert_eq!(
            *paragraph_style.paragraph_properties,
            ParagraphProperties {
                start_on_next_page: Some(true),
                text_alignment: Some(TextAlignment::Center),
                ..Default::default()
            }
        );

        assert_eq!(
            *paragraph_style.run_properties,
            RunProperties {
                italic: Some(true),
                underline: Some(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_run_style() {
        let package = package_for_test();

        let run_properties = package
            .resolve_run_style(&run_with_style_for_test().run_properties.unwrap())
            .unwrap();
        assert_eq!(
            *run_properties.run_properties,
            RunProperties {
                italic: Some(true),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_style_inheritance() {
        let package = package_for_test();

        let style = package
            .resolve_style_inheritance(&paragraph_with_style_for_test(), &run_with_style_for_test())
            .unwrap();
        assert_eq!(
            *style.paragraph_properties,
            ParagraphProperties {
                style: Some(String::from("Child")),
                start_on_next_page: Some(true),
                keep_lines_on_one_page: Some(true),
                text_alignment: Some(TextAlignment::Center),
                ..Default::default()
            }
        );
        assert_eq!(
            *style.run_properties,
            RunProperties {
                style: Some(String::from("Emphasis")),
                bold: Some(true),
                italic: Some(false),
                underline: Some(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_footnote_separator_style() {
        let package = package_for_test();
        let style = package.resolve_footnote_style(FtnEdnType::Separator).unwrap();
        assert_eq!(
            *style.paragraph_properties,
            ParagraphProperties {
                spacing: Some(Spacing {
                    line: Some(SignedTwipsMeasure::Decimal(240)),
                    line_rule: Some(LineSpacingRule::Auto),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }
}
