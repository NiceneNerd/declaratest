use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MatchingQuestion {
    pub left: String,
    pub right: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextQuestion {
    pub text: String,
    pub lines: Option<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BlankQuestion {
    pub text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OralQuestion {
    pub text: String,
    pub sub_points: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum Question {
    Matching(MatchingQuestion),
    Text(TextQuestion),
    Blank(BlankQuestion),
    Oral(OralQuestion),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SectionType {
    Short,
    Long,
    MatchingV,
    MatchingH,
    Blanks,
    Oral,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Section {
    pub name: String,
    pub questions: Vec<Question>,
    pub section_type: Option<SectionType>,
    pub separate_sheet: bool,
    pub subtitle: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TestData {
    pub subject: String,
    pub title: String,
    pub sections: Vec<Section>,
}
