using System.Collections.Generic;

namespace Declaratest.Gui.Models;

public class TestData
{
    public string Subject { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Name { get; set; } = string.Empty;
    public SectionType Type { get; set; } = SectionType.Short;
    public List<Question> Questions { get; set; } = new();
    public bool SeparateSheet { get; set; }
    public string? Subtitle { get; set; }
}

public enum SectionType
{
    Short,
    Long,
    MatchingV,
    MatchingH,
    Blanks,
    Oral
}

public abstract class Question
{
    public string Text { get; set; } = string.Empty;
}

public class TextQuestion : Question
{
    public int? Lines { get; set; }
}

public class MatchingQuestion : Question
{
    public string Left { get; set; } = string.Empty;
    public string Right { get; set; } = string.Empty;
}

public class BlankQuestion : Question
{
}

public class OralQuestion : Question
{
    public List<string> SubPoints { get; set; } = new();
}
