using System;
using System.Globalization;
using Declaratest.Gui.Models;

namespace Declaratest.Gui.Services;

public class MarkdownParser : IMarkdownParser
{
    private const string SubPointPrefix = "    -";

    public TestData Parse(string markdown)
    {
        var data = new TestData();
        Section? currentSection = null;
        Question? lastQuestion = null;

        foreach (var rawLine in markdown.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
        {
            var line = rawLine.TrimEnd();
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (line.StartsWith("# "))
            {
                data.Title = line[2..].Trim();
                continue;
            }

            if (line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
            {
                data.Subject = line[8..].Trim();
                continue;
            }

            if (line.StartsWith("Title:", StringComparison.OrdinalIgnoreCase))
            {
                data.Title = line[6..].Trim();
                continue;
            }

            if (line.StartsWith("## Section:", StringComparison.OrdinalIgnoreCase))
            {
                if (currentSection != null)
                {
                    data.Sections.Add(currentSection);
                }

                currentSection = new Section
                {
                    Name = line["## Section:".Length..].Trim()
                };
                lastQuestion = null;
                continue;
            }

            if (line.StartsWith("Type:", StringComparison.OrdinalIgnoreCase) && currentSection != null)
            {
                var typeValue = line["Type:".Length..].Trim();
                if (Enum.TryParse<SectionType>(NormalizeSectionType(typeValue), ignoreCase: true, out var sectionType))
                {
                    currentSection.Type = sectionType;
                }
                continue;
            }

            if (line.StartsWith("Separate Sheet", StringComparison.OrdinalIgnoreCase) && currentSection != null)
            {
                currentSection.SeparateSheet = line.Contains("yes", StringComparison.OrdinalIgnoreCase);
                continue;
            }

            if (line.StartsWith("- "))
            {
                if (currentSection == null)
                {
                    continue;
                }

                var content = line[2..].Trim();
                lastQuestion = ParseQuestion(currentSection, content);
                if (lastQuestion != null)
                {
                    currentSection.Questions.Add(lastQuestion);
                }
                continue;
            }

            if (rawLine.StartsWith(SubPointPrefix) && currentSection?.Type == SectionType.Oral && lastQuestion is OralQuestion oral)
            {
                var subPoint = rawLine[SubPointPrefix.Length..].Trim();
                if (!string.IsNullOrWhiteSpace(subPoint))
                {
                    oral.SubPoints.Add(subPoint);
                }
            }
        }

        if (currentSection != null)
        {
            data.Sections.Add(currentSection);
        }

        return data;
    }

    private static string NormalizeSectionType(string typeValue)
    {
        return typeValue.ToLower(CultureInfo.InvariantCulture) switch
        {
            "matching_v" => nameof(SectionType.MatchingV),
            "matching_h" => nameof(SectionType.MatchingH),
            "short" => nameof(SectionType.Short),
            "long" => nameof(SectionType.Long),
            "blanks" => nameof(SectionType.Blanks),
            "oral" => nameof(SectionType.Oral),
            _ => typeValue
        };
    }

    private static Question? ParseQuestion(Section section, string content)
    {
        return section.Type switch
        {
            SectionType.MatchingV or SectionType.MatchingH => ParseMatchingQuestion(content),
            SectionType.Blanks => new BlankQuestion { Text = content },
            SectionType.Oral => new OralQuestion { Text = content },
            SectionType.Long => new TextQuestion { Text = content, Lines = 6 },
            _ => new TextQuestion { Text = content, Lines = 2 }
        };
    }

    private static MatchingQuestion ParseMatchingQuestion(string content)
    {
        var parts = content.Split(new[] { "->" }, 2, StringSplitOptions.RemoveEmptyEntries);
        var left = parts.Length > 0 ? parts[0].Trim() : string.Empty;
        var right = parts.Length > 1 ? parts[1].Trim() : string.Empty;
        return new MatchingQuestion { Left = left, Right = right, Text = content };
    }
}
