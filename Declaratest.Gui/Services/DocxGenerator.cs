using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Declaratest.Gui.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Declaratest.Gui.Services;

public class DocxGenerator : IDocxGenerator
{
    private static bool _cleanupDone;
    private string? _lastOutputPath;

    public DocxGenerator()
    {
        if (!_cleanupDone)
        {
            CleanupOldTempFiles();
            _cleanupDone = true;
        }
    }

    public async Task<string> GenerateAsync(TestData data, string? templatePath = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DeletePreviousOutput();
        var outputPath = Path.Combine(Path.GetTempPath(), $"declaratest_{Guid.NewGuid():N}.docx");

        await Task.Run(() =>
        {
            CreateDocument(data, outputPath);
        }, cancellationToken);

        _lastOutputPath = outputPath;
        return outputPath;
    }

    private void DeletePreviousOutput()
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath))
        {
            return;
        }

        try
        {
            if (File.Exists(_lastOutputPath))
            {
                File.Delete(_lastOutputPath);
            }
        }
        catch (IOException)
        {
            // ignore cleanup issues
        }
    }

    private static void CleanupOldTempFiles()
    {
        try
        {
            var directory = new DirectoryInfo(Path.GetTempPath());
            foreach (var file in directory.GetFiles("declaratest_*.docx"))
            {
                if (file.CreationTimeUtc < DateTime.UtcNow.AddHours(-12))
                {
                    file.Delete();
                }
            }
        }
        catch (IOException)
        {
            // ignore cleanup issues
        }
        catch (UnauthorizedAccessException)
        {
            // ignore cleanup issues
        }
    }

    private static void CreateDocument(TestData data, string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        if (!string.IsNullOrWhiteSpace(data.Subject))
        {
            body.Append(CreateParagraph($"Subject: {data.Subject}", bold: true));
        }

        if (!string.IsNullOrWhiteSpace(data.Title))
        {
            body.Append(CreateParagraph(data.Title, bold: true, larger: true));
        }

        foreach (var section in data.Sections)
        {
            body.Append(CreateParagraph(section.Name, bold: true));
            if (!string.IsNullOrWhiteSpace(section.Subtitle))
            {
                body.Append(CreateParagraph(section.Subtitle));
            }

            var number = 1;
            foreach (var question in section.Questions)
            {
                body.Append(CreateParagraph($"{number}. {RenderQuestion(question)}"));

                if (question is TextQuestion { Lines: var lines } && lines > 0)
                {
                    for (var i = 0; i < lines; i++)
                    {
                        body.Append(CreateParagraph(string.Empty));
                    }
                }

                number++;
            }
        }

        mainPart.Document.Save();
    }

    private static string RenderQuestion(Question question) =>
        question switch
        {
            MatchingQuestion matching => $"{matching.Left} -> {matching.Right}",
            OralQuestion oral when oral.SubPoints.Count > 0 =>
                $"{oral.Text}{Environment.NewLine}  - {string.Join(Environment.NewLine + "  - ", oral.SubPoints)}",
            _ => question.Text
        };

    private static Paragraph CreateParagraph(string text, bool bold = false, bool larger = false)
    {
        var run = new Run();
        var runProperties = new RunProperties();

        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (larger)
        {
            runProperties.Append(new FontSize { Val = "32" });
        }

        if (runProperties.HasChildren)
        {
            run.Append(runProperties);
        }

        run.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(run);
    }
}
