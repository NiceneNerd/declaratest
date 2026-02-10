using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;
using Declaratest.Gui.Helpers;
using Declaratest.Gui.Models;
using Declaratest.Gui.Services;

namespace Declaratest.Gui.ViewModels;

public class MainViewModel : INotifyPropertyChanged
{
    private readonly IMarkdownParser _parser;
    private readonly IDocxGenerator _generator;
    private readonly IPreviewService _previewService;
    private readonly IFileService _fileService;

    private string _documentText = DefaultTemplate;
    private string _statusMessage = "Ready";
    private string _previewStatus = string.Empty;
    private string _previewMessage = string.Empty;
    private string? _currentFile;
    private bool _isBusy;

    public MainViewModel(IMarkdownParser parser, IDocxGenerator generator, IPreviewService previewService, IFileService fileService)
    {
        _parser = parser;
        _generator = generator;
        _previewService = previewService;
        _fileService = fileService;

        PreviewStatus = _previewService.IsAvailable ? "Preview: Microsoft Word detected" : "Preview: Word not available";
        PreviewMessage = _previewService.IsAvailable
            ? "Live preview will use embedded Word."
            : "Install Microsoft Word to enable live preview.";

        NewCommand = new RelayCommand(_ => ResetDocument(), _ => !IsBusy);
        OpenCommand = new RelayCommand(OpenDocumentAsync, _ => !IsBusy, HandleError);
        SaveCommand = new RelayCommand(SaveDocumentAsync, _ => !IsBusy, HandleError);
        SaveAsCommand = new RelayCommand(SaveDocumentAsAsync, _ => !IsBusy, HandleError);
        GeneratePreviewCommand = new RelayCommand(GeneratePreviewAsync, _ => !IsBusy, HandleError);
        TogglePreviewCommand = new RelayCommand(_ => TogglePreviewMessage());
        ResetZoomCommand = new RelayCommand(_ => StatusMessage = "Zoom reset to 100%");
        ShowAboutCommand = new RelayCommand(_ => StatusMessage = "Declaratest GUI - Markdown to DOCX preview");
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event EventHandler<string>? PreviewRequested;

    public string DocumentText
    {
        get => _documentText;
        set => SetField(ref _documentText, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetField(ref _statusMessage, value);
    }

    public string PreviewStatus
    {
        get => _previewStatus;
        set => SetField(ref _previewStatus, value);
    }

    public string PreviewMessage
    {
        get => _previewMessage;
        set => SetField(ref _previewMessage, value);
    }

    public ICommand NewCommand { get; }
    public ICommand OpenCommand { get; }
    public ICommand SaveCommand { get; }
    public ICommand SaveAsCommand { get; }
    public ICommand GeneratePreviewCommand { get; }
    public ICommand TogglePreviewCommand { get; }
    public ICommand ResetZoomCommand { get; }
    public ICommand ShowAboutCommand { get; }
    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetField(ref _isBusy, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    private void ResetDocument()
    {
        DocumentText = DefaultTemplate;
        StatusMessage = "New document started.";
        _currentFile = null;
    }

    private async Task OpenDocumentAsync(object? parameter)
    {
        await WithBusy(async () =>
        {
            var content = await _fileService.OpenFileAsync();
            if (content != null)
            {
                DocumentText = content;
                StatusMessage = "File loaded.";
            }
        });
    }

    private Task SaveDocumentAsync(object? parameter) => SaveDocumentInternalAsync(forcePrompt: false);

    private Task SaveDocumentAsAsync(object? parameter) => SaveDocumentInternalAsync(forcePrompt: true);

    private async Task GeneratePreviewAsync(object? parameter)
    {
        await WithBusy(async () =>
        {
            var parsed = _parser.Parse(DocumentText);
            var output = await _generator.GenerateAsync(parsed);
            PreviewRequested?.Invoke(this, output);
            StatusMessage = "Preview refreshed.";
        });
    }

    private void TogglePreviewMessage()
    {
        PreviewMessage = _previewService.IsAvailable
            ? "Live preview ready."
            : "Preview disabled until Word is installed.";
    }

    private async Task WithBusy(Func<Task> work)
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;
        try
        {
            await work();
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task SaveDocumentInternalAsync(bool forcePrompt)
    {
        await WithBusy(async () =>
        {
            var savedPath = await _fileService.SaveFileAsync(DocumentText, forcePrompt ? null : _currentFile);
            if (!string.IsNullOrWhiteSpace(savedPath))
            {
                _currentFile = savedPath;
                StatusMessage = "File saved.";
            }
        });
    }

    private void HandleError(Exception ex)
    {
        StatusMessage = $"Error: {ex.Message}";
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }

    private const string DefaultTemplate =
@"# Declaratest GUI Preview
Subject: General Knowledge
Title: Declaratest GUI Preview

## Section: Quick Facts
Type: short
- What is the most common element in the Earth's atmosphere?
- Name a fruit that is both red and green when ripe.

## Section: Match the Term
Type: matching_v
- photosynthesis -> Process by which plants make food
- alloy -> Mixture of metals

## Section: Fill in the Blanks
Type: blanks
- The largest planet in our solar system is ___________.

## Section: Oral Questions
Type: oral
- What are the main causes of climate change?
    - What happens during evaporation?
    - How do clouds form?
";
}
