using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Declaratest.Gui.Services;
using Declaratest.Gui.ViewModels;

namespace Declaratest.Gui.Views;

public partial class MainWindow : Window
{
    private readonly IPreviewService _previewService;

    public MainWindow()
    {
        InitializeComponent();

        var parser = new MarkdownParser();
        var generator = new DocxGenerator();
        _previewService = new WordInteropPreviewService();
        var fileService = new FileService();

        var viewModel = new MainViewModel(parser, generator, _previewService, fileService);
        DataContext = viewModel;

        viewModel.PreviewRequested += OnPreviewRequested;
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Close, (_, _) => Close()));
    }

    private async void OnPreviewRequested(object? sender, string documentPath)
    {
        try
        {
            await _previewService.LoadDocumentAsync(documentPath, PreviewHost);
        }
        catch (Exception ex)
        {
            if (DataContext is MainViewModel vm)
            {
                vm.StatusMessage = $"Preview failed: {ex.Message}";
            }
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        _previewService.Dispose();
    }
}
