using System.IO;
using System.Threading.Tasks;
using Win32OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Win32SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace Declaratest.Gui.Services;

public class FileService : IFileService
{
    public Task<string?> OpenFileAsync()
    {
        var dialog = new Win32OpenFileDialog
        {
            Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
            Title = "Open Declaratest Markdown"
        };

        if (dialog.ShowDialog() == true && File.Exists(dialog.FileName))
        {
            var content = File.ReadAllText(dialog.FileName);
            return Task.FromResult<string?>(content);
        }

        return Task.FromResult<string?>(null);
    }

    public Task<string?> SaveFileAsync(string content, string? existingPath = null)
    {
        var path = existingPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            var dialog = new Win32SaveFileDialog
            {
                Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*",
                Title = "Save Declaratest Markdown",
                AddExtension = true,
                DefaultExt = "md"
            };

            if (dialog.ShowDialog() == true)
            {
                path = dialog.FileName;
            }
        }

        if (!string.IsNullOrWhiteSpace(path))
        {
            File.WriteAllText(path, content);
            return Task.FromResult<string?>(path);
        }

        return Task.FromResult<string?>(null);
    }
}
