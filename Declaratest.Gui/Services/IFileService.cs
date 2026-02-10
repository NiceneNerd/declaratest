using System.Threading.Tasks;

namespace Declaratest.Gui.Services;

public interface IFileService
{
    Task<string?> OpenFileAsync();

    Task<string?> SaveFileAsync(string content, string? existingPath = null);
}
