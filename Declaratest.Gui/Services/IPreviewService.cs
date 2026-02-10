using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.Integration;

namespace Declaratest.Gui.Services;

public interface IPreviewService : IDisposable
{
    bool IsAvailable { get; }

    Task LoadDocumentAsync(string documentPath, WindowsFormsHost host, CancellationToken cancellationToken = default);
}
