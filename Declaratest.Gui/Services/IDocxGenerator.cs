using System.Threading;
using System.Threading.Tasks;
using Declaratest.Gui.Models;

namespace Declaratest.Gui.Services;

public interface IDocxGenerator
{
    Task<string> GenerateAsync(TestData data, string? templatePath = null, CancellationToken cancellationToken = default);
}
