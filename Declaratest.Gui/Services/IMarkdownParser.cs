using Declaratest.Gui.Models;

namespace Declaratest.Gui.Services;

public interface IMarkdownParser
{
    TestData Parse(string markdown);
}
