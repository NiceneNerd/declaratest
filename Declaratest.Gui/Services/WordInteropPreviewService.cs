using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.Integration;
using WF = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Declaratest.Gui.Services;

public class WordInteropPreviewService : IPreviewService
{
    private readonly bool _wordInstalled;
    private Word.Application? _application;
    private Word.Document? _document;

    public WordInteropPreviewService()
    {
        _wordInstalled = Type.GetTypeFromProgID("Word.Application") != null;
    }

    public bool IsAvailable => _wordInstalled;

    public Task LoadDocumentAsync(string documentPath, WindowsFormsHost host, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!IsAvailable)
        {
            host.Child = new WF.Label
            {
                Text = "Microsoft Word is required for live preview.",
                Dock = WF.DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            return Task.CompletedTask;
        }

        EnsureApplication();
        OpenDocument(documentPath);
        AttachToHost(host);

        return Task.CompletedTask;
    }

    private void EnsureApplication()
    {
        _application ??= new Word.Application
        {
            DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
        };
    }

    private void OpenDocument(string path)
    {
        if (_application == null)
        {
            return;
        }

        _document?.Close(false);
        try
        {
            _document = _application.Documents.Open(path, ReadOnly: true);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"Unable to open Word preview: {ex.Message}", ex);
        }

        _application.Visible = true;
    }

    private void AttachToHost(WindowsFormsHost host)
    {
        if (host.Child is not WF.Panel panel)
        {
            panel = new WF.Panel { Dock = WF.DockStyle.Fill };
            host.Child = panel;
        }

        var wordWindow = _application?.ActiveWindow;
        if (wordWindow == null)
        {
            return;
        }

        var wordHandle = new IntPtr(wordWindow.Hwnd);
        if (wordHandle == IntPtr.Zero)
        {
            host.Child = new WF.Label
            {
                Text = "Unable to locate the Word preview window.",
                Dock = WF.DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            return;
        }

        var setParentResult = SetParent(wordHandle, panel.Handle);
        if (setParentResult == IntPtr.Zero)
        {
            host.Child = new WF.Label
            {
                Text = "Unable to embed the Word preview window.",
                Dock = WF.DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            return;
        }

        wordWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;
    }

    public void Dispose()
    {
        try
        {
            _document?.Close(false);
            if (_document != null)
            {
                Marshal.FinalReleaseComObject(_document);
            }

            _application?.Quit(false);
            if (_application != null)
            {
                Marshal.FinalReleaseComObject(_application);
            }
        }
        finally
        {
            _document = null;
            _application = null;
        }
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
}
