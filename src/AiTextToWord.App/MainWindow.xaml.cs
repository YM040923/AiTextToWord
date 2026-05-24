using AiTextToWord.Core.Conversion;
using AiTextToWord.Core.Model;
using AiTextToWord.Docx;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace AiTextToWord.App;

public sealed partial class MainWindow : Window
{
    private readonly TextConversionService conversionService = TextConversionService.CreateDefault();
    private ConversionResult? currentResult;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        ConvertInput();
    }

    private async void Paste_Click(object sender, RoutedEventArgs e)
    {
        var package = Clipboard.GetContent();
        if (package.Contains(StandardDataFormats.Text))
        {
            InputTextBox.Text = await package.GetTextAsync();
        }
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        InputTextBox.Text = string.Empty;
        PreviewPanel.Children.Clear();
        StatusText.Text = string.Empty;
        currentResult = null;
    }

    private async void ExportWord_Click(object sender, RoutedEventArgs e)
    {
        if (currentResult is null || currentResult.Document.Blocks.Count == 0)
        {
            StatusText.Text = "Paste text before exporting.";
            return;
        }

        var picker = new FileSavePicker
        {
            SuggestedFileName = "ai-text-export"
        };
        picker.FileTypeChoices.Add("Word document", [".docx"]);
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSaveFileAsync();
        if (file is null)
        {
            return;
        }

        await using var stream = File.Create(file.Path);
        stream.SetLength(0);
        new DocxExporter().Export(currentResult.Document, stream, new DocxExportOptions("AI Text Export"));
        StatusText.Text = "Exported Word document.";
    }

    private void ConvertInput()
    {
        if (string.IsNullOrWhiteSpace(InputTextBox.Text))
        {
            PreviewPanel.Children.Clear();
            StatusText.Text = string.Empty;
            currentResult = null;
            return;
        }

        currentResult = conversionService.Convert(InputTextBox.Text);
        RenderPreview(currentResult.Document);
        StatusText.Text = $"{currentResult.Document.Blocks.Count} blocks ready";
    }

    private void RenderPreview(DocumentModel document)
    {
        PreviewPanel.Children.Clear();

        foreach (var block in document.Blocks)
        {
            PreviewPanel.Children.Add(CreatePreviewElement(block));
        }
    }

    private static FrameworkElement CreatePreviewElement(DocumentBlock block)
    {
        return block switch
        {
            HeadingBlock heading => new TextBlock
            {
                Text = heading.Text,
                FontSize = heading.Level == 1 ? 24 : heading.Level == 2 ? 20 : 17,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap
            },
            ParagraphBlock paragraph => PreviewText(paragraph.Text),
            BlockQuoteBlock quote => PreviewText(quote.Text, 0.72),
            CodeBlock code => new TextBox
            {
                Text = code.Code,
                FontFamily = new FontFamily("Consolas"),
                IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap
            },
            DividerBlock => new Border { Height = 1, Opacity = 0.3 },
            ListBlock list => PreviewText(string.Join(Environment.NewLine, list.Items.Select(item => $"{(list.IsOrdered ? "1." : "-")} {item}"))),
            _ => PreviewText(string.Empty)
        };
    }

    private static TextBlock PreviewText(string text, double opacity = 1)
    {
        return new TextBlock
        {
            Text = text,
            Opacity = opacity,
            TextWrapping = TextWrapping.Wrap
        };
    }
}
