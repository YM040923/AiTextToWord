using AiTextToWord.Core.Conversion;
using AiTextToWord.Core.Export;
using AiTextToWord.Core.Fonts;
using AiTextToWord.Core.Model;
using AiTextToWord.Docx;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.UI.Windowing;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Text;
using WinRT.Interop;

namespace AiTextToWord.App;

public sealed partial class MainWindow : Window
{
    private readonly TextConversionService conversionService = TextConversionService.CreateDefault();
    private ConversionResult? currentResult;
    private IReadOnlyList<string> installedFonts = [];
    private bool isLoadingExportSettings;
    private static readonly double[] BodyFontSizes = [10.5, 11, 12, 14];
    private static readonly double[] HeadingFontSizes = [16, 18, 20, 22];
    private static readonly double[] LineSpacings = [1.15, 1.3, 1.5, 2.0];
    private const string ExportFormatSettingKey = "Export.Format";
    private const string PresetSettingKey = "Export.Preset";
    private const string FontSettingKey = "Export.Font";
    private const string BodyFontSizeSettingKey = "Export.BodyFontSize";
    private const string HeadingFontSizeSettingKey = "Export.HeadingFontSize";
    private const string LineSpacingSettingKey = "Export.LineSpacing";
    private const string PageMarginSettingKey = "Export.PageMargin";
    private const string QuoteStyleSettingKey = "Export.QuoteStyle";
    private const string ListDensitySettingKey = "Export.ListDensity";
    private const string LastExportFolderSettingKey = "Export.LastFolder";
    private const string SampleText = """
        # AI 文本整理示例

        这是一段从 AI 对话中复制出来的 Markdown 内容。它包含**加粗文字**、`行内代码`、列表和引用块。

        ## 可以识别的结构

        - 标题会变成 Word 标题样式
        - 列表会保持项目符号
        - 引用会整理成独立段落

        > 这是一段引用内容，适合放提示词、说明、摘录或 AI 给出的重点结论。

        ```csharp
        Console.WriteLine("Hello, Word");
        ```

        你可以修改导出设置，然后导出 Word 查看效果。
        """;

    public MainWindow()
    {
        InitializeComponent();
        isLoadingExportSettings = true;
        ConfigureWindowChrome();
        LoadInstalledFonts();
        LoadExportSettings();
        isLoadingExportSettings = false;
        SystemBackdrop = new MicaBackdrop();
        UpdateSettingsSummary();
        UpdateExportButton();
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
            StatusText.Text = "已粘贴剪贴板文本。";
        }
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        InputTextBox.Text = string.Empty;
        PreviewPanel.Children.Clear();
        PreviewScrollViewer.Visibility = Visibility.Collapsed;
        PreviewPage.Visibility = Visibility.Collapsed;
        EmptyPreviewPanel.Visibility = Visibility.Visible;
        InputMetaText.Text = "等待粘贴内容";
        PreviewMetaText.Text = "暂无文档块";
        StatusText.Text = "就绪";
        currentResult = null;
    }

    private void Sample_Click(object sender, RoutedEventArgs e)
    {
        InputTextBox.Text = SampleText;
        StatusText.Text = "已载入示例文本。";
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (currentResult is null || currentResult.Document.Blocks.Count == 0)
        {
            StatusText.Text = "请先粘贴文本再导出。";
            return;
        }

        var picker = new FileSavePicker
        {
            SettingsIdentifier = ExportLocationMemory.PickerSettingsIdentifier,
            SuggestedFileName = "AI文本导出",
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary
        };
        if (IsPdfExportSelected())
        {
            picker.FileTypeChoices.Add("PDF 文档", [".pdf"]);
        }
        else
        {
            picker.FileTypeChoices.Add("Word 文档", [".docx"]);
        }

        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSaveFileAsync();
        if (file is null)
        {
            return;
        }

        await using var stream = File.Create(file.Path);
        stream.SetLength(0);
        if (IsPdfExportSelected())
        {
            new PdfExporter().Export(currentResult.Document, stream, CreateExportOptions());
        }
        else
        {
            new DocxExporter().Export(currentResult.Document, stream, CreateExportOptions());
        }

        SaveLastExportFolder(file.Path);
        StatusText.Text = $"{SelectedExportFormatName()} 文档已导出：{file.Name}";
    }

    private void PresetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!AreExportSettingControlsReady())
        {
            return;
        }

        if (isLoadingExportSettings)
        {
            return;
        }

        switch (PresetComboBox.SelectedIndex)
        {
            case 1:
                SelectPreferredFont();
                BodyFontSizeComboBox.SelectedIndex = 0;
                HeadingFontSizeComboBox.SelectedIndex = 0;
                LineSpacingComboBox.SelectedIndex = 0;
                PageMarginComboBox.SelectedIndex = 0;
                QuoteStyleComboBox.SelectedIndex = 0;
                ListDensityComboBox.SelectedIndex = 0;
                break;
            case 2:
                SelectPreferredFont();
                BodyFontSizeComboBox.SelectedIndex = 2;
                HeadingFontSizeComboBox.SelectedIndex = 3;
                LineSpacingComboBox.SelectedIndex = 2;
                PageMarginComboBox.SelectedIndex = 2;
                QuoteStyleComboBox.SelectedIndex = 1;
                ListDensityComboBox.SelectedIndex = 1;
                break;
            default:
                SelectPreferredFont();
                BodyFontSizeComboBox.SelectedIndex = 1;
                HeadingFontSizeComboBox.SelectedIndex = 2;
                LineSpacingComboBox.SelectedIndex = 1;
                PageMarginComboBox.SelectedIndex = 1;
                QuoteStyleComboBox.SelectedIndex = 0;
                ListDensityComboBox.SelectedIndex = 0;
                break;
        }

        UpdateSettingsSummary();
        UpdateExportButton();
        RefreshPreview();
        SaveExportSettings();
    }

    private void ExportSettings_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!AreExportSettingControlsReady())
        {
            return;
        }

        if (isLoadingExportSettings)
        {
            return;
        }

        UpdateSettingsSummary();
        UpdateExportButton();
        RefreshPreview();
        SaveExportSettings();
    }

    private DocxExportOptions CreateExportOptions()
    {
        return new DocxExportOptions("AI Text Export")
        {
            Preset = PresetComboBox.SelectedIndex switch
            {
                1 => DocxStylePreset.CompactNotes,
                2 => DocxStylePreset.FormalReport,
                _ => DocxStylePreset.StandardDocument
            },
            FontFamily = SelectedFontFamily(),
            BodyFontSize = SelectedDouble(BodyFontSizes, BodyFontSizeComboBox.SelectedIndex, 11),
            HeadingFontSize = SelectedDouble(HeadingFontSizes, HeadingFontSizeComboBox.SelectedIndex, 20),
            LineSpacing = SelectedDouble(LineSpacings, LineSpacingComboBox.SelectedIndex, 1.3),
            PageMargin = PageMarginComboBox.SelectedIndex switch
            {
                0 => DocxPageMargin.Narrow,
                2 => DocxPageMargin.Wide,
                _ => DocxPageMargin.Standard
            },
            QuoteStyle = QuoteStyleComboBox.SelectedIndex == 1 ? DocxQuoteStyle.GrayBlock : DocxQuoteStyle.Indented,
            ListDensity = ListDensityComboBox.SelectedIndex == 1 ? DocxListDensity.Comfortable : DocxListDensity.Compact
        };
    }

    private void UpdateSettingsSummary()
    {
        if (!AreExportSettingControlsReady())
        {
            return;
        }

        var preset = PresetComboBox.SelectedIndex switch
        {
            1 => "紧凑笔记",
            2 => "正式报告",
            _ => "标准文档"
        };
        var font = SelectedFontFamily();
        var bodySize = SelectedDouble(BodyFontSizes, BodyFontSizeComboBox.SelectedIndex, 11);
        var lineSpacing = SelectedDouble(LineSpacings, LineSpacingComboBox.SelectedIndex, 1.3);
        SettingsSummaryText.Text = $"{SelectedExportFormatName()} · {preset} · {font} · {bodySize:0.#}pt · {lineSpacing:0.##} 倍行距";
    }

    private void UpdateExportButton()
    {
        if (ExportButtonText is not null)
        {
            ExportButtonText.Text = $"导出 {SelectedExportFormatName()}";
        }
    }

    private bool IsPdfExportSelected()
    {
        return ExportFormatComboBox.SelectedIndex == 1;
    }

    private string SelectedExportFormatName()
    {
        return IsPdfExportSelected() ? "PDF" : "Word";
    }

    private void LoadInstalledFonts()
    {
        installedFonts = InstalledFontProvider.GetInstalledFontFamilies();
        UpdateFontSuggestions(string.Empty);
        SelectPreferredFont();
    }

    private void SelectPreferredFont()
    {
        var preferredFont = FindPreferredFont();
        if (!string.IsNullOrWhiteSpace(preferredFont))
        {
            FontSearchBox.Text = preferredFont;
        }
    }

    private string SelectedFontFamily()
    {
        var typedFont = FontSearchBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(typedFont))
        {
            return FindInstalledFont(typedFont) ?? typedFont;
        }

        return FindPreferredFont() ?? "Microsoft YaHei";
    }

    private string? FindPreferredFont()
    {
        string[] preferredFonts =
        [
            "微软雅黑",
            "Microsoft YaHei",
            "Microsoft YaHei UI",
            "等线",
            "DengXian",
            "宋体",
            "SimSun",
            "Arial"
        ];

        foreach (var preferredFont in preferredFonts)
        {
            var match = installedFonts.FirstOrDefault(font =>
                string.Equals(font, preferredFont, StringComparison.CurrentCultureIgnoreCase));
            if (!string.IsNullOrWhiteSpace(match))
            {
                return match;
            }
        }

        return installedFonts.FirstOrDefault();
    }

    private void FontSearchBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (!AreExportSettingControlsReady())
        {
            return;
        }

        if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
        {
            UpdateFontSuggestions(sender.Text);
            UpdateSettingsSummary();
            RefreshPreview();
            SaveExportSettings();
        }
    }

    private void FontSearchBox_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
    {
        if (args.SelectedItem is not string font)
        {
            return;
        }

        sender.Text = font;
        UpdateSettingsSummary();
        RefreshPreview();
        SaveExportSettings();
    }

    private void FontSearchBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        var font = args.ChosenSuggestion as string
            ?? FontFamilySearch.Filter(installedFonts, args.QueryText, 1).FirstOrDefault()
            ?? FindPreferredFont();

        if (!string.IsNullOrWhiteSpace(font))
        {
            sender.Text = font;
        }

        UpdateSettingsSummary();
        RefreshPreview();
        SaveExportSettings();
    }

    private void UpdateFontSuggestions(string query)
    {
        FontSearchBox.ItemsSource = FontFamilySearch.Filter(installedFonts, query);
    }

    private string? FindInstalledFont(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return null;
        }

        return installedFonts.FirstOrDefault(font =>
            string.Equals(font, fontName.Trim(), StringComparison.CurrentCultureIgnoreCase));
    }

    private void LoadExportSettings()
    {
        var settings = ApplicationData.Current.LocalSettings.Values;

        ExportFormatComboBox.SelectedIndex = ReadIndexSetting(settings, ExportFormatSettingKey, 0, ExportFormatComboBox.Items.Count);
        PresetComboBox.SelectedIndex = ReadIndexSetting(settings, PresetSettingKey, 0, PresetComboBox.Items.Count);
        BodyFontSizeComboBox.SelectedIndex = ReadIndexSetting(settings, BodyFontSizeSettingKey, 1, BodyFontSizeComboBox.Items.Count);
        HeadingFontSizeComboBox.SelectedIndex = ReadIndexSetting(settings, HeadingFontSizeSettingKey, 2, HeadingFontSizeComboBox.Items.Count);
        LineSpacingComboBox.SelectedIndex = ReadIndexSetting(settings, LineSpacingSettingKey, 1, LineSpacingComboBox.Items.Count);
        PageMarginComboBox.SelectedIndex = ReadIndexSetting(settings, PageMarginSettingKey, 1, PageMarginComboBox.Items.Count);
        QuoteStyleComboBox.SelectedIndex = ReadIndexSetting(settings, QuoteStyleSettingKey, 0, QuoteStyleComboBox.Items.Count);
        ListDensityComboBox.SelectedIndex = ReadIndexSetting(settings, ListDensitySettingKey, 0, ListDensityComboBox.Items.Count);

        if (settings.TryGetValue(FontSettingKey, out var value) && value is string savedFont)
        {
            FontSearchBox.Text = FindInstalledFont(savedFont) ?? savedFont;
            UpdateFontSuggestions(FontSearchBox.Text);
        }
    }

    private void SaveExportSettings()
    {
        if (isLoadingExportSettings || !AreExportSettingControlsReady())
        {
            return;
        }

        var settings = ApplicationData.Current.LocalSettings.Values;
        settings[ExportFormatSettingKey] = ExportFormatComboBox.SelectedIndex;
        settings[PresetSettingKey] = PresetComboBox.SelectedIndex;
        settings[FontSettingKey] = SelectedFontFamily();
        settings[BodyFontSizeSettingKey] = BodyFontSizeComboBox.SelectedIndex;
        settings[HeadingFontSizeSettingKey] = HeadingFontSizeComboBox.SelectedIndex;
        settings[LineSpacingSettingKey] = LineSpacingComboBox.SelectedIndex;
        settings[PageMarginSettingKey] = PageMarginComboBox.SelectedIndex;
        settings[QuoteStyleSettingKey] = QuoteStyleComboBox.SelectedIndex;
        settings[ListDensitySettingKey] = ListDensityComboBox.SelectedIndex;
    }

    private static void SaveLastExportFolder(string filePath)
    {
        var folderPath = ExportLocationMemory.ParentFolderFromFilePath(filePath);
        if (folderPath is not null)
        {
            ApplicationData.Current.LocalSettings.Values[LastExportFolderSettingKey] = folderPath;
        }
    }

    private static int ReadIndexSetting(
        IPropertySet settings,
        string key,
        int fallback,
        int itemCount)
    {
        if (!settings.TryGetValue(key, out var value) || value is not int index)
        {
            return fallback;
        }

        return index >= 0 && index < itemCount ? index : fallback;
    }

    private void ConfigureWindowChrome()
    {
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(TitleBarDragRegion);

        var windowHandle = WindowNative.GetWindowHandle(this);
        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(windowHandle);
        var appWindow = AppWindow.GetFromWindowId(windowId);
        appWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));

        appWindow.TitleBar.ButtonBackgroundColor = Microsoft.UI.Colors.Transparent;
        appWindow.TitleBar.ButtonInactiveBackgroundColor = Microsoft.UI.Colors.Transparent;
        appWindow.TitleBar.ButtonHoverBackgroundColor = Windows.UI.Color.FromArgb(24, 0, 0, 0);
        appWindow.TitleBar.ButtonPressedBackgroundColor = Windows.UI.Color.FromArgb(36, 0, 0, 0);
    }

    private bool AreExportSettingControlsReady()
    {
        return SettingsSummaryText is not null
            && ExportFormatComboBox is not null
            && ExportButtonText is not null
            && PresetComboBox is not null
            && FontSearchBox is not null
            && BodyFontSizeComboBox is not null
            && HeadingFontSizeComboBox is not null
            && LineSpacingComboBox is not null
            && PageMarginComboBox is not null
            && QuoteStyleComboBox is not null
            && ListDensityComboBox is not null;
    }

    private static double SelectedDouble(double[] values, int index, double fallback)
    {
        return index >= 0 && index < values.Length ? values[index] : fallback;
    }

    private void ConvertInput()
    {
        if (string.IsNullOrWhiteSpace(InputTextBox.Text))
        {
            PreviewPanel.Children.Clear();
            PreviewScrollViewer.Visibility = Visibility.Collapsed;
            PreviewPage.Visibility = Visibility.Collapsed;
            EmptyPreviewPanel.Visibility = Visibility.Visible;
            InputMetaText.Text = "等待粘贴内容";
            PreviewMetaText.Text = "暂无文档块";
            StatusText.Text = "就绪";
            currentResult = null;
            return;
        }

        currentResult = conversionService.Convert(InputTextBox.Text);
        RenderPreview(currentResult.Document);
        InputMetaText.Text = $"{InputTextBox.Text.Length:N0} 个字符";
        PreviewMetaText.Text = $"{currentResult.Document.Blocks.Count} 个文档块";
        StatusText.Text = $"{currentResult.Document.Blocks.Count} 个文档块已就绪";
    }

    private void RenderPreview(DocumentModel document)
    {
        PreviewPanel.Children.Clear();
        var hasBlocks = document.Blocks.Count > 0;
        EmptyPreviewPanel.Visibility = hasBlocks ? Visibility.Collapsed : Visibility.Visible;
        PreviewScrollViewer.Visibility = hasBlocks ? Visibility.Visible : Visibility.Collapsed;
        PreviewPage.Visibility = hasBlocks ? Visibility.Visible : Visibility.Collapsed;

        var options = CreateExportOptions();
        var metrics = PreviewLayoutMetrics.FromExportOptions(options);
        ApplyPreviewPageLayout(metrics);

        foreach (var block in document.Blocks)
        {
            PreviewPanel.Children.Add(CreatePreviewElement(block, metrics, options.FontFamily));
        }
    }

    private void RefreshPreview()
    {
        if (currentResult is not null)
        {
            RenderPreview(currentResult.Document);
        }
    }

    private void ApplyPreviewPageLayout(PreviewLayoutMetrics metrics)
    {
        PreviewPage.Padding = new Thickness(metrics.PagePadding);
    }

    private static FrameworkElement CreatePreviewElement(
        DocumentBlock block,
        PreviewLayoutMetrics metrics,
        string fontFamily)
    {
        return block switch
        {
            HeadingBlock heading => CreateHeadingPreview(heading, metrics, fontFamily),
            ParagraphBlock paragraph => CreateParagraphPreview(paragraph, metrics, fontFamily),
            BlockQuoteBlock quote => CreateQuotePreview(quote, metrics, fontFamily),
            CodeBlock code => CreateCodePreview(code, metrics),
            DividerBlock => new Border
            {
                Height = 1,
                Margin = new Thickness(0, 14, 0, 14),
                Background = new SolidColorBrush(Microsoft.UI.Colors.Gray),
                Opacity = 0.36
            },
            ListBlock list => CreateListPreview(list, metrics, fontFamily),
            _ => PreviewText(string.Empty, metrics.BodyFontSize, FontWeights.Normal, metrics, fontFamily)
        };
    }

    private static FrameworkElement CreateHeadingPreview(HeadingBlock heading, PreviewLayoutMetrics metrics, string fontFamily)
    {
        var fontSize = heading.Level switch
        {
            1 => metrics.Heading1FontSize,
            2 => metrics.Heading2FontSize,
            _ => metrics.Heading3FontSize
        };
        var text = PreviewText(heading.Text, fontSize, FontWeights.SemiBold, metrics, fontFamily, heading.Inlines);
        text.Margin = heading.Level == 1 ? new Thickness(0, 0, 0, 14) : new Thickness(0, 12, 0, 8);
        return text;
    }

    private static FrameworkElement CreateParagraphPreview(ParagraphBlock paragraph, PreviewLayoutMetrics metrics, string fontFamily)
    {
        var text = PreviewText(paragraph.Text, metrics.BodyFontSize, FontWeights.Normal, metrics, fontFamily, paragraph.Inlines);
        text.Margin = new Thickness(0, 0, 0, 10);
        return text;
    }

    private static TextBlock PreviewText(
        string text,
        double fontSize,
        FontWeight fontWeight,
        PreviewLayoutMetrics metrics,
        string fontFamily,
        IReadOnlyList<DocumentInline>? inlines = null,
        double opacity = 1)
    {
        var textBlock = new TextBlock
        {
            FontSize = fontSize,
            FontWeight = fontWeight,
            FontFamily = new FontFamily(fontFamily),
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.Black),
            LineHeight = Math.Max(metrics.LineHeight, fontSize * 1.15),
            LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
            Opacity = opacity,
            TextWrapping = TextWrapping.Wrap
        };
        AppendPreviewInlines(textBlock, text, inlines);
        return textBlock;
    }

    private static FrameworkElement CreateQuotePreview(BlockQuoteBlock quote, PreviewLayoutMetrics metrics, string fontFamily)
    {
        var quoteText = PreviewText(
            quote.Text,
            metrics.BodyFontSize,
            FontWeights.Normal,
            metrics,
            fontFamily,
            quote.Inlines,
            opacity: 0.78);
        Grid.SetColumn(quoteText, 1);

        var grid = new Grid
        {
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
            },
            ColumnSpacing = 10,
            Children =
            {
                new Border
                {
                    Width = 3,
                    CornerRadius = new CornerRadius(2),
                    Background = (Brush)Application.Current.Resources["SystemControlForegroundAccentBrush"]
                },
                quoteText
            }
        };

        return new Border
        {
            Margin = new Thickness(0, 4, 0, 14),
            Padding = metrics.UseGrayQuoteBlock ? new Thickness(12, 10, 12, 10) : new Thickness(0),
            Background = metrics.UseGrayQuoteBlock ? new SolidColorBrush(Windows.UI.Color.FromArgb(255, 246, 247, 249)) : null,
            CornerRadius = new CornerRadius(4),
            Child = grid
        };
    }

    private static FrameworkElement CreateCodePreview(CodeBlock code, PreviewLayoutMetrics metrics)
    {
        var panel = new StackPanel { Spacing = 8 };

        if (!string.IsNullOrWhiteSpace(code.Language))
        {
            panel.Children.Add(new TextBlock
            {
                Text = code.Language,
                FontSize = 12,
                Foreground = new SolidColorBrush(Microsoft.UI.Colors.DimGray),
                Opacity = 0.62
            });
        }

        panel.Children.Add(new TextBlock
        {
            Text = code.Code,
            FontFamily = new FontFamily("Cascadia Mono, Consolas"),
            FontSize = 13,
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.Black),
            LineHeight = Math.Max(metrics.LineHeight, 16),
            LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
            TextWrapping = TextWrapping.NoWrap
        });

        return new Border
        {
            Margin = new Thickness(0, 4, 0, 14),
            Padding = new Thickness(12),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 246, 247, 249)),
            CornerRadius = new CornerRadius(4),
            Child = panel
        };
    }

    private static FrameworkElement CreateListPreview(ListBlock list, PreviewLayoutMetrics metrics, string fontFamily)
    {
        var panel = new StackPanel
        {
            Spacing = metrics.ListItemSpacing,
            Margin = new Thickness(0, 0, 0, 10)
        };

        for (var index = 0; index < list.Items.Count; index++)
        {
            var row = new Grid { ColumnSpacing = 10 };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var marker = new TextBlock
            {
                Text = list.IsOrdered ? $"{index + 1}." : "-",
                FontFamily = new FontFamily(fontFamily),
                FontSize = metrics.BodyFontSize,
                Foreground = new SolidColorBrush(Microsoft.UI.Colors.Black),
                LineHeight = metrics.LineHeight,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
                Opacity = 0.68,
                MinWidth = 28
            };
            var text = PreviewText(
                list.Items[index],
                metrics.BodyFontSize,
                FontWeights.Normal,
                metrics,
                fontFamily,
                list.ItemInlines is not null && index < list.ItemInlines.Count ? list.ItemInlines[index] : null);
            Grid.SetColumn(text, 1);

            row.Children.Add(marker);
            row.Children.Add(text);
            panel.Children.Add(row);
        }

        return panel;
    }

    private static void AppendPreviewInlines(
        TextBlock textBlock,
        string fallbackText,
        IReadOnlyList<DocumentInline>? inlines)
    {
        if (inlines is null || inlines.Count == 0)
        {
            textBlock.Text = fallbackText;
            return;
        }

        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case TextInline text:
                    textBlock.Inlines.Add(new Run { Text = text.Text });
                    break;
                case BoldInline bold:
                    textBlock.Inlines.Add(new Microsoft.UI.Xaml.Documents.Bold
                    {
                        Inlines = { new Run { Text = bold.Text } }
                    });
                    break;
                case ItalicInline italic:
                    textBlock.Inlines.Add(new Italic
                    {
                        Inlines = { new Run { Text = italic.Text } }
                    });
                    break;
                case CodeInline code:
                    textBlock.Inlines.Add(new Run
                    {
                        Text = code.Text,
                        FontFamily = new FontFamily("Cascadia Mono, Consolas"),
                        Foreground = new SolidColorBrush(Microsoft.UI.Colors.DarkSlateGray)
                    });
                    break;
            }
        }
    }
}
