using AiTextToWord.Core.Conversion;
using AiTextToWord.Core.Export;
using AiTextToWord.Core.Fonts;
using AiTextToWord.Core.Model;
using AiTextToWord.Docx;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.UI.Windowing;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation.Collections;
using Windows.Security.Credentials;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.System;
using Windows.UI.Text;
using WinRT.Interop;

namespace AiTextToWord.App;

public sealed partial class MainWindow : Window
{
    private readonly TextConversionService conversionService = TextConversionService.CreateDefault();
    private ConversionResult? currentResult;
    private ExportedFileActionState? lastExportedFile;
    private ExportSettingsSuggestion? pendingAiSuggestion;
    private IReadOnlyList<string> installedFonts = [];
    private bool isLoadingExportSettings;
    private bool isLoadingAiSettings;
    private static readonly HttpClient AiHttpClient = new();
    private static readonly double[] LegacyBodyFontSizes = [10.5, 11, 12, 14];
    private static readonly double[] LegacyHeadingFontSizes = [16, 18, 20, 22];
    private static readonly double[] LegacyLineSpacings = [1.15, 1.3, 1.5, 2.0];
    private const string ExportFormatSettingKey = "Export.Format";
    private const string PresetSettingKey = "Export.Preset";
    private const string FontSettingKey = "Export.Font";
    private const string BodyFontSizeSettingKey = "Export.BodyFontSize";
    private const string HeadingFontSizeSettingKey = "Export.HeadingFontSize";
    private const string LineSpacingSettingKey = "Export.LineSpacing";
    private const string PageMarginSettingKey = "Export.PageMargin";
    private const string CustomPageMarginSettingKey = "Export.CustomPageMargin";
    private const string QuoteStyleSettingKey = "Export.QuoteStyle";
    private const string ListDensitySettingKey = "Export.ListDensity";
    private const string LastExportFolderSettingKey = "Export.LastFolder";
    private const string AiEndpointSettingKey = "AI.Endpoint";
    private const string AiModelSettingKey = "AI.Model";
    private const string AiCredentialResource = "AiTextToWord.AiSettingsAssistant";
    private const string AiCredentialUserName = "ApiKey";
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
        LoadAiSettings();
        isLoadingExportSettings = false;
        SystemBackdrop = new MicaBackdrop();
        UpdateSettingsSummary();
        UpdateExportButton();
        ShowWorkbenchView();
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
        lastExportedFile = null;
        UpdateExportActionButtons();
    }

    private void Sample_Click(object sender, RoutedEventArgs e)
    {
        InputTextBox.Text = SampleText;
        StatusText.Text = "已载入示例文本。";
    }

    private void ShowWorkbench_Click(object sender, RoutedEventArgs e)
    {
        ShowWorkbenchView();
    }

    private void ShowSettings_Click(object sender, RoutedEventArgs e)
    {
        ShowSettingsView();
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
        lastExportedFile = ExportedFileActionState.FromFilePath(file.Path);
        UpdateExportActionButtons();
        StatusText.Text = $"{SelectedExportFormatName()} 文档已导出：{lastExportedFile?.FileName ?? file.Name}";
    }

    private async void OpenExport_Click(object sender, RoutedEventArgs e)
    {
        if (lastExportedFile is null || !File.Exists(lastExportedFile.FilePath))
        {
            StatusText.Text = "导出的文件不存在，可能已被移动或删除。";
            lastExportedFile = null;
            UpdateExportActionButtons();
            return;
        }

        var file = await StorageFile.GetFileFromPathAsync(lastExportedFile.FilePath);
        var opened = await Launcher.LaunchFileAsync(file);
        StatusText.Text = opened ? $"已打开：{lastExportedFile.FileName}" : "无法打开导出的文件。";
    }

    private async void OpenExportFolder_Click(object sender, RoutedEventArgs e)
    {
        if (lastExportedFile is null || !Directory.Exists(lastExportedFile.FolderPath))
        {
            StatusText.Text = "导出文件夹不存在，可能已被移动或删除。";
            lastExportedFile = null;
            UpdateExportActionButtons();
            return;
        }

        var folder = await StorageFolder.GetFolderFromPathAsync(lastExportedFile.FolderPath);
        var options = new FolderLauncherOptions();
        if (File.Exists(lastExportedFile.FilePath))
        {
            options.ItemsToSelect.Add(await StorageFile.GetFileFromPathAsync(lastExportedFile.FilePath));
        }

        var opened = await Launcher.LaunchFolderAsync(folder, options);
        StatusText.Text = opened ? $"已打开文件夹：{lastExportedFile.FileName}" : "无法打开导出文件夹。";
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
                BodyFontSizeNumberBox.Value = 10.5;
                HeadingFontSizeNumberBox.Value = 16;
                LineSpacingNumberBox.Value = 1.15;
                PageMarginComboBox.SelectedIndex = 0;
                QuoteStyleComboBox.SelectedIndex = 0;
                ListDensityComboBox.SelectedIndex = 0;
                break;
            case 2:
                SelectPreferredFont();
                BodyFontSizeNumberBox.Value = 12;
                HeadingFontSizeNumberBox.Value = 22;
                LineSpacingNumberBox.Value = 1.5;
                PageMarginComboBox.SelectedIndex = 2;
                QuoteStyleComboBox.SelectedIndex = 1;
                ListDensityComboBox.SelectedIndex = 1;
                break;
            default:
                SelectPreferredFont();
                BodyFontSizeNumberBox.Value = 11;
                HeadingFontSizeNumberBox.Value = 20;
                LineSpacingNumberBox.Value = 1.3;
                PageMarginComboBox.SelectedIndex = 1;
                QuoteStyleComboBox.SelectedIndex = 0;
                ListDensityComboBox.SelectedIndex = 0;
                break;
        }

        UpdateSettingsSummary();
        UpdateExportButton();
        UpdateCustomMarginVisibility();
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
        UpdateCustomMarginVisibility();
        RefreshPreview();
        SaveExportSettings();
    }

    private void ExportNumberSettings_Changed(NumberBox sender, NumberBoxValueChangedEventArgs args)
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
            BodyFontSize = SelectedBodyFontSize(),
            HeadingFontSize = SelectedHeadingFontSize(),
            LineSpacing = SelectedLineSpacing(),
            PageMargin = PageMarginComboBox.SelectedIndex switch
            {
                0 => DocxPageMargin.Narrow,
                2 => DocxPageMargin.Wide,
                3 => DocxPageMargin.Custom,
                _ => DocxPageMargin.Standard
            },
            CustomPageMarginCentimeters = SelectedCustomPageMargin(),
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
        var bodySize = SelectedBodyFontSize();
        var lineSpacing = SelectedLineSpacing();
        SettingsSummaryText.Text = $"{SelectedExportFormatName()} · {preset} · {font} · {bodySize:0.#}pt · {lineSpacing:0.##} 倍行距";
    }

    private void UpdateExportButton()
    {
        if (ExportButtonText is not null)
        {
            ExportButtonText.Text = $"导出 {SelectedExportFormatName()}";
        }
    }

    private void UpdateExportActionButtons()
    {
        if (OpenExportButton is null || OpenExportFolderButton is null)
        {
            return;
        }

        var visibility = lastExportedFile is null ? Visibility.Collapsed : Visibility.Visible;
        OpenExportButton.Visibility = visibility;
        OpenExportFolderButton.Visibility = visibility;
    }

    private void ShowWorkbenchView()
    {
        WorkbenchHeader.Visibility = Visibility.Visible;
        WorkbenchBody.Visibility = Visibility.Visible;
        WorkbenchFooter.Visibility = Visibility.Visible;
        SettingsView.Visibility = Visibility.Collapsed;
        WorkbenchNavButton.Opacity = 1;
        SettingsNavButton.Opacity = 0.72;
        StatusText.Text = "工作台";
    }

    private void ShowSettingsView()
    {
        WorkbenchHeader.Visibility = Visibility.Collapsed;
        WorkbenchBody.Visibility = Visibility.Collapsed;
        WorkbenchFooter.Visibility = Visibility.Collapsed;
        SettingsView.Visibility = Visibility.Visible;
        WorkbenchNavButton.Opacity = 0.72;
        SettingsNavButton.Opacity = 1;
        StatusText.Text = "设置";
    }

    private async void SuggestSettings_Click(object sender, RoutedEventArgs e)
    {
        var request = AiInstructionTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(request))
        {
            AiAssistantStatusText.Text = "先输入你想要的文档效果。";
            return;
        }

        try
        {
            AiSuggestButton.IsEnabled = false;
            AiAssistantStatusText.Text = HasAiServiceConfiguration()
                ? "正在请求 AI 模型生成设置建议..."
                : "未配置模型，正在使用本地规则生成建议。";

            pendingAiSuggestion = HasAiServiceConfiguration()
                ? await RequestAiSettingsSuggestionAsync(request)
                : AiSettingsSuggestionEngine.Suggest(request, SelectedFontFamily());

            ShowAiSuggestion(pendingAiSuggestion);
            AiAssistantStatusText.Text = HasAiServiceConfiguration()
                ? "AI 已生成建议，确认后才会应用。"
                : "已生成本地建议，确认后才会应用。";
        }
        catch (Exception ex) when (ex is HttpRequestException or JsonException or InvalidOperationException)
        {
            pendingAiSuggestion = AiSettingsSuggestionEngine.Suggest(request, SelectedFontFamily());
            ShowAiSuggestion(pendingAiSuggestion);
            AiAssistantStatusText.Text = $"模型请求失败，已使用本地建议：{ex.Message}";
        }
        finally
        {
            AiSuggestButton.IsEnabled = true;
        }
    }

    private void ApplyAiSuggestion_Click(object sender, RoutedEventArgs e)
    {
        if (pendingAiSuggestion is null)
        {
            return;
        }

        ApplySettingsSuggestion(pendingAiSuggestion);
        AiSuggestionPanel.Visibility = Visibility.Collapsed;
        pendingAiSuggestion = null;
        UpdateSettingsSummary();
        UpdateExportButton();
        RefreshPreview();
        SaveExportSettings();
        AiAssistantStatusText.Text = "已应用建议到导出设置。";
    }

    private void DismissAiSuggestion_Click(object sender, RoutedEventArgs e)
    {
        pendingAiSuggestion = null;
        AiSuggestionPanel.Visibility = Visibility.Collapsed;
        AiAssistantStatusText.Text = "已取消建议。";
    }

    private void ShowAiSuggestion(ExportSettingsSuggestion suggestion)
    {
        AiSuggestionSummaryText.Text = suggestion.DisplaySummary();
        AiSuggestionPanel.Visibility = Visibility.Visible;
    }

    private void ApplySettingsSuggestion(ExportSettingsSuggestion suggestion)
    {
        ExportFormatComboBox.SelectedIndex = suggestion.Format == ExportFormat.Pdf ? 1 : 0;
        PresetComboBox.SelectedIndex = suggestion.Preset switch
        {
            ExportPreset.CompactNotes => 1,
            ExportPreset.FormalReport => 2,
            _ => 0
        };
        FontSearchBox.Text = FindInstalledFont(suggestion.FontFamily) ?? suggestion.FontFamily;
        BodyFontSizeNumberBox.Value = suggestion.BodyFontSize;
        HeadingFontSizeNumberBox.Value = suggestion.HeadingFontSize;
        LineSpacingNumberBox.Value = suggestion.LineSpacing;
        PageMarginComboBox.SelectedIndex = suggestion.PageMargin switch
        {
            ExportPageMargin.Narrow => 0,
            ExportPageMargin.Wide => 2,
            _ => 1
        };
        QuoteStyleComboBox.SelectedIndex = suggestion.QuoteStyle == ExportQuoteStyle.GrayBlock ? 1 : 0;
        ListDensityComboBox.SelectedIndex = suggestion.ListDensity == ExportListDensity.Comfortable ? 1 : 0;
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
        BodyFontSizeNumberBox.Value = ReadNumberSetting(
            settings,
            BodyFontSizeSettingKey,
            11,
            ExportSettingRanges.MinimumBodyFontSize,
            ExportSettingRanges.MaximumBodyFontSize,
            LegacyBodyFontSizes);
        HeadingFontSizeNumberBox.Value = ReadNumberSetting(
            settings,
            HeadingFontSizeSettingKey,
            20,
            ExportSettingRanges.MinimumHeadingFontSize,
            ExportSettingRanges.MaximumHeadingFontSize,
            LegacyHeadingFontSizes);
        LineSpacingNumberBox.Value = ReadNumberSetting(
            settings,
            LineSpacingSettingKey,
            1.3,
            ExportSettingRanges.MinimumLineSpacing,
            ExportSettingRanges.MaximumLineSpacing,
            LegacyLineSpacings);
        PageMarginComboBox.SelectedIndex = ReadIndexSetting(settings, PageMarginSettingKey, 1, PageMarginComboBox.Items.Count);
        CustomMarginNumberBox.Value = ReadNumberSetting(settings, CustomPageMarginSettingKey, 2.54, 0.5, 6, []);
        UpdateCustomMarginVisibility();
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
        settings[BodyFontSizeSettingKey] = SelectedBodyFontSize();
        settings[HeadingFontSizeSettingKey] = SelectedHeadingFontSize();
        settings[LineSpacingSettingKey] = SelectedLineSpacing();
        settings[PageMarginSettingKey] = PageMarginComboBox.SelectedIndex;
        settings[CustomPageMarginSettingKey] = SelectedCustomPageMargin();
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

    private void LoadAiSettings()
    {
        isLoadingAiSettings = true;
        var settings = ApplicationData.Current.LocalSettings.Values;
        AiEndpointTextBox.Text = settings.TryGetValue(AiEndpointSettingKey, out var endpoint) && endpoint is string endpointText
            ? endpointText
            : string.Empty;
        AiModelTextBox.Text = settings.TryGetValue(AiModelSettingKey, out var model) && model is string modelText
            ? modelText
            : string.Empty;
        AiApiKeyPasswordBox.Password = LoadAiApiKey();
        isLoadingAiSettings = false;
    }

    private void AiTextSettings_Changed(object sender, TextChangedEventArgs e)
    {
        SaveAiSettings();
    }

    private void AiPasswordSettings_Changed(object sender, RoutedEventArgs e)
    {
        SaveAiSettings();
    }

    private void SaveAiSettings()
    {
        if (isLoadingAiSettings)
        {
            return;
        }

        var settings = ApplicationData.Current.LocalSettings.Values;
        settings[AiEndpointSettingKey] = AiEndpointTextBox.Text.Trim();
        settings[AiModelSettingKey] = AiModelTextBox.Text.Trim();
        SaveAiApiKey(AiApiKeyPasswordBox.Password);
    }

    private bool HasAiServiceConfiguration()
    {
        return !string.IsNullOrWhiteSpace(AiEndpointTextBox.Text)
            && !string.IsNullOrWhiteSpace(AiModelTextBox.Text)
            && !string.IsNullOrWhiteSpace(AiApiKeyPasswordBox.Password);
    }

    private async Task<ExportSettingsSuggestion> RequestAiSettingsSuggestionAsync(string request)
    {
        var endpoint = AiEndpointTextBox.Text.Trim();
        var model = AiModelTextBox.Text.Trim();
        var apiKey = AiApiKeyPasswordBox.Password;
        if (string.IsNullOrWhiteSpace(endpoint) || string.IsNullOrWhiteSpace(model) || string.IsNullOrWhiteSpace(apiKey))
        {
            return AiSettingsSuggestionEngine.Suggest(request, SelectedFontFamily());
        }

        using var message = new HttpRequestMessage(HttpMethod.Post, endpoint);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        var payload = new
        {
            model,
            temperature = 0.1,
            response_format = new { type = "json_object" },
            messages = new object[]
            {
                new
                {
                    role = "system",
                    content = """
                        你是一个桌面软件的设置助手。只返回 JSON 对象，不要解释。
                        允许字段：
                        format: word|pdf
                        preset: standard|compact|formal
                        fontFamily: 字体名称
                        bodyFontSize: 8 到 72
                        headingFontSize: 10 到 96
                        lineSpacing: 1 到 3
                        pageMargin: narrow|standard|wide
                        quoteStyle: indented|gray
                        listDensity: compact|comfortable
                        """
                },
                new
                {
                    role = "user",
                    content = request
                }
            }
        };
        message.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var response = await AiHttpClient.SendAsync(message);
        var responseText = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException($"HTTP {(int)response.StatusCode}: {response.ReasonPhrase}");
        }

        var modelContent = ExtractChatCompletionContent(responseText) ?? responseText;
        if (!AiSettingsSuggestionEngine.TryParseModelJson(modelContent, SelectedFontFamily(), out var suggestion))
        {
            throw new JsonException("模型没有返回可识别的设置 JSON。");
        }

        return suggestion;
    }

    private static string? ExtractChatCompletionContent(string responseText)
    {
        using var document = JsonDocument.Parse(responseText);
        if (!document.RootElement.TryGetProperty("choices", out var choices)
            || choices.ValueKind != JsonValueKind.Array
            || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        if (first.TryGetProperty("message", out var message)
            && message.TryGetProperty("content", out var content)
            && content.ValueKind == JsonValueKind.String)
        {
            return content.GetString();
        }

        return null;
    }

    private static string LoadAiApiKey()
    {
        try
        {
            var vault = new PasswordVault();
            var credential = vault.Retrieve(AiCredentialResource, AiCredentialUserName);
            credential.RetrievePassword();
            return credential.Password;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void SaveAiApiKey(string apiKey)
    {
        try
        {
            var vault = new PasswordVault();
            try
            {
                foreach (var credential in vault.FindAllByResource(AiCredentialResource))
                {
                    vault.Remove(credential);
                }
            }
            catch
            {
                // No saved credential yet.
            }

            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                vault.Add(new PasswordCredential(AiCredentialResource, AiCredentialUserName, apiKey));
            }
        }
        catch
        {
            // Credential Locker can be unavailable in some unpackaged/debug contexts.
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

    private static double ReadNumberSetting(
        IPropertySet settings,
        string key,
        double fallback,
        double minimum,
        double maximum,
        IReadOnlyList<double> legacyValues)
    {
        return settings.TryGetValue(key, out var value)
            ? ExportSettingRanges.NumberFromSetting(value, fallback, minimum, maximum, legacyValues)
            : fallback;
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
            && BodyFontSizeNumberBox is not null
            && HeadingFontSizeNumberBox is not null
            && LineSpacingNumberBox is not null
            && PageMarginComboBox is not null
            && CustomMarginPanel is not null
            && CustomMarginNumberBox is not null
            && QuoteStyleComboBox is not null
            && ListDensityComboBox is not null;
    }

    private double SelectedBodyFontSize()
    {
        return ExportSettingRanges.ClampBodyFontSize(BodyFontSizeNumberBox.Value);
    }

    private double SelectedHeadingFontSize()
    {
        return ExportSettingRanges.ClampHeadingFontSize(HeadingFontSizeNumberBox.Value);
    }

    private double SelectedLineSpacing()
    {
        return ExportSettingRanges.ClampLineSpacing(LineSpacingNumberBox.Value);
    }

    private double SelectedCustomPageMargin()
    {
        var value = CustomMarginNumberBox.Value;
        return double.IsNaN(value) || double.IsInfinity(value) ? 2.54 : Math.Clamp(value, 0.5, 6);
    }

    private void UpdateCustomMarginVisibility()
    {
        if (CustomMarginPanel is not null)
        {
            CustomMarginPanel.Visibility = PageMarginComboBox.SelectedIndex == 3 ? Visibility.Visible : Visibility.Collapsed;
        }
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
            TableBlock table => CreateTablePreview(table, metrics, fontFamily),
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

    private static FrameworkElement CreateTablePreview(TableBlock table, PreviewLayoutMetrics metrics, string fontFamily)
    {
        var columnCount = Math.Max(
            table.Headers.Count,
            table.Rows.Select(row => row.Count).DefaultIfEmpty(0).Max());
        if (columnCount == 0)
        {
            return PreviewText(string.Empty, metrics.BodyFontSize, FontWeights.Normal, metrics, fontFamily);
        }

        var grid = new Grid
        {
            Margin = new Thickness(0, 4, 0, 14),
            BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 209, 213, 219)),
            BorderThickness = new Thickness(1)
        };
        for (var index = 0; index < columnCount; index++)
        {
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        }

        var totalRowCount = table.Rows.Count + 1;
        AddPreviewTableRow(grid, table.Headers, columnCount, metrics, fontFamily, isHeader: true, totalRowCount);
        foreach (var row in table.Rows)
        {
            AddPreviewTableRow(grid, row, columnCount, metrics, fontFamily, isHeader: false, totalRowCount);
        }

        return grid;
    }

    private static void AddPreviewTableRow(
        Grid grid,
        IReadOnlyList<TableCell> cells,
        int columnCount,
        PreviewLayoutMetrics metrics,
        string fontFamily,
        bool isHeader,
        int totalRowCount)
    {
        var rowIndex = grid.RowDefinitions.Count;
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            var cell = columnIndex < cells.Count ? cells[columnIndex] : new TableCell(string.Empty, []);
            var text = PreviewText(
                cell.Text,
                metrics.BodyFontSize,
                isHeader ? FontWeights.SemiBold : FontWeights.Normal,
                metrics,
                fontFamily,
                isHeader ? HeaderPreviewInlines(cell) : cell.Inlines);
            var border = new Border
            {
                Padding = new Thickness(8, 6, 8, 6),
                BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 209, 213, 219)),
                BorderThickness = new Thickness(
                    0,
                    0,
                    columnIndex == columnCount - 1 ? 0 : 1,
                    rowIndex == totalRowCount - 1 ? 0 : 1),
                Background = isHeader ? new SolidColorBrush(Windows.UI.Color.FromArgb(255, 246, 247, 249)) : null,
                Child = text
            };
            Grid.SetRow(border, rowIndex);
            Grid.SetColumn(border, columnIndex);
            grid.Children.Add(border);
        }
    }

    private static IReadOnlyList<DocumentInline> HeaderPreviewInlines(TableCell cell)
    {
        return cell.Inlines.Count == 0
            ? [new BoldInline(cell.Text)]
            : cell.Inlines.Select<DocumentInline, DocumentInline>(inline => inline switch
            {
                TextInline text => new BoldInline(text.Text),
                BoldInline bold => bold,
                ItalicInline italic => new BoldInline(italic.Text),
                CodeInline code => code,
                _ => inline
            }).ToList();
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
