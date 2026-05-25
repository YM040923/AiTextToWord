using System.Text.Json;

namespace AiTextToWord.Core.Export;

public enum ExportFormat
{
    Word,
    Pdf
}

public enum ExportPreset
{
    StandardDocument,
    CompactNotes,
    FormalReport
}

public enum ExportPageMargin
{
    Narrow,
    Standard,
    Wide
}

public enum ExportQuoteStyle
{
    Indented,
    GrayBlock
}

public enum ExportListDensity
{
    Compact,
    Comfortable
}

public sealed record ExportSettingsSuggestion(
    ExportFormat Format,
    ExportPreset Preset,
    string FontFamily,
    double BodyFontSize,
    double HeadingFontSize,
    double LineSpacing,
    ExportPageMargin PageMargin,
    ExportQuoteStyle QuoteStyle,
    ExportListDensity ListDensity)
{
    public string DisplaySummary()
    {
        return $"{FormatText()} · {PresetText()} · {FontFamily} · {BodyFontSize:0.#}pt · {LineSpacing:0.##} 倍行距";
    }

    private string FormatText() => Format == ExportFormat.Pdf ? "PDF" : "Word";

    private string PresetText()
    {
        return Preset switch
        {
            ExportPreset.CompactNotes => "紧凑笔记",
            ExportPreset.FormalReport => "正式报告",
            _ => "标准文档"
        };
    }
}

public static class AiSettingsSuggestionEngine
{
    public static ExportSettingsSuggestion Suggest(string request, string fallbackFont)
    {
        var normalized = request.Trim().ToLowerInvariant();
        if (ContainsAny(normalized, "紧凑", "省纸", "笔记", "自己看", "简洁"))
        {
            return new ExportSettingsSuggestion(
                ExportFormat.Word,
                ExportPreset.CompactNotes,
                SafeFont(fallbackFont),
                10.5,
                16,
                1.15,
                ExportPageMargin.Narrow,
                ExportQuoteStyle.Indented,
                ExportListDensity.Compact);
        }

        if (ContainsAny(normalized, "正式", "报告", "老师", "老板", "提交", "宽松", "严肃"))
        {
            return new ExportSettingsSuggestion(
                ExportFormat.Word,
                ExportPreset.FormalReport,
                SafeFont(fallbackFont),
                12,
                22,
                1.5,
                ExportPageMargin.Wide,
                ExportQuoteStyle.GrayBlock,
                ExportListDensity.Comfortable);
        }

        return new ExportSettingsSuggestion(
            ContainsAny(normalized, "pdf") ? ExportFormat.Pdf : ExportFormat.Word,
            ExportPreset.StandardDocument,
            SafeFont(fallbackFont),
            11,
            20,
            1.3,
            ExportPageMargin.Standard,
            ExportQuoteStyle.Indented,
            ExportListDensity.Compact);
    }

    public static bool TryParseModelJson(
        string response,
        string fallbackFont,
        out ExportSettingsSuggestion suggestion)
    {
        suggestion = Suggest(string.Empty, fallbackFont);
        var json = ExtractJsonObject(response);
        if (json is null)
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            suggestion = new ExportSettingsSuggestion(
                Format: ParseFormat(ReadString(root, "format")),
                Preset: ParsePreset(ReadString(root, "preset")),
                FontFamily: SafeFont(ReadString(root, "fontFamily") ?? fallbackFont),
                BodyFontSize: ExportSettingRanges.ClampBodyFontSize(ReadDouble(root, "bodyFontSize", 11)),
                HeadingFontSize: ExportSettingRanges.ClampHeadingFontSize(ReadDouble(root, "headingFontSize", 20)),
                LineSpacing: ExportSettingRanges.ClampLineSpacing(ReadDouble(root, "lineSpacing", 1.3)),
                PageMargin: ParsePageMargin(ReadString(root, "pageMargin")),
                QuoteStyle: ParseQuoteStyle(ReadString(root, "quoteStyle")),
                ListDensity: ParseListDensity(ReadString(root, "listDensity")));
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string? ExtractJsonObject(string response)
    {
        var start = response.IndexOf('{');
        var end = response.LastIndexOf('}');
        if (start < 0 || end <= start)
        {
            return null;
        }

        return response[start..(end + 1)];
    }

    private static bool ContainsAny(string text, params string[] keywords)
    {
        return keywords.Any(text.Contains);
    }

    private static string SafeFont(string? fontFamily)
    {
        return string.IsNullOrWhiteSpace(fontFamily) ? "Microsoft YaHei" : fontFamily.Trim();
    }

    private static string? ReadString(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var element) && element.ValueKind == JsonValueKind.String
            ? element.GetString()
            : null;
    }

    private static double ReadDouble(JsonElement root, string propertyName, double fallback)
    {
        return root.TryGetProperty(propertyName, out var element) && element.TryGetDouble(out var value)
            ? value
            : fallback;
    }

    private static ExportFormat ParseFormat(string? value)
    {
        return ContainsAny(Normalize(value), "pdf") ? ExportFormat.Pdf : ExportFormat.Word;
    }

    private static ExportPreset ParsePreset(string? value)
    {
        var normalized = Normalize(value);
        if (ContainsAny(normalized, "compact", "note", "紧凑", "笔记"))
        {
            return ExportPreset.CompactNotes;
        }

        if (ContainsAny(normalized, "formal", "report", "正式", "报告"))
        {
            return ExportPreset.FormalReport;
        }

        return ExportPreset.StandardDocument;
    }

    private static ExportPageMargin ParsePageMargin(string? value)
    {
        var normalized = Normalize(value);
        if (ContainsAny(normalized, "narrow", "窄"))
        {
            return ExportPageMargin.Narrow;
        }

        if (ContainsAny(normalized, "wide", "宽"))
        {
            return ExportPageMargin.Wide;
        }

        return ExportPageMargin.Standard;
    }

    private static ExportQuoteStyle ParseQuoteStyle(string? value)
    {
        return ContainsAny(Normalize(value), "gray", "block", "灰")
            ? ExportQuoteStyle.GrayBlock
            : ExportQuoteStyle.Indented;
    }

    private static ExportListDensity ParseListDensity(string? value)
    {
        return ContainsAny(Normalize(value), "comfortable", "loose", "舒展", "宽松")
            ? ExportListDensity.Comfortable
            : ExportListDensity.Compact;
    }

    private static string Normalize(string? value)
    {
        return value?.Trim().ToLowerInvariant() ?? string.Empty;
    }
}
