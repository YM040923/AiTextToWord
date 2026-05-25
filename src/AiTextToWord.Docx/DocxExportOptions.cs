namespace AiTextToWord.Docx;

public sealed record DocxExportOptions(string Title)
{
    public DocxStylePreset Preset { get; init; } = DocxStylePreset.StandardDocument;

    public string FontFamily { get; init; } = "Microsoft YaHei";

    public double BodyFontSize { get; init; } = 11;

    public double HeadingFontSize { get; init; } = 14;

    public double LineSpacing { get; init; } = 1.3;

    public DocxPageMargin PageMargin { get; init; } = DocxPageMargin.Standard;

    public double? CustomPageMarginCentimeters { get; init; }

    public DocxQuoteStyle QuoteStyle { get; init; } = DocxQuoteStyle.Indented;

    public DocxListDensity ListDensity { get; init; } = DocxListDensity.Compact;
}

public enum DocxStylePreset
{
    StandardDocument,
    CompactNotes,
    FormalReport
}

public enum DocxPageMargin
{
    Narrow,
    Standard,
    Wide,
    Custom
}

public enum DocxQuoteStyle
{
    Indented,
    GrayBlock
}

public enum DocxListDensity
{
    Compact,
    Comfortable
}
