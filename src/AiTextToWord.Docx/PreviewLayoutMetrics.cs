namespace AiTextToWord.Docx;

public sealed record PreviewLayoutMetrics(
    double BodyFontSize,
    double Heading1FontSize,
    double Heading2FontSize,
    double Heading3FontSize,
    double LineHeight,
    double PagePadding,
    double ListItemSpacing,
    bool UseGrayQuoteBlock)
{
    public static PreviewLayoutMetrics FromExportOptions(DocxExportOptions options)
    {
        return new PreviewLayoutMetrics(
            BodyFontSize: options.BodyFontSize,
            Heading1FontSize: options.HeadingFontSize,
            Heading2FontSize: Math.Max(options.HeadingFontSize - 2, options.BodyFontSize + 2),
            Heading3FontSize: Math.Max(options.HeadingFontSize - 4, options.BodyFontSize + 1),
            LineHeight: options.BodyFontSize * options.LineSpacing,
            PagePadding: ComputePagePadding(options),
            ListItemSpacing: options.ListDensity == DocxListDensity.Comfortable ? 6 : 3,
            UseGrayQuoteBlock: options.QuoteStyle == DocxQuoteStyle.GrayBlock);
    }

    private static double ComputePagePadding(DocxExportOptions options)
    {
        if (options.PageMargin == DocxPageMargin.Custom && options.CustomPageMarginCentimeters is { } centimeters)
        {
            return Math.Clamp(centimeters, 0.5, 6) * 22;
        }

        return options.PageMargin switch
        {
            DocxPageMargin.Narrow => 40,
            DocxPageMargin.Wide => 72,
            _ => 56
        };
    }
}
