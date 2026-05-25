using AiTextToWord.Core.Model;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Fonts;

namespace AiTextToWord.Docx;

public sealed class PdfExporter
{
    public void Export(DocumentModel document, Stream output, DocxExportOptions options)
    {
        GlobalFontSettings.FontResolver ??= new WindowsPdfFontResolver();

        var pdfDocument = CreateDocument(document, options);
        var renderer = new PdfDocumentRenderer
        {
            Document = pdfDocument
        };

        renderer.RenderDocument();
        renderer.Save(output, closeStream: false);
    }

    private static Document CreateDocument(DocumentModel document, DocxExportOptions options)
    {
        var pdfDocument = new Document();
        pdfDocument.Info.Title = options.Title;
        DefineStyles(pdfDocument, options);

        var section = pdfDocument.AddSection();
        section.PageSetup.PageFormat = PageFormat.A4;
        ApplyMargins(section.PageSetup, options.PageMargin);

        for (var index = 0; index < document.Blocks.Count; index++)
        {
            var block = document.Blocks[index];
            if (block is BlockQuoteBlock)
            {
                var quotes = new List<BlockQuoteBlock>();
                while (index < document.Blocks.Count && document.Blocks[index] is BlockQuoteBlock quote)
                {
                    quotes.Add(quote);
                    index++;
                }

                index--;
                AddQuote(section, quotes, options);
                continue;
            }

            AddBlock(section, block, options);
        }

        return pdfDocument;
    }

    private static void DefineStyles(Document document, DocxExportOptions options)
    {
        var normal = document.Styles["Normal"]!;
        normal.Font.Name = options.FontFamily;
        normal.Font.Size = Unit.FromPoint(options.BodyFontSize);
        normal.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
        normal.ParagraphFormat.LineSpacing = Unit.FromPoint(options.BodyFontSize * options.LineSpacing);

        AddParagraphStyle(document, "BodyText", "Normal", options.BodyFontSize, options.FontFamily, options.LineSpacing, bold: false);
        AddParagraphStyle(document, "CodeBlock", "Normal", Math.Max(options.BodyFontSize - 0.5, 10), "Consolas", options.LineSpacing, bold: false);
        AddParagraphStyle(document, "QuoteBlock", "Normal", options.BodyFontSize, options.FontFamily, options.LineSpacing, bold: false);
        AddParagraphStyle(document, "Heading1", "Normal", options.HeadingFontSize, options.FontFamily, options.LineSpacing, bold: true);
        AddParagraphStyle(document, "Heading2", "Normal", Math.Max(options.HeadingFontSize - 2, options.BodyFontSize + 2), options.FontFamily, options.LineSpacing, bold: true);
        AddParagraphStyle(document, "Heading3", "Normal", Math.Max(options.HeadingFontSize - 4, options.BodyFontSize + 1), options.FontFamily, options.LineSpacing, bold: true);

        document.Styles["BodyText"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(8);
        document.Styles["BodyText"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);
        document.Styles["CodeBlock"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(8);
        document.Styles["CodeBlock"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);
        document.Styles["CodeBlock"]!.ParagraphFormat.LeftIndent = Unit.FromPoint(18);
        document.Styles["CodeBlock"]!.ParagraphFormat.RightIndent = Unit.FromPoint(18);
        document.Styles["QuoteBlock"]!.ParagraphFormat.LeftIndent = Unit.FromPoint(18);
        document.Styles["QuoteBlock"]!.ParagraphFormat.RightIndent = Unit.FromPoint(18);
        document.Styles["QuoteBlock"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(8);
        document.Styles["QuoteBlock"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);
        document.Styles["Heading1"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(18);
        document.Styles["Heading1"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(8);
        document.Styles["Heading2"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(12);
        document.Styles["Heading2"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(6);
        document.Styles["Heading3"]!.ParagraphFormat.SpaceBefore = Unit.FromPoint(10);
        document.Styles["Heading3"]!.ParagraphFormat.SpaceAfter = Unit.FromPoint(6);
    }

    private static void AddParagraphStyle(
        Document document,
        string name,
        string baseStyle,
        double fontSize,
        string fontFamily,
        double lineSpacing,
        bool bold)
    {
        var style = document.Styles.AddStyle(name, baseStyle);
        style.Font.Name = fontFamily;
        style.Font.Size = Unit.FromPoint(fontSize);
        style.Font.Bold = bold;
        style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
        style.ParagraphFormat.LineSpacing = Unit.FromPoint(fontSize * lineSpacing);
    }

    private static void ApplyMargins(PageSetup pageSetup, DocxPageMargin pageMargin)
    {
        var margin = pageMargin switch
        {
            DocxPageMargin.Narrow => Unit.FromCentimeter(1.27),
            DocxPageMargin.Wide => Unit.FromCentimeter(3.18),
            _ => Unit.FromCentimeter(2.54)
        };

        pageSetup.TopMargin = margin;
        pageSetup.RightMargin = margin;
        pageSetup.BottomMargin = margin;
        pageSetup.LeftMargin = margin;
    }

    private static void AddBlock(Section section, DocumentBlock block, DocxExportOptions options)
    {
        switch (block)
        {
            case HeadingBlock heading:
                AddInlineParagraph(section, heading.Text, heading.Inlines, $"Heading{Math.Clamp(heading.Level, 1, 3)}");
                break;
            case ParagraphBlock paragraph:
                AddInlineParagraph(section, paragraph.Text, paragraph.Inlines, "BodyText");
                break;
            case ListBlock list:
                AddList(section, list, options);
                break;
            case CodeBlock code:
                AddCode(section, code);
                break;
            case DividerBlock:
                AddDivider(section);
                break;
        }
    }

    private static void AddList(Section section, ListBlock list, DocxExportOptions options)
    {
        var space = options.ListDensity == DocxListDensity.Comfortable ? 5 : 2;
        for (var index = 0; index < list.Items.Count; index++)
        {
            var paragraph = section.AddParagraph();
            paragraph.Style = "BodyText";
            paragraph.Format.SpaceBefore = Unit.FromPoint(space);
            paragraph.Format.SpaceAfter = Unit.FromPoint(space);
            paragraph.Format.ListInfo.ListType = list.IsOrdered ? ListType.NumberList1 : ListType.BulletList1;
            AddInlines(paragraph, list.Items[index], list.ItemInlines is not null && index < list.ItemInlines.Count ? list.ItemInlines[index] : null);
        }
    }

    private static void AddQuote(Section section, IReadOnlyList<BlockQuoteBlock> quotes, DocxExportOptions options)
    {
        var paragraph = section.AddParagraph();
        paragraph.Style = "QuoteBlock";
        paragraph.Format.Font.Color = Colors.DimGray;
        paragraph.Format.Borders.Left.Width = Unit.FromPoint(2);
        paragraph.Format.Borders.Left.Color = Colors.SteelBlue;
        paragraph.Format.Borders.DistanceFromLeft = Unit.FromPoint(8);

        if (options.QuoteStyle == DocxQuoteStyle.GrayBlock)
        {
            paragraph.Format.Shading.Color = Colors.WhiteSmoke;
        }

        for (var index = 0; index < quotes.Count; index++)
        {
            if (index > 0)
            {
                paragraph.AddText(" ");
            }

            AddInlines(paragraph, quotes[index].Text, quotes[index].Inlines);
        }
    }

    private static void AddCode(Section section, CodeBlock code)
    {
        if (!string.IsNullOrWhiteSpace(code.Language))
        {
            var language = section.AddParagraph(code.Language);
            language.Format.Font.Size = Unit.FromPoint(9);
            language.Format.Font.Color = Colors.DimGray;
            language.Format.SpaceBefore = Unit.FromPoint(8);
            language.Format.SpaceAfter = Unit.FromPoint(2);
        }

        var paragraph = section.AddParagraph();
        paragraph.Style = "CodeBlock";
        paragraph.Format.Shading.Color = Colors.WhiteSmoke;
        AddTextWithBreaks(paragraph, code.Code);
    }

    private static void AddDivider(Section section)
    {
        var paragraph = section.AddParagraph();
        paragraph.Format.Borders.Bottom.Width = Unit.FromPoint(0.5);
        paragraph.Format.Borders.Bottom.Color = Colors.LightGray;
        paragraph.Format.SpaceBefore = Unit.FromPoint(8);
        paragraph.Format.SpaceAfter = Unit.FromPoint(8);
    }

    private static void AddInlineParagraph(
        Section section,
        string text,
        IReadOnlyList<DocumentInline>? inlines,
        string style)
    {
        var paragraph = section.AddParagraph();
        paragraph.Style = style;
        AddInlines(paragraph, text, inlines);
    }

    private static void AddInlines(Paragraph paragraph, string fallbackText, IReadOnlyList<DocumentInline>? inlines)
    {
        if (inlines is null || inlines.Count == 0)
        {
            AddTextWithBreaks(paragraph, fallbackText);
            return;
        }

        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case TextInline text:
                    AddTextWithBreaks(paragraph, text.Text);
                    break;
                case BoldInline bold:
                    AddFormattedText(paragraph, bold.Text, isBold: true);
                    break;
                case ItalicInline italic:
                    AddFormattedText(paragraph, italic.Text, isItalic: true);
                    break;
                case CodeInline code:
                    var formatted = paragraph.AddFormattedText(code.Text);
                    formatted.Font.Name = "Consolas";
                    formatted.Font.Color = Colors.DarkSlateGray;
                    break;
            }
        }
    }

    private static void AddFormattedText(Paragraph paragraph, string text, bool isBold = false, bool isItalic = false)
    {
        var formatted = paragraph.AddFormattedText(text);
        formatted.Bold = isBold;
        formatted.Italic = isItalic;
    }

    private static void AddTextWithBreaks(Paragraph paragraph, string text)
    {
        var segments = text.Replace("\r\n", "\n").Split('\n');
        for (var index = 0; index < segments.Length; index++)
        {
            paragraph.AddText(segments[index]);
            if (index < segments.Length - 1)
            {
                paragraph.AddLineBreak();
            }
        }
    }
}
