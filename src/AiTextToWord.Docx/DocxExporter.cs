using AiTextToWord.Core.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AiTextToWord.Docx;

public sealed class DocxExporter
{
    private const int BulletNumberingId = 1;
    private const int OrderedNumberingId = 2;

    public void Export(DocumentModel document, Stream output, DocxExportOptions options)
    {
        using var word = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        AddStyles(mainPart, options);
        AddNumbering(mainPart);
        var body = mainPart.Document.Body!;

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
                body.Append(BlockQuoteParagraph(quotes));
                continue;
            }

            foreach (var paragraph in ToParagraphs(block, IsFirstBodyParagraph(document.Blocks, index), options))
            {
                body.Append(paragraph);
            }
        }

        body.Append(DefaultSectionProperties(options));
        mainPart.Document.Save();
    }

    private static IEnumerable<Paragraph> ToParagraphs(DocumentBlock block, bool isFirstBodyParagraph, DocxExportOptions options)
    {
        if (block is ListBlock list)
        {
            for (var index = 0; index < list.Items.Count; index++)
            {
                var inlines = list.ItemInlines is not null && index < list.ItemInlines.Count
                    ? list.ItemInlines[index]
                    : null;
                yield return ListParagraph(list.Items[index], inlines, list.IsOrdered, options);
            }

            yield break;
        }

        yield return block switch
        {
            HeadingBlock heading => ParagraphWithText(heading.Text, heading.Inlines, $"Heading{heading.Level}"),
            ParagraphBlock paragraph => ParagraphWithText(paragraph.Text, paragraph.Inlines, isFirstBodyParagraph ? "FirstParagraph" : "BodyText"),
            BlockQuoteBlock quote => BlockQuoteParagraph([quote]),
            CodeBlock code => ParagraphWithText(code.Code, style: "BlockText", font: "Consolas"),
            DividerBlock => SpacerParagraph(),
            _ => ParagraphWithText(string.Empty)
        };
    }

    private static void AddNumbering(MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            AbstractNumbering(0, NumberFormatValues.Bullet, "\u2022"),
            AbstractNumbering(1, NumberFormatValues.Decimal, "%1."),
            new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = BulletNumberingId },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = OrderedNumberingId });
    }

    private static void AddStyles(MainDocumentPart mainPart, DocxExportOptions options)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var bodySize = HalfPoints(options.BodyFontSize);
        var heading1Size = HalfPoints(options.HeadingFontSize);
        var heading2Size = HalfPoints(Math.Max(options.HeadingFontSize - 2, options.BodyFontSize + 2));
        var heading3Size = HalfPoints(Math.Max(options.HeadingFontSize - 4, options.BodyFontSize + 1));
        var line = LineSpacing(options.LineSpacing);
        var listSpacing = options.ListDensity == DocxListDensity.Comfortable ? "100" : "36";
        stylesPart.Styles = new Styles(
            ParagraphStyle(
                "Normal",
                "Normal",
                bodySize,
                options.FontFamily,
                bold: false,
                spacingAfter: "0",
                line: line,
                isDefault: true),
            ParagraphStyle(
                "FirstParagraph",
                "First Paragraph",
                bodySize,
                options.FontFamily,
                bold: false,
                spacingAfter: "180",
                line: line),
            ParagraphStyle(
                "BodyText",
                "Body Text",
                bodySize,
                options.FontFamily,
                bold: false,
                spacingAfter: "180",
                line: line,
                spacingBefore: "180"),
            ParagraphStyle(
                "BlockText",
                "Block Text",
                HalfPoints(Math.Max(options.BodyFontSize - 0.5, 10)),
                options.FontFamily,
                bold: false,
                spacingAfter: "100",
                line: line,
                spacingBefore: "100",
                color: "333333",
                leftIndent: "480",
                rightIndent: "480",
                shadingFill: options.QuoteStyle == DocxQuoteStyle.GrayBlock ? "F3F4F6" : null),
            ParagraphStyle(
                "Compact",
                "Compact",
                bodySize,
                options.FontFamily,
                bold: false,
                spacingAfter: listSpacing,
                line: line,
                spacingBefore: listSpacing),
            ParagraphStyle(
                "Heading1",
                "Heading 1",
                heading1Size,
                options.FontFamily,
                bold: true,
                spacingAfter: "80",
                line: line,
                spacingBefore: "360",
                color: "0F4761"),
            ParagraphStyle(
                "Heading2",
                "Heading 2",
                heading2Size,
                options.FontFamily,
                bold: true,
                spacingAfter: "80",
                line: line,
                spacingBefore: "160",
                color: "0F4761"),
            ParagraphStyle(
                "Heading3",
                "Heading 3",
                heading3Size,
                options.FontFamily,
                bold: true,
                spacingAfter: "80",
                line: line,
                spacingBefore: "160",
                color: "0F4761"));
    }

    private static Style ParagraphStyle(
        string styleId,
        string name,
        string fontSize,
        string fontFamily,
        bool bold,
        string spacingAfter,
        string line,
        string spacingBefore = "0",
        string? color = null,
        string? leftIndent = null,
        string? rightIndent = null,
        string? shadingFill = null,
        bool isDefault = false)
    {
        var runProperties = new StyleRunProperties(
            new RunFonts
            {
                Ascii = fontFamily,
                HighAnsi = fontFamily,
                EastAsia = fontFamily
            },
            new FontSize { Val = fontSize });

        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (!string.IsNullOrWhiteSpace(color))
        {
            runProperties.Append(new Color { Val = color });
        }

        var paragraphProperties = new StyleParagraphProperties(
            new SpacingBetweenLines
            {
                Before = spacingBefore,
                After = spacingAfter,
                Line = line,
                LineRule = LineSpacingRuleValues.Auto
            });
        if (!string.IsNullOrWhiteSpace(leftIndent) || !string.IsNullOrWhiteSpace(rightIndent))
        {
            paragraphProperties.Append(new Indentation
            {
                Left = leftIndent,
                Right = rightIndent,
                FirstLine = "0"
            });
        }

        if (!string.IsNullOrWhiteSpace(shadingFill))
        {
            paragraphProperties.Append(new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = shadingFill
            });
        }

        return new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new PrimaryStyle(),
            paragraphProperties,
            runProperties)
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
            Default = isDefault
        };
    }

    private static SectionProperties DefaultSectionProperties(DocxExportOptions options)
    {
        var margin = PageMarginTwips(options);

        return new SectionProperties(
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin
            {
                Top = margin,
                Right = (uint)margin,
                Bottom = margin,
                Left = (uint)margin,
                Header = 720,
                Footer = 720,
                Gutter = 0
            });
    }

    private static int PageMarginTwips(DocxExportOptions options)
    {
        if (options.PageMargin == DocxPageMargin.Custom && options.CustomPageMarginCentimeters is { } centimeters)
        {
            return (int)Math.Round(Math.Clamp(centimeters, 0.5, 6) * 567, MidpointRounding.AwayFromZero);
        }

        return options.PageMargin switch
        {
            DocxPageMargin.Narrow => 720,
            DocxPageMargin.Wide => 1800,
            _ => 1440
        };
    }

    private static string HalfPoints(double points)
    {
        return Math.Round(points * 2, MidpointRounding.AwayFromZero).ToString("0");
    }

    private static string LineSpacing(double lineSpacing)
    {
        return Math.Round(lineSpacing * 240, MidpointRounding.AwayFromZero).ToString("0");
    }

    private static AbstractNum AbstractNumbering(int id, NumberFormatValues format, string levelText)
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = format },
                new LevelText { Val = levelText },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new ParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" }))
            {
                LevelIndex = 0
            })
        {
            AbstractNumberId = id
        };
    }

    private static Paragraph ListParagraph(string text, IReadOnlyList<DocumentInline>? inlines, bool ordered, DocxExportOptions options)
    {
        var paragraph = ParagraphWithText(text, inlines, "Compact");
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.Append(new NumberingProperties(
            new NumberingLevelReference { Val = 0 },
            new NumberingId { Val = ordered ? OrderedNumberingId : BulletNumberingId }));
        return paragraph;
    }

    private static Paragraph BlockQuoteParagraph(IReadOnlyList<BlockQuoteBlock> quotes)
    {
        var inlines = new List<DocumentInline>();
        foreach (var quote in quotes)
        {
            if (inlines.Count > 0)
            {
                inlines.Add(new TextInline(" "));
            }

            if (quote.Inlines is { Count: > 0 })
            {
                inlines.AddRange(quote.Inlines);
            }
            else
            {
                inlines.Add(new TextInline(quote.Text));
            }
        }

        var text = string.Concat(inlines.Select(InlineText));
        return ParagraphWithText(text, inlines, "BlockText");
    }

    private static Paragraph SpacerParagraph()
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "80", After = "80" }));
    }

    private static Paragraph ParagraphWithText(
        string text,
        IReadOnlyList<DocumentInline>? inlines = null,
        string? style = null,
        bool italic = false,
        string? font = null)
    {
        var runProperties = new RunProperties();

        if (italic)
        {
            runProperties.Append(new Italic());
        }

        if (!string.IsNullOrWhiteSpace(font))
        {
            runProperties.Append(new RunFonts { Ascii = font, HighAnsi = font });
        }

        var paragraph = new Paragraph();
        if (!string.IsNullOrWhiteSpace(style))
        {
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style });
        }

        if (inlines is { Count: > 0 })
        {
            foreach (var inline in inlines)
            {
                paragraph.Append(RunForInline(inline, runProperties));
            }

            return paragraph;
        }

        AppendTextWithBreaks(paragraph, text, runProperties);
        return paragraph;
    }

    private static Run RunForInline(DocumentInline inline, RunProperties baseProperties)
    {
        var runProperties = (RunProperties)baseProperties.CloneNode(true);
        var text = inline switch
        {
            TextInline textInline => textInline.Text,
            BoldInline boldInline => boldInline.Text,
            ItalicInline italicInline => italicInline.Text,
            CodeInline codeInline => codeInline.Text,
            _ => string.Empty
        };

        switch (inline)
        {
            case BoldInline:
                runProperties.Append(new Bold());
                break;
            case ItalicInline:
                runProperties.Append(new Italic());
                break;
            case CodeInline:
                runProperties.Append(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" });
                break;
        }

        return new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static string InlineText(DocumentInline inline)
    {
        return inline switch
        {
            TextInline textInline => textInline.Text,
            BoldInline boldInline => boldInline.Text,
            ItalicInline italicInline => italicInline.Text,
            CodeInline codeInline => codeInline.Text,
            _ => string.Empty
        };
    }

    private static bool IsFirstBodyParagraph(IReadOnlyList<DocumentBlock> blocks, int index)
    {
        if (blocks[index] is not ParagraphBlock)
        {
            return false;
        }

        return index == 0 || blocks[index - 1] is not ParagraphBlock;
    }

    private static void AppendTextWithBreaks(Paragraph paragraph, string text, RunProperties runProperties)
    {
        var segments = text.Replace("\r\n", "\n").Split('\n');
        for (var index = 0; index < segments.Length; index++)
        {
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text(segments[index]) { Space = SpaceProcessingModeValues.Preserve }));
            if (index < segments.Length - 1)
            {
                paragraph.Append(new Run(new Break()));
            }
        }
    }
}
