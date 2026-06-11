using AiTextToWord.Core.Markdown;
using AiTextToWord.Core.Model;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf.IO;

namespace AiTextToWord.Docx.Tests;

public sealed class DocxExporterTests
{
    [Fact]
    public void PreviewLayoutMetrics_FromExportOptions_ReflectsPageAndTypographySettings()
    {
        var metrics = PreviewLayoutMetrics.FromExportOptions(
            new DocxExportOptions("Preview")
            {
                BodyFontSize = 12,
                HeadingFontSize = 22,
                LineSpacing = 1.5,
                PageMargin = DocxPageMargin.Wide,
                QuoteStyle = DocxQuoteStyle.GrayBlock,
                ListDensity = DocxListDensity.Comfortable
            });

        Assert.Equal(12, metrics.BodyFontSize);
        Assert.Equal(22, metrics.Heading1FontSize);
        Assert.Equal(18, metrics.Heading3FontSize);
        Assert.Equal(18, metrics.LineHeight);
        Assert.Equal(72, metrics.PagePadding);
        Assert.Equal(6, metrics.ListItemSpacing);
        Assert.True(metrics.UseGrayQuoteBlock);
    }

    [Fact]
    public void PdfExport_CreatesReadablePdf()
    {
        var document = new DocumentModel([
            new HeadingBlock(1, "PDF Export Title"),
            new ParagraphBlock("A paragraph for PDF export."),
            new ListBlock(false, ["One", "Two"]),
            new BlockQuoteBlock("A quote block."),
            new CodeBlock("csharp", "Console.WriteLine(\"hi\");")
        ]);
        using var stream = new MemoryStream();

        new PdfExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var pdf = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.Equal(1, pdf.PageCount);
    }

    [Fact]
    public void Export_CreatesReadableDocx()
    {
        var document = new DocumentModel([
            new HeadingBlock(1, "Export Title"),
            new ParagraphBlock("A paragraph."),
            new ListBlock(false, ["One", "Two"]),
            new CodeBlock("csharp", "Console.WriteLine(\"hi\");")
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        Assert.NotNull(word.MainDocumentPart);
        Assert.NotNull(word.MainDocumentPart.Document);
        var body = Assert.IsType<DocumentFormat.OpenXml.Wordprocessing.Body>(word.MainDocumentPart.Document.Body);
        var text = body.InnerText;
        Assert.Contains("Export Title", text);
        Assert.Contains("A paragraph.", text);
        Assert.Contains("Console.WriteLine", text);
    }

    [Fact]
    public void Export_WritesMarkdownListsAsWordNumberedParagraphs()
    {
        var document = new DocumentModel([
            new ListBlock(false, ["Bullet one", "Bullet two"]),
            new ListBlock(true, ["Step one", "Step two"])
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var mainPart = Assert.IsType<MainDocumentPart>(word.MainDocumentPart);
        Assert.NotNull(mainPart.NumberingDefinitionsPart);

        var exportedDocument = Assert.IsType<Document>(mainPart.Document);
        var body = Assert.IsType<Body>(exportedDocument.Body);
        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(4, paragraphs.Count);
        Assert.All(paragraphs, paragraph => Assert.NotNull(paragraph.ParagraphProperties?.NumberingProperties));
        Assert.All(paragraphs, paragraph => Assert.Equal("Compact", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value));
        Assert.Equal("Bullet one", paragraphs[0].InnerText);
        Assert.Equal("Bullet two", paragraphs[1].InnerText);
        Assert.Equal("Step one", paragraphs[2].InnerText);
        Assert.Equal("Step two", paragraphs[3].InnerText);
    }

    [Fact]
    public void Export_WritesMarkdownTablesAsWordTables()
    {
        var document = new MarkdownDocumentParser().Parse("""
        | Name | Status | Notes |
        | --- | --- | --- |
        | Word | **Ready** | Uses `docx` |
        | PDF | Preview | Export too |
        """);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        var table = Assert.Single(body.Elements<Table>());
        var borders = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableBorders>();
        Assert.NotNull(borders?.InsideHorizontalBorder);
        Assert.Equal(BorderValues.Single, borders!.InsideHorizontalBorder!.Val!.Value);
        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(["Name", "Status", "Notes"], rows[0].Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Select(cell => cell.InnerText));
        Assert.Equal(["Word", "Ready", "Uses docx"], rows[1].Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Select(cell => cell.InnerText));
        Assert.Contains(rows[1].Descendants<Run>(), run => run.RunProperties?.Bold is not null && run.InnerText == "Ready");
        Assert.Contains(rows[1].Descendants<Run>(), run => run.RunProperties?.RunFonts?.Ascii?.Value == "Consolas" && run.InnerText == "docx");
    }

    [Fact]
    public void Export_PreservesInlineBoldAndCodeRuns()
    {
        var document = new DocumentModel([
            new ParagraphBlock(
                "请按照以下“现代沉浸式”蓝图重构这个页面：",
                [
                    new TextInline("请按照以下"),
                    new BoldInline("“现代沉浸式”"),
                    new TextInline("蓝图重构这个页面：")
                ]),
            new ParagraphBlock(
                "请使用 filter: blur(80px) saturate(200%)。",
                [
                    new TextInline("请使用 "),
                    new CodeInline("filter: blur(80px) saturate(200%)"),
                    new TextInline("。")
                ])
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Contains(paragraphs[0].Elements<Run>(), run => run.RunProperties?.Bold is not null && run.InnerText == "“现代沉浸式”");
        Assert.Contains(paragraphs[1].Elements<Run>(), run => run.RunProperties?.RunFonts?.Ascii?.Value == "Consolas" && run.InnerText == "filter: blur(80px) saturate(200%)");
    }

    [Fact]
    public void Export_WritesSectionPropertiesForNormalWordLayout()
    {
        var document = new DocumentModel([new ParagraphBlock("正文")]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        Assert.NotNull(body.Elements<SectionProperties>().SingleOrDefault());
    }

    [Fact]
    public void Export_GroupsAdjacentQuoteBlocksAsOneBlockTextParagraph()
    {
        var document = new DocumentModel([
            new ParagraphBlock("引导段落"),
            new BlockQuoteBlock("第一句"),
            new BlockQuoteBlock(
                "第二句加粗",
                [
                    new TextInline("第二句"),
                    new BoldInline("加粗")
                ]),
            new ListBlock(false, ["列表项"]),
            new BlockQuoteBlock("收尾句")
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal("FirstParagraph", paragraphs[0].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("BlockText", paragraphs[1].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("第一句 第二句加粗", paragraphs[1].InnerText);
        Assert.DoesNotContain(paragraphs[1].Elements<Run>(), run => run.RunProperties?.Italic is not null);
        Assert.Contains(paragraphs[1].Elements<Run>(), run => run.RunProperties?.Bold is not null && run.InnerText == "加粗");
        Assert.Equal("Compact", paragraphs[2].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.NotNull(paragraphs[2].ParagraphProperties?.NumberingProperties);
        Assert.Equal("BlockText", paragraphs[3].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Export_UsesCustomTypographyLineSpacingAndMargins()
    {
        var document = new DocumentModel([
            new HeadingBlock(1, "标题"),
            new ParagraphBlock("正文")
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(
            document,
            stream,
            new DocxExportOptions("AI Text Export")
            {
                FontFamily = "SimSun",
                BodyFontSize = 12,
                HeadingFontSize = 18,
                LineSpacing = 1.5,
                PageMargin = DocxPageMargin.Wide
            });

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var styles = Assert.IsType<Styles>(word.MainDocumentPart!.StyleDefinitionsPart!.Styles);
        var normalStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "Normal");
        var headingStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "Heading1");
        Assert.Equal("SimSun", normalStyle.StyleRunProperties!.RunFonts!.EastAsia!.Value);
        Assert.Equal("24", normalStyle.StyleRunProperties!.FontSize!.Val!.Value);
        Assert.Equal("36", headingStyle.StyleRunProperties!.FontSize!.Val!.Value);

        var bodyTextStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "BodyText");
        var spacing = Assert.IsType<SpacingBetweenLines>(bodyTextStyle.StyleParagraphProperties!.GetFirstChild<SpacingBetweenLines>());
        Assert.Equal("360", spacing.Line!.Value);

        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart.Document).Body);
        var margins = Assert.IsType<SectionProperties>(body.Elements<SectionProperties>().Single()).GetFirstChild<PageMargin>();
        Assert.NotNull(margins);
        Assert.Equal(1800U, margins!.Left!.Value);
        Assert.Equal(1800U, margins.Right!.Value);
    }

    [Fact]
    public void Export_UsesCustomPageMarginCentimeters()
    {
        var document = new DocumentModel([new ParagraphBlock("正文")]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(
            document,
            stream,
            new DocxExportOptions("AI Text Export")
            {
                PageMargin = DocxPageMargin.Custom,
                CustomPageMarginCentimeters = 2
            });

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        var margins = Assert.IsType<SectionProperties>(body.Elements<SectionProperties>().Single()).GetFirstChild<PageMargin>();
        Assert.NotNull(margins);
        Assert.Equal(1134U, margins!.Left!.Value);
        Assert.Equal(1134U, margins.Right!.Value);
        Assert.Equal(1134, margins.Top!.Value);
        Assert.Equal(1134, margins.Bottom!.Value);
    }

    [Fact]
    public void Export_AppliesGrayQuoteBlockAndComfortableListDensity()
    {
        var document = new DocumentModel([
            new BlockQuoteBlock("引用内容"),
            new ListBlock(false, ["列表项"])
        ]);
        using var stream = new MemoryStream();

        new DocxExporter().Export(
            document,
            stream,
            new DocxExportOptions("AI Text Export")
            {
                QuoteStyle = DocxQuoteStyle.GrayBlock,
                ListDensity = DocxListDensity.Comfortable
            });

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var styles = Assert.IsType<Styles>(word.MainDocumentPart!.StyleDefinitionsPart!.Styles);
        var blockTextStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "BlockText");
        var blockShading = blockTextStyle.StyleParagraphProperties!.GetFirstChild<Shading>();
        Assert.NotNull(blockShading);
        Assert.Equal("F3F4F6", blockShading!.Fill!.Value);

        var compactStyle = Assert.Single(styles.Elements<Style>(), style => style.StyleId == "Compact");
        var spacing = Assert.IsType<SpacingBetweenLines>(compactStyle.StyleParagraphProperties!.GetFirstChild<SpacingBetweenLines>());
        Assert.Equal("100", spacing.Before!.Value);
        Assert.Equal("100", spacing.After!.Value);
    }

    [Fact]
    public void Export_FromAiMarkdownInput_WritesCleanWordStructure()
    {
        var document = new MarkdownDocumentParser().Parse("""
        ### 第三步：拿“中央主内容区（Main Content / 列表与卡片）”开刀 ###

        > 请按照以下**“现代沉浸式”**蓝图重构这个页面：
        > **1. 全屏动态毛玻璃背景（沉浸感的核心）**：
        > 废弃原本死板的纯色背景。请使用 `filter: blur(80px) saturate(200%)`。

        * **左栏（视觉焦点）**：展示一张**极其巨大**的专辑封面。
        """);
        using var stream = new MemoryStream();

        new DocxExporter().Export(document, stream, new DocxExportOptions("AI Text Export"));

        stream.Position = 0;
        using var word = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(Assert.IsType<Document>(word.MainDocumentPart!.Document).Body);
        var paragraphs = body.Elements<Paragraph>().ToList();
        var text = body.InnerText;
        Assert.DoesNotContain("**", text);
        Assert.DoesNotContain('`', text);
        Assert.Equal("Heading3", paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Contains(paragraphs.SelectMany(paragraph => paragraph.Elements<Run>()), run => run.RunProperties?.Bold is not null && run.InnerText == "“现代沉浸式”");
        Assert.Contains(paragraphs.SelectMany(paragraph => paragraph.Elements<Run>()), run => run.RunProperties?.RunFonts?.Ascii?.Value == "Consolas" && run.InnerText == "filter: blur(80px) saturate(200%)");
        Assert.Contains(paragraphs, paragraph => paragraph.ParagraphProperties?.NumberingProperties is not null && paragraph.InnerText == "左栏（视觉焦点）：展示一张极其巨大的专辑封面。");
    }
}
