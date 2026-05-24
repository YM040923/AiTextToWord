using AiTextToWord.Core.Model;
using DocumentFormat.OpenXml.Packaging;

namespace AiTextToWord.Docx.Tests;

public sealed class DocxExporterTests
{
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
}
