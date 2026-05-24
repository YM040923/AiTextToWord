using AiTextToWord.Core.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AiTextToWord.Docx;

public sealed class DocxExporter
{
    public void Export(DocumentModel document, Stream output, DocxExportOptions options)
    {
        using var word = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true);
        var mainPart = word.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        foreach (var block in document.Blocks)
        {
            body.Append(ToParagraph(block));
        }

        mainPart.Document.Save();
    }

    private static Paragraph ToParagraph(DocumentBlock block)
    {
        return block switch
        {
            HeadingBlock heading => ParagraphWithText(heading.Text, $"Heading{heading.Level}"),
            ParagraphBlock paragraph => ParagraphWithText(paragraph.Text),
            BlockQuoteBlock quote => ParagraphWithText(quote.Text, italic: true),
            CodeBlock code => ParagraphWithText(code.Code, font: "Consolas"),
            DividerBlock => ParagraphWithText(string.Empty),
            ListBlock list => ParagraphWithText(string.Join(Environment.NewLine, list.Items.Select(item => $"- {item}"))),
            _ => ParagraphWithText(string.Empty)
        };
    }

    private static Paragraph ParagraphWithText(string text, string? style = null, bool italic = false, string? font = null)
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

        var segments = text.Replace("\r\n", "\n").Split('\n');
        for (var index = 0; index < segments.Length; index++)
        {
            paragraph.Append(new Run(runProperties.CloneNode(true), new Text(segments[index]) { Space = SpaceProcessingModeValues.Preserve }));
            if (index < segments.Length - 1)
            {
                paragraph.Append(new Run(new Break()));
            }
        }

        return paragraph;
    }
}
