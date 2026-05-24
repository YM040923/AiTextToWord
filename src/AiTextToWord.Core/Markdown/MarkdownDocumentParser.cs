using AiTextToWord.Core.Model;

namespace AiTextToWord.Core.Markdown;

public sealed class MarkdownDocumentParser
{
    public DocumentModel Parse(string text)
    {
        var blocks = new List<DocumentBlock>();
        var lines = text.Replace("\r\n", "\n").Split('\n');

        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd();

            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (line.StartsWith("# "))
            {
                blocks.Add(new HeadingBlock(1, line[2..].Trim()));
                continue;
            }

            if (line.StartsWith("- "))
            {
                var items = new List<string>();
                while (index < lines.Length && lines[index].TrimStart().StartsWith("- "))
                {
                    items.Add(lines[index].Trim()[2..].Trim());
                    index++;
                }

                index--;
                blocks.Add(new ListBlock(false, items));
                continue;
            }

            blocks.Add(new ParagraphBlock(line.Trim()));
        }

        return new DocumentModel(blocks);
    }
}
