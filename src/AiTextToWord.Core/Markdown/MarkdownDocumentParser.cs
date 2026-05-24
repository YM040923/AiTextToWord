using System.Text;
using System.Text.RegularExpressions;
using AiTextToWord.Core.Model;

namespace AiTextToWord.Core.Markdown;

public sealed partial class MarkdownDocumentParser
{
    public DocumentModel Parse(string text)
    {
        var blocks = new List<DocumentBlock>();
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd();
            var trimmed = line.Trim();

            if (trimmed.Length == 0)
            {
                continue;
            }

            if (trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                var language = trimmed.Length > 3 ? trimmed[3..].Trim() : null;
                var code = new StringBuilder();
                index++;

                while (index < lines.Length && !lines[index].TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    code.AppendLine(lines[index].TrimEnd());
                    index++;
                }

                blocks.Add(new CodeBlock(string.IsNullOrWhiteSpace(language) ? null : language, code.ToString().TrimEnd()));
                continue;
            }

            if (trimmed is "---" or "***" or "___")
            {
                blocks.Add(new DividerBlock());
                continue;
            }

            if (trimmed.StartsWith("# ", StringComparison.Ordinal))
            {
                blocks.Add(new HeadingBlock(1, trimmed[2..].Trim()));
                continue;
            }

            if (trimmed.StartsWith("## ", StringComparison.Ordinal))
            {
                blocks.Add(new HeadingBlock(2, trimmed[3..].Trim()));
                continue;
            }

            if (trimmed.StartsWith("### ", StringComparison.Ordinal))
            {
                blocks.Add(new HeadingBlock(3, trimmed[4..].Trim()));
                continue;
            }

            if (trimmed.StartsWith("> ", StringComparison.Ordinal))
            {
                blocks.Add(new BlockQuoteBlock(trimmed[2..].Trim()));
                continue;
            }

            if (trimmed.StartsWith("- ", StringComparison.Ordinal) || trimmed.StartsWith("* ", StringComparison.Ordinal))
            {
                var items = new List<string>();
                while (index < lines.Length)
                {
                    var itemLine = lines[index].Trim();
                    if (!itemLine.StartsWith("- ", StringComparison.Ordinal) && !itemLine.StartsWith("* ", StringComparison.Ordinal))
                    {
                        break;
                    }

                    items.Add(itemLine[2..].Trim());
                    index++;
                }

                index--;
                blocks.Add(new ListBlock(false, items));
                continue;
            }

            if (OrderedListRegex().IsMatch(trimmed))
            {
                var items = new List<string>();
                while (index < lines.Length)
                {
                    var itemLine = lines[index].Trim();
                    var match = OrderedListRegex().Match(itemLine);
                    if (!match.Success)
                    {
                        break;
                    }

                    items.Add(match.Groups["text"].Value.Trim());
                    index++;
                }

                index--;
                blocks.Add(new ListBlock(true, items));
                continue;
            }

            blocks.Add(new ParagraphBlock(trimmed));
        }

        return new DocumentModel(blocks);
    }

    [GeneratedRegex(@"^\d+\.\s+(?<text>.+)$")]
    private static partial Regex OrderedListRegex();
}
