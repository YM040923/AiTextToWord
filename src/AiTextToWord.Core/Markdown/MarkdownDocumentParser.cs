using System.Text;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Model = AiTextToWord.Core.Model;
using MarkdownCodeInline = Markdig.Syntax.Inlines.CodeInline;

namespace AiTextToWord.Core.Markdown;

public sealed class MarkdownDocumentParser
{
    public Model.DocumentModel Parse(string text)
    {
        var normalizedText = NormalizeLooseTableSpacing(text);
        var markdown = Markdig.Markdown.Parse(normalizedText, ParserPipeline);
        var blocks = new List<Model.DocumentBlock>();

        foreach (var block in markdown)
        {
            AddBlock(block, blocks, isQuote: false, source: normalizedText);
        }

        return new Model.DocumentModel(blocks);
    }

    private static string NormalizeLooseTableSpacing(string text)
    {
        if (!text.Contains('|', StringComparison.Ordinal))
        {
            return text;
        }

        var lines = text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');
        var output = new List<string>(lines.Length);

        for (var index = 0; index < lines.Length; index++)
        {
            if (lines[index].Trim().Length == 0
                && PreviousNonEmptyLineIsTableLine(lines, index)
                && NextNonEmptyLineIsTableLine(lines, index))
            {
                continue;
            }

            output.Add(lines[index]);
        }

        return string.Join("\n", output);
    }

    private static bool PreviousNonEmptyLineIsTableLine(IReadOnlyList<string> lines, int index)
    {
        for (var cursor = index - 1; cursor >= 0; cursor--)
        {
            var line = lines[cursor].Trim();
            if (line.Length == 0)
            {
                continue;
            }

            return IsPipeTableLine(line);
        }

        return false;
    }

    private static bool NextNonEmptyLineIsTableLine(IReadOnlyList<string> lines, int index)
    {
        for (var cursor = index + 1; cursor < lines.Count; cursor++)
        {
            var line = lines[cursor].Trim();
            if (line.Length == 0)
            {
                continue;
            }

            return IsPipeTableLine(line);
        }

        return false;
    }

    private static bool IsPipeTableLine(string line)
    {
        return line.Count(character => character == '|') >= 2;
    }

    private static void AddBlock(Block block, List<Model.DocumentBlock> blocks, bool isQuote, string source)
    {
        switch (block)
        {
            case QuoteBlock quote:
                var quoteSource = ExtractSource(source, quote);
                foreach (var child in quote)
                {
                    if (child is Markdig.Syntax.ParagraphBlock paragraph)
                    {
                        AddTextBlock(
                            blocks,
                            paragraph.Inline,
                            headingLevel: null,
                            isQuote: true,
                            ExtractQuoteParagraphSource(quoteSource, ExtractSource(source, paragraph)));
                    }
                    else
                    {
                        AddBlock(child, blocks, isQuote: true, source);
                    }
                }
                break;

            case Markdig.Syntax.HeadingBlock heading:
                AddTextBlock(blocks, heading.Inline, heading.Level, isQuote, ExtractSource(source, heading));
                break;

            case Markdig.Syntax.ParagraphBlock paragraph:
                AddTextBlock(blocks, paragraph.Inline, headingLevel: null, isQuote, ExtractSource(source, paragraph));
                break;

            case Markdig.Syntax.ListBlock list:
                blocks.Add(ToListBlock(list, source));
                break;

            case Table table:
                var tableBlock = ToTableBlock(table, source);
                if (tableBlock is not null)
                {
                    blocks.Add(tableBlock);
                }
                break;

            case FencedCodeBlock fencedCode:
                blocks.Add(new Model.CodeBlock(
                    string.IsNullOrWhiteSpace(fencedCode.Info) ? null : fencedCode.Info,
                    fencedCode.Lines.ToString().TrimEnd()));
                break;

            case CodeBlock code:
                blocks.Add(new Model.CodeBlock(null, code.Lines.ToString().TrimEnd()));
                break;

            case ThematicBreakBlock:
                blocks.Add(new Model.DividerBlock());
                break;
        }
    }

    private static void AddTextBlock(List<Model.DocumentBlock> blocks, ContainerInline? inline, int? headingLevel, bool isQuote, string? source)
    {
        var parsed = ParseInline(inline, source);
        if (parsed.Text.Length == 0)
        {
            return;
        }

        if (isQuote)
        {
            blocks.Add(new Model.BlockQuoteBlock(parsed.Text, parsed.Inlines));
        }
        else if (headingLevel is { } level)
        {
            blocks.Add(new Model.HeadingBlock(level, parsed.Text, parsed.Inlines));
        }
        else
        {
            blocks.Add(new Model.ParagraphBlock(parsed.Text, parsed.Inlines));
        }
    }

    private static Model.ListBlock ToListBlock(Markdig.Syntax.ListBlock list, string source)
    {
        var items = new List<string>();
        var itemInlines = new List<IReadOnlyList<Model.DocumentInline>>();

        foreach (var child in list)
        {
            if (child is not ListItemBlock listItem)
            {
                continue;
            }

            var parsed = ParseListItem(listItem, source);
            if (parsed.Text.Length == 0)
            {
                continue;
            }

            items.Add(parsed.Text);
            itemInlines.Add(parsed.Inlines);
        }

        return new Model.ListBlock(list.IsOrdered, items, itemInlines);
    }

    private static Model.TableBlock? ToTableBlock(Table table, string source)
    {
        var rows = table
            .OfType<TableRow>()
            .Select(row => ToTableRow(row, source))
            .Where(row => row.Count > 0)
            .ToList();

        if (rows.Count == 0)
        {
            return null;
        }

        var headers = rows[0];
        var bodyRows = rows.Skip(1).ToList();
        return new Model.TableBlock(headers, bodyRows);
    }

    private static IReadOnlyList<Model.TableCell> ToTableRow(TableRow row, string source)
    {
        return row
            .OfType<TableCell>()
            .Select(cell => ToTableCell(cell, source))
            .ToList();
    }

    private static Model.TableCell ToTableCell(TableCell cell, string source)
    {
        var text = new StringBuilder();
        var inlines = new List<Model.DocumentInline>();

        foreach (var block in cell)
        {
            if (block is not Markdig.Syntax.ParagraphBlock paragraph)
            {
                continue;
            }

            var parsed = ParseInline(paragraph.Inline, ExtractSource(source, paragraph));
            if (parsed.Text.Length == 0)
            {
                continue;
            }

            if (text.Length > 0)
            {
                text.Append(' ');
                inlines.Add(new Model.TextInline(" "));
            }

            text.Append(parsed.Text);
            inlines.AddRange(parsed.Inlines);
        }

        return new Model.TableCell(text.ToString().Trim(), inlines);
    }

    private static InlineParseResult ParseListItem(ListItemBlock item, string source)
    {
        var text = new StringBuilder();
        var inlines = new List<Model.DocumentInline>();

        foreach (var block in item)
        {
            switch (block)
            {
                case Markdig.Syntax.ParagraphBlock paragraph:
                    AppendInline(ParseInline(paragraph.Inline, ExtractSource(source, paragraph)));
                    break;
                case Markdig.Syntax.HeadingBlock heading:
                    AppendInline(ParseInline(heading.Inline, ExtractSource(source, heading)));
                    break;
                case FencedCodeBlock fencedCode:
                    AppendPlain(fencedCode.Lines.ToString().TrimEnd());
                    break;
                case CodeBlock code:
                    AppendPlain(code.Lines.ToString().TrimEnd());
                    break;
            }
        }

        return new InlineParseResult(text.ToString().Trim(), inlines);

        void AppendInline(InlineParseResult parsed)
        {
            if (parsed.Text.Length == 0)
            {
                return;
            }

            if (text.Length > 0)
            {
                text.Append(' ');
                inlines.Add(new Model.TextInline(" "));
            }

            text.Append(parsed.Text);
            inlines.AddRange(parsed.Inlines);
        }

        void AppendPlain(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (text.Length > 0)
            {
                text.Append(' ');
                inlines.Add(new Model.TextInline(" "));
            }

            text.Append(value);
            inlines.Add(new Model.TextInline(value));
        }
    }

    private static InlineParseResult ParseInline(ContainerInline? inline, string? source)
    {
        var inlines = new List<Model.DocumentInline>();

        if (inline is not null)
        {
            foreach (var child in inline)
            {
                inlines.AddRange(ToDocumentInlines(child));
            }
        }

        var text = string.Concat(inlines.Select(InlineText)).Trim();
        var sourceInlines = TryParseSourceInlines(source, text, inlines);
        if (sourceInlines is not null)
        {
            inlines = sourceInlines.ToList();
            text = string.Concat(inlines.Select(InlineText)).Trim();
        }
        else if (HasResidualInlineMarkdown(text))
        {
            inlines = ParseLiteralInlines(text).ToList();
            text = string.Concat(inlines.Select(InlineText)).Trim();
        }

        return new InlineParseResult(text, inlines);
    }

    private static IReadOnlyList<Model.DocumentInline>? TryParseSourceInlines(
        string? source,
        string parsedText,
        IReadOnlyList<Model.DocumentInline> parsedInlines)
    {
        var normalizedSource = NormalizeInlineSource(source);
        if (!HasInlineMarkdownSource(normalizedSource))
        {
            return null;
        }

        var sourceInlines = ParseLiteralInlines(normalizedSource).ToList();
        var sourceText = string.Concat(sourceInlines.Select(InlineText)).Trim();
        var parsedPlainText = PlainTextFromInlineMarkdown(parsedText);
        if (!TextsEquivalent(sourceText, parsedPlainText))
        {
            return null;
        }

        return StyleInlineCount(sourceInlines) > StyleInlineCount(parsedInlines)
            ? sourceInlines
            : null;
    }

    private static bool HasResidualInlineMarkdown(string text)
    {
        return text.Contains("**", StringComparison.Ordinal)
            || text.Contains("__", StringComparison.Ordinal)
            || text.Contains('`')
            || text.Contains("](", StringComparison.Ordinal);
    }

    private static bool HasInlineMarkdownSource(string text)
    {
        return HasResidualInlineMarkdown(text)
            || text.Contains("![", StringComparison.Ordinal)
            || text.Contains('[', StringComparison.Ordinal);
    }

    private static IReadOnlyList<Model.DocumentInline> ToDocumentInlines(Inline inline)
    {
        return inline switch
        {
            LiteralInline literal => ParseLiteralInlines(literal.Content.ToString()),
            MarkdownCodeInline code => [new Model.CodeInline(code.Content)],
            LineBreakInline => [new Model.TextInline(" ")],
            LinkInline link => [new Model.TextInline(InlineText(link))],
            EmphasisInline emphasis => [ToEmphasisInline(emphasis)],
            ContainerInline container => FlattenContainer(container),
            _ => []
        };
    }

    private static IReadOnlyList<Model.DocumentInline> FlattenContainer(ContainerInline container)
    {
        var inlines = new List<Model.DocumentInline>();
        foreach (var child in container)
        {
            inlines.AddRange(ToDocumentInlines(child));
        }

        return inlines;
    }

    private static IReadOnlyList<Model.DocumentInline> ParseLiteralInlines(string source)
    {
        var inlines = new List<Model.DocumentInline>();
        var cursor = 0;

        foreach (Match match in InlineFallbackRegex.Matches(source))
        {
            if (match.Index > cursor)
            {
                inlines.Add(new Model.TextInline(source[cursor..match.Index]));
            }

            inlines.Add(CreateFallbackInline(match));
            cursor = match.Index + match.Length;
        }

        if (cursor < source.Length)
        {
            inlines.Add(new Model.TextInline(source[cursor..]));
        }

        return inlines;
    }

    private static Model.DocumentInline CreateFallbackInline(Match match)
    {
        if (match.Groups["imageAlt"].Success)
        {
            return new Model.TextInline(match.Groups["imageAlt"].Value);
        }

        if (match.Groups["linkText"].Success)
        {
            return new Model.TextInline(match.Groups["linkText"].Value);
        }

        if (match.Groups["code"].Success)
        {
            return new Model.CodeInline(match.Groups["code"].Value);
        }

        if (match.Groups["boldStar"].Success || match.Groups["boldUnder"].Success)
        {
            return new Model.BoldInline(match.Groups["boldStar"].Success ? match.Groups["boldStar"].Value : match.Groups["boldUnder"].Value);
        }

        return new Model.ItalicInline(match.Groups["italicStar"].Success ? match.Groups["italicStar"].Value : match.Groups["italicUnder"].Value);
    }

    private static Model.DocumentInline ToEmphasisInline(EmphasisInline emphasis)
    {
        var text = InlineText(emphasis);
        return emphasis.DelimiterCount >= 2
            ? new Model.BoldInline(text)
            : new Model.ItalicInline(text);
    }

    private static string InlineText(ContainerInline container)
    {
        var text = new StringBuilder();
        foreach (var child in container)
        {
            foreach (var inline in ToDocumentInlines(child))
            {
                text.Append(InlineText(inline));
            }
        }

        return text.ToString();
    }

    private static string InlineText(Model.DocumentInline inline)
    {
        return inline switch
        {
            Model.TextInline text => text.Text,
            Model.BoldInline bold => bold.Text,
            Model.ItalicInline italic => italic.Text,
            Model.CodeInline code => code.Text,
            _ => string.Empty
        };
    }

    private static string? ExtractSource(string source, Block block)
    {
        if (block.Span.Start < 0 || block.Span.End < block.Span.Start || block.Span.Start >= source.Length)
        {
            return null;
        }

        var end = Math.Min(block.Span.End, source.Length - 1);
        return source.Substring(block.Span.Start, end - block.Span.Start + 1);
    }

    private static string? ExtractQuoteParagraphSource(string? quoteSource, string? paragraphSource)
    {
        if (string.IsNullOrWhiteSpace(quoteSource))
        {
            return paragraphSource;
        }

        var lines = quoteSource
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');
        var paragraphLines = new List<string>();

        foreach (var line in lines)
        {
            var value = line.Trim();
            if (value.Length == 0)
            {
                if (paragraphLines.Count > 0)
                {
                    break;
                }

                continue;
            }

            var unquoted = Regex.Replace(value, @"^\s{0,3}>\s?", string.Empty);
            if (paragraphLines.Count > 0 && IsMarkdownListLine(unquoted))
            {
                break;
            }

            paragraphLines.Add(value);
        }

        return paragraphLines.Count == 0
            ? paragraphSource
            : string.Join("\n", paragraphLines);
    }

    private static bool IsMarkdownListLine(string line)
    {
        return Regex.IsMatch(line, @"^\s{0,3}(?:[-*+]\s+|\d+[.)]\s+)");
    }

    private static string NormalizeInlineSource(string? source)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            return string.Empty;
        }

        var lines = source
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');
        var normalized = new List<string>();

        foreach (var line in lines)
        {
            var value = line.Trim();
            if (value.Length == 0)
            {
                continue;
            }

            value = Regex.Replace(value, @"^\s{0,3}>\s?", string.Empty);
            value = Regex.Replace(value, @"^\s{0,3}(?:[-*+]\s+|\d+[.)]\s+)", string.Empty);
            value = Regex.Replace(value, @"^\s{0,3}#{1,6}\s*", string.Empty);
            value = Regex.Replace(value, @"\s+#{1,}\s*$", string.Empty);

            if (value.Length > 0)
            {
                normalized.Add(value.Trim());
            }
        }

        return string.Join(" ", normalized);
    }

    private static bool TextsEquivalent(string left, string right)
    {
        return NormalizeWhitespace(left).Equals(NormalizeWhitespace(right), StringComparison.Ordinal);
    }

    private static string PlainTextFromInlineMarkdown(string value)
    {
        if (!HasInlineMarkdownSource(value))
        {
            return value;
        }

        return string.Concat(ParseLiteralInlines(value).Select(InlineText)).Trim();
    }

    private static string NormalizeWhitespace(string value)
    {
        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static int StyleInlineCount(IReadOnlyList<Model.DocumentInline> inlines)
    {
        return inlines.Count(inline => inline is Model.BoldInline or Model.ItalicInline or Model.CodeInline);
    }

    private sealed record InlineParseResult(string Text, IReadOnlyList<Model.DocumentInline> Inlines);

    private static readonly MarkdownPipeline ParserPipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .Build();

    private static readonly Regex InlineFallbackRegex = new(
        @"!\[(?<imageAlt>[^\]]*)\]\([^)]+\)|\[(?<linkText>[^\]]+)\]\([^)]+\)|`(?<code>[^`]+)`|\*\*(?<boldStar>.+?)\*\*|__(?<boldUnder>.+?)__|(?<!\*)\*(?<italicStar>[^*]+)\*(?!\*)|(?<!_)_(?<italicUnder>[^_]+)_(?!_)",
        RegexOptions.Compiled);
}
