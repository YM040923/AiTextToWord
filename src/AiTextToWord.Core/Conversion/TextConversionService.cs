using AiTextToWord.Core.Cleanup;
using AiTextToWord.Core.Markdown;

namespace AiTextToWord.Core.Conversion;

public sealed class TextConversionService(AiChatCleaner cleaner, MarkdownDocumentParser parser)
{
    public static TextConversionService CreateDefault()
    {
        return new TextConversionService(
            new AiChatCleaner(new TextNormalizer()),
            new MarkdownDocumentParser());
    }

    public ConversionResult Convert(string input)
    {
        var cleaned = cleaner.Clean(input);
        var document = parser.Parse(cleaned);
        return new ConversionResult(cleaned, document);
    }
}
