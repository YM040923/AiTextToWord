using AiTextToWord.Core.Conversion;
using AiTextToWord.Core.Model;

namespace AiTextToWord.Core.Tests;

public sealed class TextConversionServiceTests
{
    [Fact]
    public void Convert_CleansAndParsesInput()
    {
        var service = TextConversionService.CreateDefault();

        var result = service.Convert("""
        > **[copy this to AI]**

        # Clean Title

        - Keep me
        """);

        Assert.DoesNotContain("copy this to AI", result.CleanedText, StringComparison.OrdinalIgnoreCase);
        Assert.Collection(result.Document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<ListBlock>(block));
    }
}
