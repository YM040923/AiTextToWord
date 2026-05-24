using AiTextToWord.Core.Cleanup;

namespace AiTextToWord.Core.Tests;

public sealed class AiChatCleanerTests
{
    [Fact]
    public void Clean_RemovesCopyInstructionLabelsAndReducesBlankLines()
    {
        var cleaner = new AiChatCleaner(new TextNormalizer());

        var cleaned = cleaner.Clean("""
        > **[copy this to AI]**


        # Useful Title



        Keep this paragraph.
        """);

        Assert.DoesNotContain("copy this to AI", cleaned, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("# Useful Title", cleaned);
        Assert.Contains("Keep this paragraph.", cleaned);
        Assert.DoesNotContain("\n\n\n", cleaned);
    }

    [Fact]
    public void Clean_PreservesFenceWhitespace()
    {
        var cleaner = new AiChatCleaner(new TextNormalizer());

        var cleaned = cleaner.Clean("""
        ```csharp
        if (true)
        {
            Console.WriteLine("kept");
        }
        ```
        """);

        Assert.Contains("    Console.WriteLine", cleaned);
    }
}
