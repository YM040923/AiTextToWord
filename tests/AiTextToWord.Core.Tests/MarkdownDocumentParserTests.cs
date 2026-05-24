using AiTextToWord.Core.Markdown;
using AiTextToWord.Core.Model;

namespace AiTextToWord.Core.Tests;

public sealed class MarkdownDocumentParserTests
{
    [Fact]
    public void Parse_CreatesHeadingParagraphAndListBlocks()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        # Title

        A normal paragraph.

        - First
        - Second
        """);

        Assert.Collection(document.Blocks,
            block =>
            {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(1, heading.Level);
                Assert.Equal("Title", heading.Text);
            },
            block =>
            {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("A normal paragraph.", paragraph.Text);
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.False(list.IsOrdered);
                Assert.Equal(["First", "Second"], list.Items);
            });
    }
}
