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

    [Fact]
    public void Convert_ParsesLooseAiMarkdownTables()
    {
        var service = TextConversionService.CreateDefault();

        var result = service.Convert("""
        这里有三个关键词：

        | 关键词 | 含义 |

        | ------ | ---------------------------- |

        | 污染物进入环境 | 有害物质进入空气、水、土壤等介质 |

        | 超过自净能力 | 环境本来能稀释、降解、转化部分污染物，但超过限度就失控 |

        | 影响健康 | 不一定立刻中毒，也可能是慢性损害、遗传损害、致癌、致畸等 |
        """);

        var table = Assert.IsType<TableBlock>(result.Document.Blocks[1]);
        Assert.Equal(["关键词", "含义"], table.Headers.Select(cell => cell.Text));
        Assert.Equal(3, table.Rows.Count);
    }
}
