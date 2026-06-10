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

    [Fact]
    public void Parse_SupportsFirstReleaseBlockTypes()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        ## Section

        > Important quote

        1. First step
        2. Second step

        ---

        ```powershell
        dotnet test
        ```
        """);

        Assert.Collection(document.Blocks,
            block =>
            {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(2, heading.Level);
            },
            block =>
            {
                var quote = Assert.IsType<BlockQuoteBlock>(block);
                Assert.Equal("Important quote", quote.Text);
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.True(list.IsOrdered);
                Assert.Equal(["First step", "Second step"], list.Items);
            },
            block => Assert.IsType<DividerBlock>(block),
            block =>
            {
                var code = Assert.IsType<CodeBlock>(block);
                Assert.Equal("powershell", code.Language);
                Assert.Equal("dotnet test", code.Code.Trim());
            });
    }

    [Fact]
    public void Parse_CreatesTableBlocksFromMarkdownPipeTables()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        | Name | Status | Notes |
        | --- | --- | --- |
        | Word | **Ready** | Uses `docx` |
        | PDF | _Preview_ | Export too |
        """);

        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        Assert.Equal(["Name", "Status", "Notes"], table.Headers.Select(cell => cell.Text));
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(["Word", "Ready", "Uses docx"], table.Rows[0].Select(cell => cell.Text));
        Assert.Equal(["PDF", "Preview", "Export too"], table.Rows[1].Select(cell => cell.Text));
        Assert.Contains(table.Rows[0][1].Inlines, inline => inline is BoldInline { Text: "Ready" });
        Assert.Contains(table.Rows[0][2].Inlines, inline => inline is CodeInline { Text: "docx" });
        Assert.Contains(table.Rows[1][1].Inlines, inline => inline is ItalicInline { Text: "Preview" });
    }

    [Fact]
    public void Parse_CreatesTableBlocksFromLooseAiMarkdownTables()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        这里有三个关键词：

        | 关键词 | 含义 |

        | ------ | ---------------------------- |

        | 污染物进入环境 | 有害物质进入空气、水、土壤等介质 |

        | 超过自净能力 | 环境本来能稀释、降解、转化部分污染物，但超过限度就失控 |

        | 影响健康 | 不一定立刻中毒，也可能是慢性损害、遗传损害、致癌、致畸等 |
        """);

        Assert.Collection(document.Blocks,
            block =>
            {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("这里有三个关键词：", paragraph.Text);
            },
            block =>
            {
                var table = Assert.IsType<TableBlock>(block);
                Assert.Equal(["关键词", "含义"], table.Headers.Select(cell => cell.Text));
                Assert.Equal(3, table.Rows.Count);
                Assert.Equal(["污染物进入环境", "有害物质进入空气、水、土壤等介质"], table.Rows[0].Select(cell => cell.Text));
                Assert.Equal(["超过自净能力", "环境本来能稀释、降解、转化部分污染物，但超过限度就失控"], table.Rows[1].Select(cell => cell.Text));
                Assert.Equal(["影响健康", "不一定立刻中毒，也可能是慢性损害、遗传损害、致癌、致畸等"], table.Rows[2].Select(cell => cell.Text));
            });
    }

    [Fact]
    public void Parse_StripsCommonInlineMarkdownMarkersFromAiChatText()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        ### 第三步：拿“中央主内容区（Main Content / 列表与卡片）”开刀 ###

        > **【复制这条给 AI】**
        > 请使用 `cc-haha` 打开 [项目](https://example.com)。

        * **无界整体感**：取消导航栏与右侧主内容区之间的实体分割线。

        1. **提取逻辑**：请先通读 `Sidebar.vue`。
        """);

        Assert.Collection(document.Blocks,
            block =>
            {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(3, heading.Level);
                Assert.Equal("第三步：拿“中央主内容区（Main Content / 列表与卡片）”开刀", heading.Text);
            },
            block =>
            {
                var quote = Assert.IsType<BlockQuoteBlock>(block);
                Assert.Equal("【复制这条给 AI】 请使用 cc-haha 打开 项目。", quote.Text);
                var inlines = Assert.IsAssignableFrom<IReadOnlyList<DocumentInline>>(quote.Inlines);
                Assert.Contains(inlines, inline => inline is BoldInline { Text: "【复制这条给 AI】" });
                Assert.Contains(inlines, inline => inline is CodeInline { Text: "cc-haha" });
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.False(list.IsOrdered);
                Assert.Equal(["无界整体感：取消导航栏与右侧主内容区之间的实体分割线。"], list.Items);
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.True(list.IsOrdered);
                Assert.Equal(["提取逻辑：请先通读 Sidebar.vue。"], list.Items);
            });
    }

    [Fact]
    public void Parse_ParsesQuotedAiInstructionsIntoStructuredBlocks()
    {
        var parser = new MarkdownDocumentParser();

        var document = parser.Parse("""
        > 请按照以下**“现代沉浸式”**蓝图重构这个页面：
        > **1. 全屏动态毛玻璃背景（沉浸感的核心）**：
        > 废弃原本死板的纯色背景。请使用 `filter: blur(80px) saturate(200%)`。
        > **2. 现代左右双栏布局（Flex/Grid）**：
        > * **左栏（视觉焦点）**：展示一张**极其巨大**的专辑封面。
        > * **右栏（沉浸式歌词）**：**正在播放的当前行歌词**必须加粗。
        """);

        Assert.Collection(document.Blocks,
            block =>
            {
                var quote = Assert.IsType<BlockQuoteBlock>(block);
                Assert.Equal("请按照以下“现代沉浸式”蓝图重构这个页面： 1. 全屏动态毛玻璃背景（沉浸感的核心）： 废弃原本死板的纯色背景。请使用 filter: blur(80px) saturate(200%)。 2. 现代左右双栏布局（Flex/Grid）：", quote.Text);
                var inlines = Assert.IsAssignableFrom<IReadOnlyList<DocumentInline>>(quote.Inlines);
                Assert.Contains(inlines, inline => inline is BoldInline { Text: "“现代沉浸式”" });
                Assert.Contains(inlines, inline => inline is BoldInline { Text: "1. 全屏动态毛玻璃背景（沉浸感的核心）" });
                Assert.Contains(inlines, inline => inline is CodeInline { Text: "filter: blur(80px) saturate(200%)" });
                Assert.Contains(inlines, inline => inline is BoldInline { Text: "2. 现代左右双栏布局（Flex/Grid）" });
            },
            block =>
            {
                var list = Assert.IsType<ListBlock>(block);
                Assert.Equal([
                    "左栏（视觉焦点）：展示一张极其巨大的专辑封面。",
                    "右栏（沉浸式歌词）：正在播放的当前行歌词必须加粗。"
                ], list.Items);
                var itemInlines = Assert.IsAssignableFrom<IReadOnlyList<IReadOnlyList<DocumentInline>>>(list.ItemInlines);
                Assert.Contains(itemInlines[0], inline => inline is BoldInline { Text: "左栏（视觉焦点）" });
                Assert.Contains(itemInlines[0], inline => inline is BoldInline { Text: "极其巨大" });
                Assert.Contains(itemInlines[1], inline => inline is BoldInline { Text: "正在播放的当前行歌词" });
            });
    }
}
