namespace AiTextToWord.Core.Model;

public abstract record DocumentBlock;

public sealed record HeadingBlock(int Level, string Text) : DocumentBlock;

public sealed record ParagraphBlock(string Text) : DocumentBlock;

public sealed record ListBlock(bool IsOrdered, IReadOnlyList<string> Items) : DocumentBlock;

public sealed record BlockQuoteBlock(string Text) : DocumentBlock;

public sealed record CodeBlock(string? Language, string Code) : DocumentBlock;

public sealed record DividerBlock : DocumentBlock;
