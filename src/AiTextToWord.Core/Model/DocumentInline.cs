namespace AiTextToWord.Core.Model;

public abstract record DocumentInline;

public sealed record TextInline(string Text) : DocumentInline;

public sealed record BoldInline(string Text) : DocumentInline;

public sealed record ItalicInline(string Text) : DocumentInline;

public sealed record CodeInline(string Text) : DocumentInline;
