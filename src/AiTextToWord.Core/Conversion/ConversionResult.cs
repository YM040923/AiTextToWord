using AiTextToWord.Core.Model;

namespace AiTextToWord.Core.Conversion;

public sealed record ConversionResult(string CleanedText, DocumentModel Document);
