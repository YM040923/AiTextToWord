using System.Text;
using System.Text.RegularExpressions;

namespace AiTextToWord.Core.Cleanup;

public sealed partial class AiChatCleaner(TextNormalizer normalizer)
{
    public string Clean(string text)
    {
        var normalized = normalizer.Normalize(text);
        var output = new StringBuilder();
        var inFence = false;

        foreach (var line in normalized.Split('\n'))
        {
            var trimmed = line.Trim();

            if (trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                inFence = !inFence;
                output.AppendLine(line);
                continue;
            }

            if (!inFence && ShouldRemoveLine(trimmed))
            {
                continue;
            }

            output.AppendLine(line);
        }

        return normalizer.Normalize(output.ToString());
    }

    private static bool ShouldRemoveLine(string line)
    {
        if (line.Length == 0)
        {
            return false;
        }

        var unquoted = line.TrimStart('>', ' ');
        var unwrapped = unquoted.Trim('*', '_', '[', ']', ' ');

        if (CopyInstructionRegex().IsMatch(unwrapped))
        {
            return true;
        }

        if (unwrapped is "---" or "----")
        {
            return true;
        }

        return false;
    }

    [GeneratedRegex(@"^copy\s+this\s+(to|for)\s+ai$", RegexOptions.IgnoreCase)]
    private static partial Regex CopyInstructionRegex();
}
