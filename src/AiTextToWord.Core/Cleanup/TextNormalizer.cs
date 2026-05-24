using System.Text;

namespace AiTextToWord.Core.Cleanup;

public sealed class TextNormalizer
{
    public string Normalize(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        var result = new StringBuilder();
        var blankCount = 0;
        var inFence = false;

        foreach (var rawLine in normalized.Split('\n'))
        {
            var line = rawLine.TrimEnd();

            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                inFence = !inFence;
                blankCount = 0;
                result.AppendLine(line);
                continue;
            }

            if (!inFence && string.IsNullOrWhiteSpace(line))
            {
                blankCount++;
                if (blankCount <= 1)
                {
                    result.AppendLine();
                }
                continue;
            }

            blankCount = 0;
            result.AppendLine(inFence ? rawLine.TrimEnd() : line.Trim());
        }

        return result.ToString().Trim();
    }
}
