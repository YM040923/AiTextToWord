using System.Text.RegularExpressions;

namespace AiTextToWord.Core.Fonts;

public static partial class WindowsFontNameParser
{
    private static readonly string[] StyleSuffixes =
    [
        "Regular",
        "Bold",
        "Italic",
        "Oblique",
        "Light",
        "Medium",
        "SemiBold",
        "Semibold",
        "SemiLight",
        "ExtraLight",
        "ExtraBold"
    ];

    public static IReadOnlyList<string> ExtractFamilyNames(string registryName)
    {
        if (string.IsNullOrWhiteSpace(registryName))
        {
            return [];
        }

        var name = FontTypeSuffixRegex().Replace(registryName, string.Empty);
        return name
            .Split('&', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Select(RemoveStyleSuffixes)
            .Where(font => !string.IsNullOrWhiteSpace(font) && !font.StartsWith('@'))
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    private static string RemoveStyleSuffixes(string name)
    {
        var result = WhitespaceRegex().Replace(name.Trim(), " ");
        var changed = true;

        while (changed)
        {
            changed = false;
            foreach (var suffix in StyleSuffixes)
            {
                if (result.EndsWith($" {suffix}", StringComparison.OrdinalIgnoreCase))
                {
                    result = result[..^(suffix.Length + 1)].Trim();
                    changed = true;
                    break;
                }
            }
        }

        return result;
    }

    [GeneratedRegex(@"\s*\((TrueType|OpenType|Type 1|Raster)\)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex FontTypeSuffixRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();
}
