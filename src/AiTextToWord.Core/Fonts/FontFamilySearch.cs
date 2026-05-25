namespace AiTextToWord.Core.Fonts;

public static class FontFamilySearch
{
    public static IReadOnlyList<string> Filter(
        IEnumerable<string> fonts,
        string? query,
        int maxResults = 80)
    {
        var normalizedQuery = query?.Trim();
        var availableFonts = fonts
            .Where(font => !string.IsNullOrWhiteSpace(font))
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .ToArray();

        if (string.IsNullOrWhiteSpace(normalizedQuery))
        {
            return availableFonts
                .Take(maxResults)
                .ToArray();
        }

        return availableFonts
            .Where(font => font.Contains(normalizedQuery, StringComparison.CurrentCultureIgnoreCase))
            .OrderBy(font => font.StartsWith(normalizedQuery, StringComparison.CurrentCultureIgnoreCase) ? 0 : 1)
            .ThenBy(font => font, StringComparer.CurrentCultureIgnoreCase)
            .Take(maxResults)
            .ToArray();
    }
}
