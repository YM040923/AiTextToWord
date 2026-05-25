namespace AiTextToWord.Core.Export;

public static class ExportSettingRanges
{
    public const double MinimumBodyFontSize = 8;
    public const double MaximumBodyFontSize = 72;
    public const double MinimumHeadingFontSize = 10;
    public const double MaximumHeadingFontSize = 96;
    public const double MinimumLineSpacing = 1;
    public const double MaximumLineSpacing = 3;

    public static double ClampBodyFontSize(double value)
    {
        return Clamp(value, MinimumBodyFontSize, MaximumBodyFontSize);
    }

    public static double ClampHeadingFontSize(double value)
    {
        return Clamp(value, MinimumHeadingFontSize, MaximumHeadingFontSize);
    }

    public static double ClampLineSpacing(double value)
    {
        return Clamp(value, MinimumLineSpacing, MaximumLineSpacing);
    }

    public static double NumberFromSetting(
        object? value,
        double fallback,
        double minimum,
        double maximum,
        IReadOnlyList<double> legacyValues)
    {
        var number = value switch
        {
            int legacyIndex when legacyIndex >= 0 && legacyIndex < legacyValues.Count => legacyValues[legacyIndex],
            double doubleValue => doubleValue,
            float floatValue => floatValue,
            string stringValue when double.TryParse(stringValue, out var parsed) => parsed,
            _ => fallback
        };

        return Clamp(number, minimum, maximum);
    }

    private static double Clamp(double value, double minimum, double maximum)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return minimum;
        }

        return Math.Min(Math.Max(value, minimum), maximum);
    }
}
