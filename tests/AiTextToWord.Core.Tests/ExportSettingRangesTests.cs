using AiTextToWord.Core.Export;

namespace AiTextToWord.Core.Tests;

public sealed class ExportSettingRangesTests
{
    [Theory]
    [InlineData(7, 8)]
    [InlineData(10.5, 10.5)]
    [InlineData(80, 72)]
    public void ClampBodyFontSize_ConstrainsValue(double value, double expected)
    {
        Assert.Equal(expected, ExportSettingRanges.ClampBodyFontSize(value));
    }

    [Theory]
    [InlineData(8, 10)]
    [InlineData(22, 22)]
    [InlineData(120, 96)]
    public void ClampHeadingFontSize_ConstrainsValue(double value, double expected)
    {
        Assert.Equal(expected, ExportSettingRanges.ClampHeadingFontSize(value));
    }

    [Theory]
    [InlineData(0.8, 1)]
    [InlineData(1.35, 1.35)]
    [InlineData(5, 3)]
    public void ClampLineSpacing_ConstrainsValue(double value, double expected)
    {
        Assert.Equal(expected, ExportSettingRanges.ClampLineSpacing(value));
    }

    [Fact]
    public void NumberFromSetting_MigratesLegacyComboBoxIndex()
    {
        var result = ExportSettingRanges.NumberFromSetting(2, fallback: 11, minimum: 8, maximum: 72, [10.5, 11, 12, 14]);

        Assert.Equal(12, result);
    }

    [Fact]
    public void NumberFromSetting_ReadsStoredDoubleAndClampsIt()
    {
        var result = ExportSettingRanges.NumberFromSetting(120d, fallback: 11, minimum: 8, maximum: 72, []);

        Assert.Equal(72, result);
    }
}
