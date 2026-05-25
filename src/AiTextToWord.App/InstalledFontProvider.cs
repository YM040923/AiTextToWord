using AiTextToWord.Core.Fonts;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AiTextToWord.App;

internal static class InstalledFontProvider
{
    private const string FontsKeyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

    private static readonly string[] FallbackFonts =
    [
        "微软雅黑",
        "Microsoft YaHei",
        "等线",
        "宋体",
        "Arial"
    ];

    public static IReadOnlyList<string> GetInstalledFontFamilies()
    {
        var fonts = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        AddFontsFromRegistry(Registry.CurrentUser, fonts);
        AddFontsFromRegistry(Registry.LocalMachine, fonts);

        foreach (var font in FallbackFonts)
        {
            fonts.Add(font);
        }

        return fonts
            .OrderBy(font => font, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    private static void AddFontsFromRegistry(RegistryKey root, ISet<string> fonts)
    {
        using var key = root.OpenSubKey(FontsKeyPath);
        if (key is null)
        {
            return;
        }

        foreach (var valueName in key.GetValueNames())
        {
            foreach (var font in WindowsFontNameParser.ExtractFamilyNames(valueName))
            {
                fonts.Add(font);
            }
        }
    }
}
