using AiTextToWord.Core.Fonts;
using Microsoft.Win32;
using PdfSharp.Fonts;

namespace AiTextToWord.Docx;

internal sealed class WindowsPdfFontResolver : IFontResolver
{
    private const string FontsKeyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";
    private static readonly Lazy<FontCatalog> Catalog = new(BuildCatalog);

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var catalog = Catalog.Value;
        var family = FindFamily(catalog, familyName)
            ?? FindFamily(catalog, "Microsoft YaHei")
            ?? FindFamily(catalog, "SimSun")
            ?? FindFamily(catalog, "Arial");

        if (family is null)
        {
            return null;
        }

        var face = family.Find(isBold, isItalic);
        return new FontResolverInfo(face.Path, isBold && !face.IsBold, isItalic && !face.IsItalic);
    }

    public byte[]? GetFont(string faceName)
    {
        return File.Exists(faceName) ? File.ReadAllBytes(faceName) : null;
    }

    private static FontFamilyEntry? FindFamily(FontCatalog catalog, string familyName)
    {
        return catalog.Families.TryGetValue(familyName, out var family) ? family : null;
    }

    private static FontCatalog BuildCatalog()
    {
        var families = new Dictionary<string, FontFamilyEntry>(StringComparer.CurrentCultureIgnoreCase);
        AddFontsFromRegistry(Registry.LocalMachine, families);
        AddFontsFromRegistry(Registry.CurrentUser, families);
        AddKnownFallbacks(families);
        return new FontCatalog(families);
    }

    private static void AddFontsFromRegistry(RegistryKey root, IDictionary<string, FontFamilyEntry> families)
    {
        using var key = root.OpenSubKey(FontsKeyPath);
        if (key is null)
        {
            return;
        }

        foreach (var valueName in key.GetValueNames())
        {
            if (key.GetValue(valueName) is not string fileName)
            {
                continue;
            }

            var path = FontPath(fileName);
            if (!File.Exists(path))
            {
                continue;
            }

            var isBold = valueName.Contains("Bold", StringComparison.OrdinalIgnoreCase);
            var isItalic = valueName.Contains("Italic", StringComparison.OrdinalIgnoreCase)
                || valueName.Contains("Oblique", StringComparison.OrdinalIgnoreCase);

            foreach (var familyName in WindowsFontNameParser.ExtractFamilyNames(valueName))
            {
                AddFace(families, familyName, new FontFace(path, isBold, isItalic, CollectionNumberFor(familyName, path)));
            }
        }
    }

    private static void AddKnownFallbacks(IDictionary<string, FontFamilyEntry> families)
    {
        AddKnownFace(families, "Microsoft YaHei", "Deng.ttf", isBold: false);
        AddKnownFace(families, "Microsoft YaHei", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "Microsoft YaHei", "msyh.ttc", isBold: false);
        AddKnownFace(families, "Microsoft YaHei", "msyhbd.ttc", isBold: true);
        AddKnownFace(families, "Microsoft YaHei UI", "Deng.ttf", isBold: false);
        AddKnownFace(families, "Microsoft YaHei UI", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "Microsoft YaHei UI", "msyh.ttc", isBold: false, collectionNumber: 1);
        AddKnownFace(families, "Microsoft YaHei UI", "msyhbd.ttc", isBold: true, collectionNumber: 1);
        AddKnownFace(families, "微软雅黑", "Deng.ttf", isBold: false);
        AddKnownFace(families, "微软雅黑", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "微软雅黑", "msyh.ttc", isBold: false);
        AddKnownFace(families, "微软雅黑", "msyhbd.ttc", isBold: true);
        AddKnownFace(families, "SimSun", "Deng.ttf", isBold: false);
        AddKnownFace(families, "SimSun", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "SimSun", "simsun.ttc", isBold: false);
        AddKnownFace(families, "宋体", "Deng.ttf", isBold: false);
        AddKnownFace(families, "宋体", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "宋体", "simsun.ttc", isBold: false);
        AddKnownFace(families, "DengXian", "Deng.ttf", isBold: false);
        AddKnownFace(families, "DengXian", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "等线", "Deng.ttf", isBold: false);
        AddKnownFace(families, "等线", "Dengb.ttf", isBold: true);
        AddKnownFace(families, "Consolas", "consola.ttf", isBold: false);
        AddKnownFace(families, "Consolas", "consolab.ttf", isBold: true);
        AddKnownFace(families, "Arial", "arial.ttf", isBold: false);
        AddKnownFace(families, "Arial", "arialbd.ttf", isBold: true);
    }

    private static void AddKnownFace(
        IDictionary<string, FontFamilyEntry> families,
        string familyName,
        string fileName,
        bool isBold,
        int collectionNumber = 0)
    {
        var path = FontPath(fileName);
        if (File.Exists(path))
        {
            AddFace(families, familyName, new FontFace(path, isBold, false, collectionNumber));
        }
    }

    private static void AddFace(IDictionary<string, FontFamilyEntry> families, string familyName, FontFace face)
    {
        if (!families.TryGetValue(familyName, out var family))
        {
            family = new FontFamilyEntry();
            families[familyName] = family;
        }

        family.Add(face);
    }

    private static string FontPath(string fileName)
    {
        if (Path.IsPathRooted(fileName))
        {
            return fileName;
        }

        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        return Path.Combine(windows, "Fonts", fileName);
    }

    private static int CollectionNumberFor(string familyName, string path)
    {
        if (!Path.GetExtension(path).Equals(".ttc", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        return familyName.EndsWith(" UI", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
    }

    private sealed record FontCatalog(IDictionary<string, FontFamilyEntry> Families);

    private sealed class FontFamilyEntry
    {
        private readonly List<FontFace> faces = [];

        public void Add(FontFace face)
        {
            if (!faces.Any(existing =>
                    string.Equals(existing.Path, face.Path, StringComparison.OrdinalIgnoreCase)
                    && existing.IsBold == face.IsBold
                    && existing.IsItalic == face.IsItalic
                    && existing.CollectionNumber == face.CollectionNumber))
            {
                faces.Add(face);
            }
        }

        public FontFace Find(bool isBold, bool isItalic)
        {
            return Preferred(faces.Where(face => face.IsBold == isBold && face.IsItalic == isItalic))
                ?? Preferred(faces.Where(face => face.IsBold == isBold))
                ?? Preferred(faces.Where(face => !face.IsBold && !face.IsItalic))
                ?? Preferred(faces)
                ?? faces[0];
        }

        private static FontFace? Preferred(IEnumerable<FontFace> candidates)
        {
            return candidates
                .OrderBy(face => Path.GetExtension(face.Path).Equals(".ttc", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();
        }
    }

    private sealed record FontFace(string Path, bool IsBold, bool IsItalic, int CollectionNumber);
}
