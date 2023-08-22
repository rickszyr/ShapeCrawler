using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal static class HexParser
{
    internal static Color GetSolidColorFromElement(TypedOpenXmlCompositeElement typedElement, SCSlideMaster slideMaster)
    {
        var colorString = FromSolidFill(typedElement, slideMaster!);
        var color = ColorTranslator.FromHtml($"#{colorString.Item2!}");
        float h = color.GetHue(), s = color.GetSaturation(), l = color.GetBrightness();

        var luminanceModulation = (float)(typedElement.Descendants<A.LuminanceModulation>().FirstOrDefault()?.Val ?? 100000);
        var luminanceOffset = (float)(typedElement.Descendants<A.LuminanceOffset>()?.FirstOrDefault()?.Val ?? 0);
        var alpha = (float)(typedElement.Descendants<A.Alpha>().FirstOrDefault()?.Val ?? 100000);
        var resultingColor = SKColor.FromHsl(h, s * 100, (l * 100 * luminanceModulation / 100000f) + (luminanceOffset / 1000f), (byte)(alpha / 100000f * 255));

        return Color.FromArgb(resultingColor.Alpha, resultingColor.Red, resultingColor.Green, resultingColor.Blue);
    }

    internal static (SCColorType, string?) FromSolidFill(TypedOpenXmlCompositeElement typedElement, SCSlideMaster slideMaster)
    {
        var colorHexVariant = GetWithoutScheme(typedElement);
        if (colorHexVariant is not null)
        {
            return ((SCColorType, string))colorHexVariant;
        }

        var aSchemeColor = typedElement.GetFirstChild<A.SchemeColor>() !;
        var fromScheme = GetByThemeColorScheme(aSchemeColor.Val!.InnerText!, slideMaster);
        return (SCColorType.Scheme, fromScheme);
    }

    internal static (SCColorType, string)? GetWithoutScheme(TypedOpenXmlCompositeElement typedElement)
    {
        var aSrgbClr = typedElement.GetFirstChild<A.RgbColorModelHex>();
        string colorHexVariant;
        if (aSrgbClr != null)
        {
            colorHexVariant = aSrgbClr.Val!;
            {
                return (SCColorType.RGB, colorHexVariant);
            }
        }

        var aSysClr = typedElement.GetFirstChild<A.SystemColor>();
        if (aSysClr != null)
        {
            colorHexVariant = aSysClr.LastColor!;
            {
                return (SCColorType.System, colorHexVariant);
            }
        }

        var aPresetColor = typedElement.GetFirstChild<A.PresetColor>();
        if (aPresetColor != null)
        {
            var coloName = aPresetColor.Val!.Value.ToString();
            {
                return (SCColorType.Preset, SCColorTranslator.HexFromName(coloName));
            }
        }

        return null;
    }

    private static string? GetByThemeColorScheme(string schemeColor, SCSlideMaster slideMaster)
    {
        var hex = GetThemeColorByString(schemeColor, slideMaster);

        if (hex == null)
        {
            hex = GetThemeMappedColor(schemeColor, slideMaster);
        }

        return hex ?? null;
    }

    private static string? GetThemeMappedColor(string fontSchemeColor, SCSlideMaster slideMaster)
    {
        var slideMasterPColorMap = slideMaster.PSlideMaster.ColorMap;
        var targetSchemeColor = slideMasterPColorMap?.GetAttributes().FirstOrDefault(a => a.LocalName == fontSchemeColor);
        return GetThemeColorByString(targetSchemeColor?.Value?.ToString() !, slideMaster);
    }

    private static string? GetThemeColorByString(string schemeColor, SCSlideMaster slideMaster)
    {
        var themeAColorScheme = slideMaster.ThemePart.Theme.ThemeElements!.ColorScheme!;
        var color = themeAColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        var hex = color?.RgbColorModelHex?.Val?.Value ?? color?.SystemColor?.LastColor?.Value;
        return hex;
    }
}