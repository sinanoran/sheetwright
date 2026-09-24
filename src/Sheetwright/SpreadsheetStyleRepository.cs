using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace Sheetwright;

internal sealed class SpreadsheetStyleRepository
{
    private readonly Stylesheet _stylesheet;
    private readonly Dictionary<string, uint> _styleIndexes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _fontIndexes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _fillIndexes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _borderIndexes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _numberFormats = new(StringComparer.Ordinal);

    internal SpreadsheetStyleRepository(WorkbookPart workbookPart)
    {
        WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        _stylesheet = CreateBaseStylesheet();
        stylesPart.Stylesheet = _stylesheet;
        _styleIndexes[CreateKey(new SpreadsheetStylePatch())] = 0;
        AddExistingIndexes(_stylesheet.Fonts!.Elements<Font>(), _fontIndexes);
        AddExistingIndexes(_stylesheet.Fills!.Elements<Fill>(), _fillIndexes);
        AddExistingIndexes(_stylesheet.Borders!.Elements<Border>(), _borderIndexes);
    }

    internal uint GetStyleIndex(SpreadsheetStylePatch patch)
    {
        string key = CreateKey(patch);
        if (_styleIndexes.TryGetValue(key, out uint existing))
        {
            return existing;
        }

        uint fontId = AddFont(_stylesheet, patch);
        uint fillId = AddFill(_stylesheet, patch);
        uint borderId = AddBorder(_stylesheet, patch);
        uint numberFormatId = AddNumberFormat(_stylesheet, patch.NumberFormat);

        CellFormat format = new()
        {
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId,
            NumberFormatId = numberFormatId,
            ApplyFont = fontId != 0,
            ApplyFill = fillId != 0,
            ApplyBorder = borderId != 0,
            ApplyNumberFormat = numberFormatId != 0
        };

        if (patch.HorizontalAlignment.HasValue || patch.WrapText.HasValue)
        {
            format.Alignment = new Alignment
            {
                Horizontal = patch.HorizontalAlignment switch
                {
                    SpreadsheetHorizontalAlignment.Left => HorizontalAlignmentValues.Left,
                    SpreadsheetHorizontalAlignment.Center => HorizontalAlignmentValues.Center,
                    SpreadsheetHorizontalAlignment.Right => HorizontalAlignmentValues.Right,
                    _ => HorizontalAlignmentValues.General
                },
                WrapText = patch.WrapText
            };
            format.ApplyAlignment = true;
        }

        if (patch.Locked.HasValue)
        {
            format.Protection = new Protection { Locked = patch.Locked.Value };
            format.ApplyProtection = true;
        }

        _stylesheet.CellFormats!.Append(format);
        _stylesheet.CellFormats.Count = (uint)_stylesheet.CellFormats.ChildElements.Count;
        uint index = _stylesheet.CellFormats.Count.Value - 1;
        _styleIndexes[key] = index;
        return index;
    }

    internal void Save()
    {
        _stylesheet.Save();
    }

    private uint AddFont(Stylesheet stylesheet, SpreadsheetStylePatch patch)
    {
        if (patch.FontName is null && patch.FontSize is null && patch.Bold is null && patch.Italic is null && patch.FontColor is null)
        {
            return 0;
        }

        Font font = new();
        if (patch.Bold == true)
        {
            font.Append(new Bold());
        }

        if (patch.Italic == true)
        {
            font.Append(new Italic());
        }

        font.Append(new FontSize { Val = patch.FontSize ?? 11D });
        if (!string.IsNullOrEmpty(patch.FontColor))
        {
            font.Append(new Color { Rgb = patch.FontColor });
        }

        font.Append(new FontName { Val = patch.FontName ?? "Calibri" });

        string xml = font.OuterXml;
        if (_fontIndexes.TryGetValue(xml, out uint existing))
        {
            return existing;
        }

        Fonts fonts = stylesheet.Fonts ?? throw new InvalidOperationException("The font collection is missing.");
        fonts.Append(font);
        fonts.Count = (uint)fonts.ChildElements.Count;
        uint index = fonts.Count.Value - 1;
        _fontIndexes[xml] = index;
        return index;
    }

    private uint AddFill(Stylesheet stylesheet, SpreadsheetStylePatch patch)
    {
        if (string.IsNullOrEmpty(patch.FillColor))
        {
            return 0;
        }

        Fill fill = new(new PatternFill(
            new ForegroundColor { Rgb = patch.FillColor },
            new BackgroundColor { Indexed = 64U })
        {
            PatternType = PatternValues.Solid
        });

        string xml = fill.OuterXml;
        if (_fillIndexes.TryGetValue(xml, out uint existing))
        {
            return existing;
        }

        Fills fills = stylesheet.Fills ?? throw new InvalidOperationException("The fill collection is missing.");
        fills.Append(fill);
        fills.Count = (uint)fills.ChildElements.Count;
        uint index = fills.Count.Value - 1;
        _fillIndexes[xml] = index;
        return index;
    }

    private uint AddBorder(Stylesheet stylesheet, SpreadsheetStylePatch patch)
    {
        if (patch.LeftBorder is null && patch.RightBorder is null && patch.TopBorder is null && patch.BottomBorder is null)
        {
            return 0;
        }

        Border border = new(
            CreateBorder<LeftBorder>(patch.LeftBorder, patch.BorderColor),
            CreateBorder<RightBorder>(patch.RightBorder, patch.BorderColor),
            CreateBorder<TopBorder>(patch.TopBorder, patch.BorderColor),
            CreateBorder<BottomBorder>(patch.BottomBorder, patch.BorderColor),
            new DiagonalBorder());

        string xml = border.OuterXml;
        if (_borderIndexes.TryGetValue(xml, out uint existing))
        {
            return existing;
        }

        Borders borders = stylesheet.Borders ?? throw new InvalidOperationException("The border collection is missing.");
        borders.Append(border);
        borders.Count = (uint)borders.ChildElements.Count;
        uint index = borders.Count.Value - 1;
        _borderIndexes[xml] = index;
        return index;
    }

    private uint AddNumberFormat(Stylesheet stylesheet, string? format)
    {
        if (format is null || format.Trim().Length == 0)
        {
            return 0;
        }

        if (_numberFormats.TryGetValue(format, out uint existing))
        {
            return existing;
        }

        uint id = 164U + (uint)_numberFormats.Count;
        stylesheet.NumberingFormats ??= new NumberingFormats();
        stylesheet.NumberingFormats.Append(new NumberingFormat { NumberFormatId = id, FormatCode = format });
        stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.ChildElements.Count;
        _numberFormats[format] = id;
        return id;
    }

    private static T CreateBorder<T>(SpreadsheetBorderStyle? style, string? color) where T : BorderPropertiesType, new()
    {
        T border = new()
        {
            Style = style switch
            {
                SpreadsheetBorderStyle.Thin => BorderStyleValues.Thin,
                SpreadsheetBorderStyle.Medium => BorderStyleValues.Medium,
                _ => BorderStyleValues.None
            }
        };

        if (!string.IsNullOrEmpty(color) && style is not null and not SpreadsheetBorderStyle.None)
        {
            border.Color = new Color { Rgb = color };
        }

        return border;
    }

    private static string CreateKey(SpreadsheetStylePatch patch)
    {
        StringBuilder key = new();
        AppendKeyPart(key, patch.FontName);
        AppendKeyPart(key, patch.FontSize?.ToString(CultureInfo.InvariantCulture));
        AppendKeyPart(key, patch.Bold?.ToString());
        AppendKeyPart(key, patch.Italic?.ToString());
        AppendKeyPart(key, patch.FontColor);
        AppendKeyPart(key, patch.FillColor);
        AppendKeyPart(key, patch.NumberFormat);
        AppendKeyPart(key, patch.HorizontalAlignment?.ToString());
        AppendKeyPart(key, patch.WrapText?.ToString());
        AppendKeyPart(key, patch.Locked?.ToString());
        AppendKeyPart(key, patch.LeftBorder?.ToString());
        AppendKeyPart(key, patch.RightBorder?.ToString());
        AppendKeyPart(key, patch.TopBorder?.ToString());
        AppendKeyPart(key, patch.BottomBorder?.ToString());
        AppendKeyPart(key, patch.BorderColor);
        return key.ToString();
    }

    private static void AppendKeyPart(StringBuilder key, string? value)
    {
        if (value is null)
        {
            key.Append("-1:");
            return;
        }

        key.Append(value.Length.ToString(CultureInfo.InvariantCulture));
        key.Append(':');
        key.Append(value);
    }

    private static void AddExistingIndexes<T>(IEnumerable<T> elements, Dictionary<string, uint> indexes)
        where T : OpenXmlElement
    {
        uint index = 0;
        foreach (T element in elements)
        {
            indexes[element.OuterXml] = index;
            index++;
        }
    }

    private static Stylesheet CreateBaseStylesheet()
    {
        return new Stylesheet(
            new NumberingFormats(),
            new Fonts(new Font(
                new FontSize { Val = 11D },
                new FontName { Val = "Calibri" },
                new FontFamilyNumbering { Val = 2 }))
            { Count = 1U },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
            { Count = 2U },
            new Borders(new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()))
            { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            })
            { Count = 1U },
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U },
            new DifferentialFormats { Count = 0U },
            new TableStyles
            {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            });
    }
}
