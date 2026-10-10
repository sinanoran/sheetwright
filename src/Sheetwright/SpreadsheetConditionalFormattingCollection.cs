using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Sheetwright;

/// <summary>
/// The conditional formats on a worksheet, in priority order.
/// </summary>
public sealed class SpreadsheetConditionalFormattingCollection : IEnumerable<SpreadsheetConditionalFormattingRule>
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly List<SpreadsheetConditionalFormattingRule> _rules = new();

    internal SpreadsheetConditionalFormattingCollection(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    /// <summary>Formats cells whose value compares with a value or formula.</summary>
    public SpreadsheetCellValueConditionalFormattingRule AddCellValue(
        string address,
        SpreadsheetConditionalFormattingOperator @operator)
    {
        return Add(new SpreadsheetCellValueConditionalFormattingRule(address, @operator));
    }

    /// <summary>Formats cells for which a formula is true.</summary>
    public SpreadsheetFormulaConditionalFormattingRule AddFormula(string address)
    {
        return Add(new SpreadsheetFormulaConditionalFormattingRule(address));
    }

    public SpreadsheetDuplicateValuesConditionalFormattingRule AddDuplicateValues(string address)
    {
        return Add(new SpreadsheetDuplicateValuesConditionalFormattingRule(address));
    }

    public SpreadsheetUniqueValuesConditionalFormattingRule AddUniqueValues(string address)
    {
        return Add(new SpreadsheetUniqueValuesConditionalFormattingRule(address));
    }

    public SpreadsheetColorScaleConditionalFormattingRule AddTwoColorScale(string address)
    {
        return Add(new SpreadsheetColorScaleConditionalFormattingRule(address, threeColors: false));
    }

    public SpreadsheetColorScaleConditionalFormattingRule AddThreeColorScale(string address)
    {
        return Add(new SpreadsheetColorScaleConditionalFormattingRule(address, threeColors: true));
    }

    internal void Validate()
    {
        foreach (SpreadsheetConditionalFormattingRule rule in _rules)
        {
            rule.Validate();
        }
    }

    /// <remarks>
    /// Rules on the same address share one <c>conditionalFormatting</c> element, which is how Excel
    /// writes them; the priority is what orders them, and it runs across the whole sheet rather
    /// than restarting in each element.
    /// </remarks>
    internal void Commit(Worksheet worksheet, SpreadsheetStyleRepository styles)
    {
        worksheet.RemoveAllChildren<ConditionalFormatting>();
        if (_rules.Count == 0)
        {
            return;
        }

        List<ConditionalFormatting> blocks = new();
        Dictionary<string, ConditionalFormatting> blocksByAddress = new(StringComparer.Ordinal);
        int priority = 1;
        foreach (SpreadsheetConditionalFormattingRule rule in _rules)
        {
            if (!blocksByAddress.TryGetValue(rule.Address, out ConditionalFormatting? block))
            {
                block = new ConditionalFormatting
                {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = rule.Address }
                };
                blocksByAddress[rule.Address] = block;
                blocks.Add(block);
            }

            block.Append(rule.ToOpenXml(priority++, styles));
        }

        // conditionalFormatting follows mergeCells and precedes dataValidations in CT_Worksheet,
        // and the order is not advisory. Finding what has to come after it, rather than what comes
        // before, keeps this right whichever of the earlier elements a sheet happens to have.
        OpenXmlElement? following = worksheet.ChildElements.FirstOrDefault(FollowsConditionalFormatting);
        foreach (ConditionalFormatting block in blocks)
        {
            if (following is null)
            {
                worksheet.Append(block);
            }
            else
            {
                worksheet.InsertBefore(block, following);
            }
        }
    }

    public IEnumerator<SpreadsheetConditionalFormattingRule> GetEnumerator()
    {
        return _rules.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private static bool FollowsConditionalFormatting(OpenXmlElement element)
    {
        return element is DataValidations
            or Hyperlinks
            or PrintOptions
            or PageMargins
            or PageSetup
            or HeaderFooter
            or RowBreaks
            or ColumnBreaks
            or CustomProperties
            or CellWatches
            or IgnoredErrors
            or Drawing
            or LegacyDrawing
            or LegacyDrawingHeaderFooter
            or DrawingHeaderFooter
            or Picture
            or OleObjects
            or Controls
            or WebPublishItems
            or TableParts
            or WorksheetExtensionList;
    }

    private T Add<T>(T rule) where T : SpreadsheetConditionalFormattingRule
    {
        _worksheet.EnsureEditable();
        _rules.Add(rule);
        return rule;
    }
}
