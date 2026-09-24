using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Sheetwright;

public sealed class SpreadsheetDefinedNameCollection : IEnumerable<SpreadsheetDefinedName>
{
    private readonly Workbook _workbook;

    internal SpreadsheetDefinedNameCollection(WorkbookPart workbookPart)
    {
        _workbook = workbookPart.Workbook ?? throw new InvalidDataException("The workbook is missing.");
    }

    public SpreadsheetDefinedName this[string name]
    {
        get
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            string sanitizedName = SpreadsheetTextSanitizer.Sanitize(name);
            DefinedName definedName = GetDefinedNames().Elements<DefinedName>()
                    .FirstOrDefault(x => string.Equals(x.Name?.Value, sanitizedName, StringComparison.OrdinalIgnoreCase))
                // First() would report "sequence contains no matching element",
                // which names LINQ rather than the workbook and sends the reader
                // looking for a bug in this class.
                ?? throw new KeyNotFoundException($"The workbook has no defined name '{name}'.");
            return new SpreadsheetDefinedName(definedName);
        }
    }

    public bool ContainsKey(string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        string sanitizedName = SpreadsheetTextSanitizer.Sanitize(name);
        return GetDefinedNames().Elements<DefinedName>()
            .Any(x => string.Equals(x.Name?.Value, sanitizedName, StringComparison.OrdinalIgnoreCase));
    }

    public SpreadsheetDefinedName Add(string name, SpreadsheetRange range)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        if (range is null)
        {
            throw new ArgumentNullException(nameof(range));
        }

        string formula = $"'{range.Worksheet.Name.Replace("'", "''")}'!${SpreadsheetAddress.GetColumnName(range.Start.Column)}${range.Start.Row}";
        if (range.Start.Row != range.End.Row || range.Start.Column != range.End.Column)
        {
            formula += $":${SpreadsheetAddress.GetColumnName(range.End.Column)}${range.End.Row}";
        }
        return AddFormula(name, formula);
    }

    public SpreadsheetDefinedName AddFormula(string name, string formula)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        // Through the property rather than straight onto Text, so that adding a
        // name and re-pointing one normalise identically. Writing it here as well
        // is what let the two drift apart in the first place.
        DefinedName definedName = new() { Name = SpreadsheetTextSanitizer.Sanitize(name) };
        SpreadsheetDefinedName defined = new(definedName) { Formula = formula };

        GetDefinedNames().Append(definedName);
        return defined;
    }

    public IEnumerator<SpreadsheetDefinedName> GetEnumerator()
    {
        return GetDefinedNames().Elements<DefinedName>().Select(x => new SpreadsheetDefinedName(x)).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private DefinedNames GetDefinedNames()
    {
        if (_workbook.DefinedNames is DefinedNames definedNames)
        {
            return definedNames;
        }

        DefinedNames newDefinedNames = new();
        _workbook.Append(newDefinedNames);
        return newDefinedNames;
    }
}
