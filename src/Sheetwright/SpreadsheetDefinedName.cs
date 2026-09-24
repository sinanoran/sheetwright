using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetDefinedName
{
    internal SpreadsheetDefinedName(DefinedName definedName)
    {
        Inner = definedName;
    }

    internal DefinedName Inner { get; }

    public string Name => Inner.Name?.Value ?? string.Empty;

    public string Formula
    {
        get => Inner.Text ?? string.Empty;
        set => Inner.Text = Normalise(value ?? throw new ArgumentNullException(nameof(value)));
    }

    /// <summary>
    /// A formula as the file stores it: without the leading equals sign.
    /// </summary>
    /// <remarks>
    /// <b>The one place that decides this, and it has to stay the one place.</b>
    /// The rule lived here and again in <see cref="SpreadsheetDefinedNameCollection.AddFormula"/>,
    /// and the two copies had already drifted: adding a name stripped the equals
    /// sign and re-pointing an existing one did not, so the same value written
    /// through the two doors produced different files — and the one Excel repairs
    /// was the door with no test on it.
    ///
    /// Excel writes the sign when a person types a formula and stores what
    /// follows it; a defined name whose text begins with one is a name whose
    /// formula is "=…", which is not a formula.
    /// </remarks>
    internal static string Normalise(string formula) =>
        SpreadsheetTextSanitizer.Sanitize(formula.TrimStart('='));
}
