using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// One conditional format: the cells it covers and the rule that decides which of them it changes.
/// </summary>
/// <remarks>
/// <para>
/// The address is parsed the way a data validation's is, and for the same reason: it goes into
/// <c>sqref</c>, which has no room for a sheet qualifier, and a mistyped one should fail on the
/// line that wrote it rather than in Excel. A list of rectangles separated by spaces is accepted.
/// So is a whole column, unlike a data validation added through a range: <c>C:C</c> is the usual
/// way to format a column somebody will keep adding rows to, and it is one <c>sqref</c> rather
/// than a million cells.
/// </para>
/// <para>
/// Rules take priority in the order they were added to their sheet: the first one added is
/// priority 1, and wins where two rules set the same property on the same cell.
/// </para>
/// </remarks>
public abstract class SpreadsheetConditionalFormattingRule
{
    internal SpreadsheetConditionalFormattingRule(string address)
    {
        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        IReadOnlyList<string> parts = SpreadsheetValidationAddress.Split(address);
        if (parts.Count == 0)
        {
            throw new ArgumentException("A conditional formatting address is required.", nameof(address));
        }

        Address = string.Join(' ', parts.Select(part => SpreadsheetAddress.Parse(part).Reference));
    }

    /// <summary>
    /// The cells the rule covers, normalised: no sheet qualifier, no <c>$</c>, rectangles separated
    /// by a single space.
    /// </summary>
    public string Address { get; }

    /// <summary>
    /// Refuses a rule that would be written incomplete.
    /// </summary>
    /// <remarks>
    /// Called for every rule on every sheet before anything is committed, because committing is
    /// not undoable: a sheet half-written when the second one throws would be written twice on the
    /// next attempt.
    /// </remarks>
    internal abstract void Validate();

    internal abstract ConditionalFormattingRule ToOpenXml(int priority, SpreadsheetStyleRepository styles);
}
