namespace Sheetwright;

/// <summary>
/// Characters that make Excel read a cell as a formula rather than as text: what <see cref="SpreadsheetGridCsv"/>
/// guards on the way out and the CSV reader unguards on the way in.
/// </summary>
/// <remarks>
/// This is the CSV-only half of a real vulnerability. In a workbook a string is written as a string and Excel
/// never evaluates it; in CSV there is no type, so a display name of <c>=1+1</c> becomes a formula, and one of
/// <c>=HYPERLINK(...)</c> or a <c>DDE</c> call is an attack that runs on the machine of whoever opened the
/// export. The data here is user-supplied — display names, search terms, audit payloads — so this is reachable.
/// The last four leaders are the full-width forms of = + - @, which a spreadsheet opened with an East Asian
/// input locale can read as the ASCII sign. The writer and the reader share this one list so they cannot
/// disagree about what a guarded field looks like.
/// </remarks>
internal static class SpreadsheetFormulaGuard
{
    private static readonly char[] Leaders =
        ['=', '+', '-', '@', '\t', '\r', '＝', '＋', '－', '＠'];

    internal static bool IsLeader(char character) => Array.IndexOf(Leaders, character) >= 0;
}
