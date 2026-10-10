namespace Sheetwright;

/// <summary>
/// Standard number, currency, percent and date format strings for spreadsheets.
/// </summary>
public static class SpreadsheetNumberFormats
{
    public const string Integer = "#,##0";
    public const string DecimalTwoPlaces = "#,##0.00";

    public const string CurrencyUsd = "$#,##0.00;($#,##0.00);\"-\"";
    public const string CurrencyUsdRounded = "$#,##0;($#,##0);\"-\"";
    public const string CurrencyEur = "[$€-2] #,##0.00;([$€-2] #,##0.00);\"-\"";
    public const string CurrencyTry = "[$₺-tr-TR] #,##0.00;([$₺-tr-TR] #,##0.00);\"-\"";
    public const string CurrencyGbp = "[$£-809] #,##0.00;([$£-809] #,##0.00);\"-\"";

    public const string Accounting = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
    public const string AccountingUsd = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";

    public const string Percent = "0.0%";
    public const string PercentTwoDecimals = "0.00%";
    public const string PercentRounded = "0%";

    public const string DateIso = "yyyy-mm-dd";
    public const string DateTimeIso = "yyyy-mm-dd hh:mm";
    public const string DateDisplay = "dd mmm yyyy";
}
