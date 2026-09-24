using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Data validation: the dropdown and the bounds a workbook enforces on itself.
/// </summary>
/// <remarks>
/// These matter for a workbook sent out to be filled in and sent back. The bound
/// that is easy to get wrong is the one nobody set: an unassigned second bound
/// must be omitted rather than written as zero, because "between 5 and 0" is a
/// rule that rejects everything and reads, in Excel, as the file being broken.
/// </remarks>
public sealed class SpreadsheetDataValidationTests
{
    [Fact]
    public void A_list_validation_carries_the_formula_it_was_given()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetListDataValidation validation = sheet.DataValidations.AddListValidation("A1:A10");
            validation.Formula.ExcelFormula = "\"Yes,No\"";
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal(DataValidationValues.List, validation.Type!.Value);
        Assert.Equal("A1:A10", validation.SequenceOfReferences!.InnerText);
        Assert.Equal("\"Yes,No\"", validation.GetFirstChild<Formula1>()!.Text);
    }

    [Fact]
    public void A_list_validation_has_no_operator_because_there_is_nothing_to_compare()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.DataValidations.AddListValidation("A1").Formula.ExcelFormula = "\"Yes,No\"";
        });

        Assert.Null(Validations(written).Single().Operator);
    }

    [Fact]
    public void A_whole_number_between_two_bounds_writes_both_of_them()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetIntegerDataValidation validation = sheet.DataValidations.AddIntegerValidation("A1");
            validation.Formula.Value = 1;
            validation.Formula2.Value = 10;
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal(DataValidationValues.Whole, validation.Type!.Value);
        Assert.Equal(DataValidationOperatorValues.Between, validation.Operator!.Value);
        Assert.Equal("1", validation.GetFirstChild<Formula1>()!.Text);
        Assert.Equal("10", validation.GetFirstChild<Formula2>()!.Text);
    }

    [Fact]
    public void A_bound_nobody_set_is_left_out_rather_than_written_as_zero()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetIntegerDataValidation validation = sheet.DataValidations.AddIntegerValidation("A1");
            validation.Operator = SpreadsheetDataValidationOperator.GreaterThan;
            validation.Formula.Value = 0;
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal(DataValidationOperatorValues.GreaterThan, validation.Operator!.Value);
        Assert.Equal("0", validation.GetFirstChild<Formula1>()!.Text);
        Assert.Null(validation.GetFirstChild<Formula2>());
    }

    [Fact]
    public void A_decimal_bound_is_written_in_the_invariant_culture()
    {
        // The file format uses a full stop whatever the machine's culture says.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetDecimalDataValidation validation = sheet.DataValidations.AddDecimalValidation("A1");
            validation.Formula.Value = 1.5M;
            validation.Formula2.Value = 9.75M;
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal("1.5", validation.GetFirstChild<Formula1>()!.Text);
        Assert.Equal("9.75", validation.GetFirstChild<Formula2>()!.Text);
    }

    [Fact]
    public void A_date_bound_becomes_a_serial_number()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetDateTimeDataValidation validation = sheet.DataValidations.AddDateTimeValidation("A1");
            validation.Formula.Value = new DateTime(2026, 8, 27);
        });

        Assert.Equal(
            new DateTime(2026, 8, 27).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
            Validations(written).Single().GetFirstChild<Formula1>()!.Text);
    }

    [Fact]
    public void A_time_bound_is_a_fraction_of_a_day()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetTimeDataValidation validation = sheet.DataValidations.AddTimeValidation("A1");
            validation.Formula.Value!.Hour = 6;
        });

        Assert.Equal("0.25", Validations(written).Single().GetFirstChild<Formula1>()!.Text);
    }

    [Fact]
    public void A_text_length_validation_measures_the_text_and_not_the_number()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetIntegerDataValidation validation = sheet.DataValidations.AddTextLengthValidation("A1");
            validation.Operator = SpreadsheetDataValidationOperator.LessThanOrEqual;
            validation.Formula.Value = 200;
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal(DataValidationValues.TextLength, validation.Type!.Value);
        Assert.Equal(DataValidationOperatorValues.LessThanOrEqual, validation.Operator!.Value);
    }

    [Fact]
    public void A_validation_that_only_carries_a_message_has_no_formula()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetAnyDataValidation validation = sheet.DataValidations.AddAnyValidation("A1");
            validation.ShowInputMessage = true;
            validation.PromptTitle = "Reference";
            validation.Prompt = "The invoice reference.";
        });

        DataValidation validation = Validations(written).Single();

        Assert.Equal(DataValidationValues.None, validation.Type!.Value);
        Assert.Null(validation.GetFirstChild<Formula1>());
        Assert.True(validation.ShowInputMessage!.Value);
        Assert.Equal("Reference", validation.PromptTitle!.Value);
    }

    [Fact]
    public void An_error_message_is_carried_with_the_style_it_should_be_shown_in()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetAnyDataValidation validation = sheet.DataValidations.AddAnyValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Not allowed";
            validation.Error = "Pick a value from the list.";
            validation.ErrorStyle = SpreadsheetDataValidationWarningStyle.Warning;
        });

        DataValidation validation = Validations(written).Single();

        Assert.True(validation.ShowErrorMessage!.Value);
        Assert.Equal("Not allowed", validation.ErrorTitle!.Value);
        Assert.Equal(DataValidationErrorStyleValues.Warning, validation.ErrorStyle!.Value);
    }

    [Fact]
    public void A_message_is_cleaned_of_anything_XML_cannot_carry()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetAnyDataValidation validation = sheet.DataValidations.AddAnyValidation("A1");
            validation.ErrorTitle = "bro" + (char)1 + "ken";
        });

        Assert.Equal("broken", Validations(written).Single().ErrorTitle!.Value);
    }

    [Fact]
    public void Several_validations_are_counted()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.DataValidations.AddAnyValidation("A1");
            sheet.DataValidations.AddAnyValidation("A2");
        });

        DataValidations validations = written.Sheet("Data").GetFirstChild<DataValidations>()!;

        Assert.Equal(2U, validations.Count!.Value);
        Assert.Equal(2, validations.Elements<DataValidation>().Count());
    }

    [Fact]
    public void A_sheet_with_no_validations_has_no_validation_element()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "x");

        Assert.Null(written.Sheet("Data").GetFirstChild<DataValidations>());
    }

    [Fact]
    public void A_validation_added_through_a_range_takes_the_ranges_address()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Cells["B2:B9"].DataValidation.AddAnyDataValidation();
        });

        Assert.Equal("B2:B9", Validations(written).Single().SequenceOfReferences!.InnerText);
    }

    [Fact]
    public void Validating_a_whole_column_through_a_range_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<InvalidOperationException>(() => sheet.Cells["A:A"].DataValidation.AddAnyDataValidation());
    }

    [Fact]
    public void A_null_address_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => sheet.DataValidations.AddAnyValidation(null!));
    }

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("$B$2:$B$9", "B2:B9")]
    [InlineData("Sheet1!A1", "A1")]
    [InlineData("'Some Sheet'!$A$1:$C$3", "A1:C3")]
    public void A_validation_address_is_normalised_rather_than_copied(string given, string expected)
    {
        // sqref has no room for a sheet qualifier, and this was the one address
        // in the library that went into the file as the caller typed it.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Equal(expected, sheet.DataValidations.AddAnyValidation(given).Address.Address);
    }

    [Fact]
    public void A_validation_can_cover_more_than_one_rectangle()
    {
        // How one dropdown reaches columns B and E without being declared twice.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.DataValidations.AddAnyValidation("B1:B9 E1:E9");
        });

        Assert.Equal("B1:B9 E1:E9", Validations(written).Single().SequenceOfReferences!.InnerText);
    }

    [Theory]
    [InlineData("not an address")]
    [InlineData("A0")]
    [InlineData("   ")]
    public void A_validation_address_that_is_not_one_fails_where_it_was_written(string address)
    {
        // Rather than in Excel, days later, on somebody else's machine.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.ThrowsAny<Exception>(() => sheet.DataValidations.AddAnyValidation(address));
    }

    [Fact]
    public void The_validations_can_be_enumerated_before_they_are_written()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.DataValidations.AddAnyValidation("A1");
        sheet.DataValidations.AddListValidation("B1");

        Assert.Equal(2, sheet.DataValidations.Count());
        Assert.Equal(["A1", "B1"], sheet.DataValidations.Select(validation => validation.Address.Address));
    }

    private static IEnumerable<DataValidation> Validations(WrittenWorkbook written) =>
        written.Sheet("Data").GetFirstChild<DataValidations>()!.Elements<DataValidation>();
}
