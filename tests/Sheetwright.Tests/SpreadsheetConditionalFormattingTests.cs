using DocumentFormat.OpenXml.Spreadsheet;
using Drawing = System.Drawing;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Conditional formatting: the rules a workbook applies to itself, and the formats they apply.
/// </summary>
/// <remarks>
/// <para>
/// Two tables are involved and both are easy to get quietly wrong. The rules live in the worksheet,
/// where <c>conditionalFormatting</c> has a fixed place in CT_Worksheet and a misplaced one is a
/// file Excel offers to repair. Their formatting lives in the stylesheet's <c>dxfs</c>, which has to
/// be de-duplicated for the same reason the cell style table is — and which differs from a cell
/// format in ways that only show when the workbook is opened.
/// </para>
/// <para>
/// So these tests read the XML back, and the ones that matter most also run it through
/// <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/>.
/// </para>
/// </remarks>
public sealed class SpreadsheetConditionalFormattingTests
{
    // ---------------------------------------------------------------- cell value comparisons

    [Fact]
    public void A_cell_value_rule_compares_each_cell_with_the_value_it_was_given()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            SpreadsheetCellValueConditionalFormattingRule rule =
                sheet.ConditionalFormatting.AddCellValue("C2:C9", SpreadsheetConditionalFormattingOperator.GreaterThan);
            rule.Formula.Value = 1000;
            rule.Style.Fill.BackgroundColor.SetColor("FFC7CE");
        });

        ConditionalFormatting block = Blocks(written).Single();
        ConditionalFormattingRule rule = block.Elements<ConditionalFormattingRule>().Single();

        Assert.Equal("C2:C9", block.SequenceOfReferences!.InnerText);
        Assert.Equal(ConditionalFormatValues.CellIs, rule.Type!.Value);
        Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, rule.Operator!.Value);
        Assert.Equal(["1000"], Formulas(rule));
    }

    [Theory]
    [InlineData(SpreadsheetConditionalFormattingOperator.Equal, "equal")]
    [InlineData(SpreadsheetConditionalFormattingOperator.NotEqual, "notEqual")]
    [InlineData(SpreadsheetConditionalFormattingOperator.GreaterThan, "greaterThan")]
    [InlineData(SpreadsheetConditionalFormattingOperator.GreaterThanOrEqual, "greaterThanOrEqual")]
    [InlineData(SpreadsheetConditionalFormattingOperator.LessThan, "lessThan")]
    [InlineData(SpreadsheetConditionalFormattingOperator.LessThanOrEqual, "lessThanOrEqual")]
    public void Every_one_bound_comparison_is_written_as_Excels_operator_with_one_formula(
        SpreadsheetConditionalFormattingOperator comparison,
        string expected)
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddCellValue("A1:A9", comparison).Formula.Value = 5);

        ConditionalFormattingRule rule = Rules(written).Single();

        Assert.Equal(expected, rule.Operator!.InnerText);
        Assert.Equal(["5"], Formulas(rule));
    }

    [Theory]
    [InlineData(SpreadsheetConditionalFormattingOperator.Between, "between")]
    [InlineData(SpreadsheetConditionalFormattingOperator.NotBetween, "notBetween")]
    public void A_range_comparison_writes_both_bounds_lower_first(
        SpreadsheetConditionalFormattingOperator comparison,
        string expected)
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetCellValueConditionalFormattingRule rule =
                Sheet(package).ConditionalFormatting.AddCellValue("A1:A9", comparison);
            rule.Formula.Value = 10;
            rule.Formula2.Value = 20;
        });

        ConditionalFormattingRule rule = Rules(written).Single();

        Assert.Equal(expected, rule.Operator!.InnerText);
        Assert.Equal(["10", "20"], Formulas(rule));
    }

    [Fact]
    public void A_text_value_is_quoted_and_its_own_quotes_doubled()
    {
        // Unquoted, Overdue is a defined name Excel cannot find; with a quote left single, the
        // string ends early and the rest of the formula is a syntax error.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal)
                .Formula.Value = "Say \"overdue\"");

        Assert.Equal(["\"Say \"\"overdue\"\"\""], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void A_character_value_is_text_like_a_one_letter_string()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal)
                .Formula.Value = 'Y');

        Assert.Equal(["\"Y\""], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void A_date_value_is_compared_as_its_serial_number()
    {
        // A date in a cell is a serial number with a format on top, so that is what it has to be
        // compared with.
        DateTime deadline = new(2026, 8, 27);

        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.LessThan)
                .Formula.Value = deadline);

        Assert.Equal(
            [deadline.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture)],
            Formulas(Rules(written).Single()));
    }

    [Theory]
    [InlineData(true, "TRUE")]
    [InlineData(false, "FALSE")]
    public void A_boolean_value_is_Excels_TRUE_or_FALSE(bool value, string expected)
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal)
                .Formula.Value = value);

        Assert.Equal([expected], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void A_decimal_value_keeps_its_digits()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.GreaterThan)
                .Formula.Value = 1234.50M);

        Assert.Equal(["1234.50"], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void A_value_of_a_type_the_writer_does_not_know_is_compared_as_its_text()
    {
        // The same fallback a cell gets: Convert.ToString, invariant, and then it is text.
        Guid id = Guid.Parse("6f9619ff-8b86-d011-b42d-00c04fc964ff");

        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal)
                .Formula.Value = id);

        Assert.Equal(["\"6f9619ff-8b86-d011-b42d-00c04fc964ff\""], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void A_formula_operand_is_written_as_given_without_its_equals_sign()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.GreaterThan)
                .Formula.ExcelFormula = "=$B$1");

        Assert.Equal(["$B$1"], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void Text_that_looks_like_a_formula_is_still_text_when_it_is_a_value()
    {
        // The only way to get a formula is to ask for one. A value is a value even when it starts
        // with an equals sign, which is the rule a cell follows too.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting
                .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal)
                .Formula.Value = "=A1");

        Assert.Equal(["\"=A1\""], Formulas(Rules(written).Single()));
    }

    [Fact]
    public void Setting_a_value_clears_the_formula_and_setting_a_formula_clears_the_value()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetConditionalFormattingFormula formula = Sheet(package).ConditionalFormatting
            .AddCellValue("A1", SpreadsheetConditionalFormattingOperator.Equal)
            .Formula;

        formula.ExcelFormula = "$B$1";
        formula.Value = 3;
        Assert.Null(formula.ExcelFormula);
        Assert.Equal(3, formula.Value);

        formula.ExcelFormula = "$B$1";
        Assert.Null(formula.Value);
        Assert.Equal("$B$1", formula.ExcelFormula);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void A_value_with_no_numeric_spelling_is_refused_on_the_line_that_set_it(double value)
    {
        using SpreadsheetPackage package = new();
        SpreadsheetCellValueConditionalFormattingRule rule = Sheet(package).ConditionalFormatting
            .AddCellValue("A1", SpreadsheetConditionalFormattingOperator.Equal);

        Assert.Throws<ArgumentOutOfRangeException>(() => rule.Formula.Value = value);
        Assert.Throws<ArgumentOutOfRangeException>(() => rule.Formula.Value = (float)value);
    }

    [Fact]
    public void A_rule_with_nothing_to_compare_with_is_refused_before_anything_is_written()
    {
        using SpreadsheetPackage package = new();
        Sheet(package).ConditionalFormatting.AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Equal);

        InvalidOperationException refused = Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());
        Assert.Contains("A1:A9", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_null_or_empty_operand_counts_as_nothing_to_compare_with()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetCellValueConditionalFormattingRule rule = Sheet(package).ConditionalFormatting
            .AddCellValue("A1", SpreadsheetConditionalFormattingOperator.Equal);
        rule.Formula.ExcelFormula = " = ";

        Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());
    }

    [Fact]
    public void A_between_rule_without_an_upper_bound_is_refused()
    {
        using SpreadsheetPackage package = new();
        Sheet(package).ConditionalFormatting
            .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Between)
            .Formula.Value = 1;

        InvalidOperationException refused = Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());
        Assert.Contains("Formula2", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_second_bound_on_a_one_bound_comparison_is_refused_rather_than_dropped()
    {
        // Whoever set it thought it meant something. Dropping it would write a rule that does
        // something other than what its author reads it as doing.
        using SpreadsheetPackage package = new();
        SpreadsheetCellValueConditionalFormattingRule rule = Sheet(package).ConditionalFormatting
            .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.GreaterThan);
        rule.Formula.Value = 1;
        rule.Formula2.Value = 9;

        Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());
    }

    [Fact]
    public void A_refused_rule_can_be_fixed_and_the_workbook_written_afterwards()
    {
        // The refusal comes before any sheet is committed. Committing adds table parts, and a
        // second attempt after a half-finished first would add every one of them twice.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet first = Sheet(package);
        first.Cells[1, 1].Value = "Name";
        first.Cells[2, 1].Value = "Ada";
        first.Tables.Add(first.Cells["A1:A2"], "People");

        SpreadsheetWorksheet second = package.Workbook.Worksheets.Add("Second");
        SpreadsheetCellValueConditionalFormattingRule rule = second.ConditionalFormatting
            .AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Between);
        rule.Formula.Value = 1;

        Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());

        rule.Formula2.Value = 9;
        using WrittenWorkbook written = WrittenWorkbook.Opening(package.GetAsByteArray());

        Assert.Single(written.WorkbookPart.WorksheetParts.SelectMany(part => part.TableDefinitionParts));
        Assert.Equal(["1", "9"], Formulas(Rules(written, "Second").Single()));
        Assert.Empty(written.SchemaErrors());
    }

    [Fact]
    public void An_operator_Excel_does_not_have_is_refused()
    {
        using SpreadsheetPackage package = new();

        Assert.Throws<ArgumentOutOfRangeException>(() => Sheet(package).ConditionalFormatting
            .AddCellValue("A1", (SpreadsheetConditionalFormattingOperator)99));
    }

    // ---------------------------------------------------------------- formula rules

    [Fact]
    public void A_formula_rule_is_an_expression_carrying_its_formula_without_the_equals_sign()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetFormulaConditionalFormattingRule rule = Sheet(package).ConditionalFormatting.AddFormula("A2:F100");
            rule.Formula.ExcelFormula = "=$D2<TODAY()";
            rule.Style.Font.SetBold(true);
        });

        ConditionalFormattingRule rule = Rules(written).Single();

        Assert.Equal(ConditionalFormatValues.Expression, rule.Type!.Value);
        Assert.Null(rule.Operator);
        Assert.Equal(["$D2<TODAY()"], Formulas(rule));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("=")]
    [InlineData("   ")]
    public void A_formula_rule_with_no_formula_is_refused(string? formula)
    {
        using SpreadsheetPackage package = new();
        Sheet(package).ConditionalFormatting.AddFormula("A1:A9").Formula.ExcelFormula = formula;

        Assert.Throws<InvalidOperationException>(() => package.GetAsByteArray());
    }

    [Fact]
    public void A_formula_is_cleaned_of_anything_XML_cannot_carry()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddFormula("A1:A9").Formula.ExcelFormula = "A1>" + (char)1 + "0");

        Assert.Equal(["A1>0"], Formulas(Rules(written).Single()));
    }

    // ---------------------------------------------------------------- duplicate and unique values

    [Fact]
    public void Duplicate_values_are_a_rule_of_their_own_with_nothing_to_compare_with()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddDuplicateValues("A1:A9").Style.Fill.BackgroundColor.SetColor("FFC7CE"));

        ConditionalFormattingRule rule = Rules(written).Single();

        Assert.Equal(ConditionalFormatValues.DuplicateValues, rule.Type!.Value);
        Assert.Empty(Formulas(rule));
        Assert.Null(rule.Operator);
        Assert.NotNull(rule.FormatId);
    }

    [Fact]
    public void Unique_values_are_a_rule_of_their_own_with_nothing_to_compare_with()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddUniqueValues("A1:A9").Style.Font.SetItalic(true));

        ConditionalFormattingRule rule = Rules(written).Single();

        Assert.Equal(ConditionalFormatValues.UniqueValues, rule.Type!.Value);
        Assert.Empty(Formulas(rule));
        Assert.NotNull(rule.FormatId);
    }

    // ---------------------------------------------------------------- colour scales

    [Fact]
    public void A_two_colour_scale_runs_from_the_lowest_value_to_the_highest()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddTwoColorScale("B2:B20"));

        ConditionalFormattingRule rule = Rules(written).Single();
        ColorScale scale = rule.GetFirstChild<ColorScale>()!;

        Assert.Equal(ConditionalFormatValues.ColorScale, rule.Type!.Value);
        Assert.Equal(["min", "max"], Anchors(scale).Select(anchor => anchor.Type));
        Assert.All(Anchors(scale), anchor => Assert.Null(anchor.Value));
        Assert.Equal(["FFF8696B", "FF63BE7B"], Colors(scale));
    }

    [Fact]
    public void A_three_colour_scale_passes_through_the_fiftieth_percentile()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddThreeColorScale("B2:B20"));

        ColorScale scale = Rules(written).Single().GetFirstChild<ColorScale>()!;

        Assert.Equal<(string, string?)[]>(
            [("min", null), ("percentile", "50"), ("max", null)],
            [.. Anchors(scale)]);
        Assert.Equal(["FFF8696B", "FFFFEB84", "FF63BE7B"], Colors(scale));
    }

    [Fact]
    public void A_colour_scale_lists_every_anchor_before_any_colour()
    {
        // CT_ColorScale pairs them by position. Interleaved, the validator rejects the rule.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddThreeColorScale("B2:B20"));

        ColorScale scale = Rules(written).Single().GetFirstChild<ColorScale>()!;

        Assert.Equal(
            ["cfvo", "cfvo", "cfvo", "color", "color", "color"],
            scale.ChildElements.Select(element => element.LocalName));
    }

    [Fact]
    public void A_two_colour_scale_has_no_middle_point()
    {
        using SpreadsheetPackage package = new();

        Assert.Null(Sheet(package).ConditionalFormatting.AddTwoColorScale("A1:A9").MiddleValue);
    }

    [Fact]
    public void Every_point_of_a_colour_scale_can_be_re_anchored_and_recoloured()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetColorScaleConditionalFormattingRule rule =
                Sheet(package).ConditionalFormatting.AddThreeColorScale("B2:B20");
            rule.LowValue.SetNumber(-2.5D);
            rule.LowValue.Color.SetColor(Drawing.Color.FromArgb(255, 0, 0, 255));
            rule.MiddleValue!.SetPercent(25D);
            rule.MiddleValue.Color.SetColor("#FFFFFF");
            rule.HighValue.SetFormula("=MAX($B$2:$B$20)*0.9");
            rule.HighValue.Color.SetColor("80FF0000");
        });

        ColorScale scale = Rules(written).Single().GetFirstChild<ColorScale>()!;

        Assert.Equal<(string, string?)[]>(
            [("num", "-2.5"), ("percent", "25"), ("formula", "MAX($B$2:$B$20)*0.9")],
            [.. Anchors(scale)]);
        Assert.Equal(["FF0000FF", "FFFFFFFF", "80FF0000"], Colors(scale));
    }

    [Fact]
    public void A_colour_scale_point_can_be_anchored_to_a_percentile_or_back_to_the_extremes()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetColorScaleConditionalFormattingRule rule =
                Sheet(package).ConditionalFormatting.AddTwoColorScale("B2:B20");
            rule.LowValue.SetPercentile(10D);
            rule.HighValue.SetNumber(5D);
            rule.HighValue.SetHighestValue();
        });

        ColorScale scale = Rules(written).Single().GetFirstChild<ColorScale>()!;

        Assert.Equal<(string, string?)[]>([("percentile", "10"), ("max", null)], [.. Anchors(scale)]);
    }

    [Fact]
    public void A_colour_scale_point_reports_how_it_is_anchored()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetColorScaleValue point = Sheet(package).ConditionalFormatting.AddTwoColorScale("A1:A9").LowValue;

        Assert.Equal(SpreadsheetColorScaleValueType.LowestValue, point.Type);
        Assert.Null(point.Value);

        point.SetPercentile(12.5D);

        Assert.Equal(SpreadsheetColorScaleValueType.Percentile, point.Type);
        Assert.Equal("12.5", point.Value);

        point.SetLowestValue();

        Assert.Equal(SpreadsheetColorScaleValueType.LowestValue, point.Type);
        Assert.Null(point.Value);
    }

    [Theory]
    [InlineData(-0.1D)]
    [InlineData(100.1D)]
    [InlineData(double.NaN)]
    public void A_percentage_or_percentile_outside_nought_to_a_hundred_is_refused(double value)
    {
        using SpreadsheetPackage package = new();
        SpreadsheetColorScaleValue point = Sheet(package).ConditionalFormatting.AddTwoColorScale("A1:A9").LowValue;

        Assert.Throws<ArgumentOutOfRangeException>(() => point.SetPercent(value));
        Assert.Throws<ArgumentOutOfRangeException>(() => point.SetPercentile(value));
        Assert.Equal(SpreadsheetColorScaleValueType.LowestValue, point.Type);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_colour_scale_number_with_no_numeric_spelling_is_refused(double value)
    {
        using SpreadsheetPackage package = new();
        SpreadsheetColorScaleValue point = Sheet(package).ConditionalFormatting.AddTwoColorScale("A1:A9").LowValue;

        Assert.Throws<ArgumentOutOfRangeException>(() => point.SetNumber(value));
    }

    [Fact]
    public void An_empty_colour_scale_formula_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetColorScaleValue point = Sheet(package).ConditionalFormatting.AddTwoColorScale("A1:A9").LowValue;

        Assert.Throws<ArgumentNullException>(() => point.SetFormula(null!));
        Assert.Throws<ArgumentException>(() => point.SetFormula("="));
    }

    [Fact]
    public void A_colour_scale_adds_nothing_to_the_differential_format_table()
    {
        // The colour is computed per cell, so there is no fixed format to point at.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddThreeColorScale("A1:A9"));

        Assert.Null(Rules(written).Single().FormatId);
        Assert.Empty(DifferentialFormats(written));
    }

    // ---------------------------------------------------------------- what a dxf carries

    [Fact]
    public void Font_changes_go_into_the_differential_format_in_schema_order()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingFontStyle font = Sheet(package).ConditionalFormatting
                .AddDuplicateValues("A1:A9").Style.Font;
            font.Color.SetColor("9C0006");
            font.SetUnderline(true);
            font.SetStrikethrough(true);
            font.SetItalic(true);
            font.SetBold(true);
        });

        Font font = DxfOf(written, Rules(written).Single()).Font!;

        Assert.Equal(["b", "i", "strike", "u", "color"], font.ChildElements.Select(element => element.LocalName));
        Assert.Equal("FF9C0006", font.GetFirstChild<Color>()!.Rgb!.Value);
        Assert.Null(font.GetFirstChild<FontSize>());
        Assert.Null(font.GetFirstChild<FontName>());
    }

    [Fact]
    public void Turning_bold_off_is_written_rather_than_left_out()
    {
        // A dxf is a patch. Leaving b out leaves a bold heading bold; b val="0" is what takes it
        // off.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingFontStyle font = Sheet(package).ConditionalFormatting
                .AddUniqueValues("A1:A9").Style.Font;
            font.SetBold(false);
            font.SetItalic(false);
            font.SetStrikethrough(false);
            font.SetUnderline(false);
        });

        Font font = DxfOf(written, Rules(written).Single()).Font!;

        Assert.False(font.GetFirstChild<Bold>()!.Val!.Value);
        Assert.False(font.GetFirstChild<Italic>()!.Val!.Value);
        Assert.False(font.GetFirstChild<Strike>()!.Val!.Value);
        Assert.Equal(UnderlineValues.None, font.GetFirstChild<Underline>()!.Val!.Value);
    }

    [Fact]
    public void A_conditional_fill_puts_its_colour_where_Excel_reads_it()
    {
        // In a cell format the colour of a solid fill is fgColor. In a dxf Excel reads bgColor and
        // writes no pattern type, and a dxf shaped like a cell format shows no fill at all.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddDuplicateValues("A1:A9").Style.Fill.BackgroundColor.SetColor("FFC7CE"));

        PatternFill fill = DxfOf(written, Rules(written).Single()).Fill!.PatternFill!;

        Assert.Null(fill.PatternType);
        Assert.Null(fill.ForegroundColor);
        Assert.Equal("FFFFC7CE", fill.BackgroundColor!.Rgb!.Value);
    }

    [Fact]
    public void A_border_around_outlines_every_side_of_each_matching_cell()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddDuplicateValues("A1:A9").Style.Border
                .BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.Red));

        Border border = DxfOf(written, Rules(written).Single()).Border!;

        Assert.Equal(["left", "right", "top", "bottom"], border.ChildElements.Select(element => element.LocalName));
        Assert.All(
            border.ChildElements.Cast<BorderPropertiesType>(),
            side =>
            {
                Assert.Equal(BorderStyleValues.Thin, side.Style!.Value);
                Assert.Equal("FFFF0000", side.Color!.Rgb!.Value);
            });
    }

    [Fact]
    public void Only_the_border_sides_that_were_set_are_written()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingBorderStyle border = Sheet(package).ConditionalFormatting
                .AddDuplicateValues("A1:A9").Style.Border;
            border.SetBottom(SpreadsheetBorderStyle.Medium, Drawing.Color.Black);
            border.SetLeft(SpreadsheetBorderStyle.None, Drawing.Color.Black);
        });

        Border border = DxfOf(written, Rules(written).Single()).Border!;

        Assert.Equal(["left", "bottom"], border.ChildElements.Select(element => element.LocalName));
        Assert.Equal(BorderStyleValues.Medium, border.BottomBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.None, border.LeftBorder!.Style!.Value);
        Assert.Null(border.LeftBorder.Color);
    }

    [Fact]
    public void Each_border_side_can_be_set_on_its_own()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingBorderStyle border = Sheet(package).ConditionalFormatting
                .AddDuplicateValues("A1:A9").Style.Border;
            border.SetTop(SpreadsheetBorderStyle.Thin, Drawing.Color.Blue);
            border.SetRight(SpreadsheetBorderStyle.Thin, Drawing.Color.Green);
        });

        Border border = DxfOf(written, Rules(written).Single()).Border!;

        Assert.Equal(["right", "top"], border.ChildElements.Select(element => element.LocalName));
        Assert.Equal("FF0000FF", border.TopBorder!.Color!.Rgb!.Value);
        Assert.Equal("FF008000", border.RightBorder!.Color!.Rgb!.Value);
    }

    [Fact]
    public void A_conditional_number_format_is_numbered_from_the_same_sequence_as_cell_formats()
    {
        // Two sequences would hand out 164 twice, and one id would then name two formats.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[1, 1].Value = 0.5D;
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("0.00");
            sheet.ConditionalFormatting.AddDuplicateValues("A1:A9").Style.NumberFormat.SetFormat("0.0%");
        });

        NumberingFormat dxfFormat = DxfOf(written, Rules(written).Single()).NumberingFormat!;
        NumberingFormat[] registered = [.. written.Stylesheet.NumberingFormats!.Elements<NumberingFormat>()];
        uint cellFormatId = written.FormatOf(written.Cell("Data", "A1")!).NumberFormatId!.Value;

        Assert.Equal("0.0%", dxfFormat.FormatCode!.Value);
        Assert.True(dxfFormat.NumberFormatId!.Value >= 164U);
        Assert.NotEqual(cellFormatId, dxfFormat.NumberFormatId!.Value);
        Assert.Contains(registered, format =>
            format.NumberFormatId!.Value == dxfFormat.NumberFormatId!.Value && format.FormatCode!.Value == "0.0%");
    }

    [Fact]
    public void A_number_format_a_cell_already_uses_is_not_registered_twice()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[1, 1].Value = 0.5D;
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("0.00");
            sheet.ConditionalFormatting.AddDuplicateValues("A1:A9").Style.NumberFormat.SetFormat("0.00");
        });

        NumberingFormat dxfFormat = DxfOf(written, Rules(written).Single()).NumberingFormat!;

        Assert.Single(written.Stylesheet.NumberingFormats!.Elements<NumberingFormat>());
        Assert.Equal(
            written.FormatOf(written.Cell("Data", "A1")!).NumberFormatId!.Value,
            dxfFormat.NumberFormatId!.Value);
    }

    [Fact]
    public void An_empty_conditional_number_format_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetConditionalFormattingNumberFormatStyle numberFormat = Sheet(package).ConditionalFormatting
            .AddDuplicateValues("A1:A9").Style.NumberFormat;

        Assert.Throws<ArgumentNullException>(() => numberFormat.SetFormat(null!));
        Assert.Throws<ArgumentException>(() => numberFormat.SetFormat(""));
        Assert.Throws<ArgumentException>(() => numberFormat.SetFormat("   "));
    }

    [Fact]
    public void A_differential_format_keeps_its_parts_in_schema_order()
    {
        // CT_Dxf is a sequence: font, numFmt, fill, border. Set in any other order, they are
        // still written in that one.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingStyle style = Sheet(package).ConditionalFormatting
                .AddDuplicateValues("A1:A9").Style;
            style.Border.BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.Black);
            style.Fill.BackgroundColor.SetColor("FFEB9C");
            style.NumberFormat.SetFormat("0.00");
            style.Font.SetBold(true);
        });

        Assert.Equal(
            ["font", "numFmt", "fill", "border"],
            DxfOf(written, Rules(written).Single()).ChildElements.Select(element => element.LocalName));
    }

    [Fact]
    public void A_rule_that_formats_nothing_points_at_no_differential_format()
    {
        // Legitimate with Stop If True — that is how lower rules are kept off some cells — and
        // an empty dxf would cost an entry to say nothing.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddDuplicateValues("A1:A9").StopIfTrue = true);

        Assert.Null(Rules(written).Single().FormatId);
        Assert.Empty(DifferentialFormats(written));
    }

    [Fact]
    public void Stop_if_true_is_written_only_when_asked_for()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            rules.AddDuplicateValues("A1:A9").StopIfTrue = true;
            rules.AddUniqueValues("A1:A9");
        });

        ConditionalFormattingRule[] rules = [.. Rules(written)];

        Assert.True(rules[0].StopIfTrue!.Value);
        Assert.Null(rules[1].StopIfTrue);
    }

    // ---------------------------------------------------------------- dxf de-duplication

    [Fact]
    public void Two_rules_formatted_the_same_share_one_differential_format()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            Highlight(rules.AddDuplicateValues("A1:A9").Style);
            SpreadsheetCellValueConditionalFormattingRule overdue =
                rules.AddCellValue("B1:B9", SpreadsheetConditionalFormattingOperator.LessThan);
            overdue.Formula.Value = 0;
            Highlight(overdue.Style);
        });

        ConditionalFormattingRule[] rules = [.. Rules(written)];

        Assert.Single(DifferentialFormats(written));
        Assert.Equal(rules[0].FormatId!.Value, rules[1].FormatId!.Value);
    }

    [Fact]
    public void Two_rules_formatted_differently_do_not()
    {
        // The negative control. A repository that returned 0 for everything would pass the test
        // above.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            rules.AddDuplicateValues("A1:A9").Style.Fill.BackgroundColor.SetColor("FFC7CE");
            rules.AddUniqueValues("A1:A9").Style.Fill.BackgroundColor.SetColor("C6EFCE");
        });

        ConditionalFormattingRule[] rules = [.. Rules(written)];

        Assert.Equal(2, DifferentialFormats(written).Count);
        Assert.NotEqual(rules[0].FormatId!.Value, rules[1].FormatId!.Value);
    }

    [Fact]
    public void The_order_formatting_was_set_in_does_not_make_a_second_entry()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;

            SpreadsheetConditionalFormattingStyle first = rules.AddDuplicateValues("A1:A9").Style;
            first.Font.SetBold(true);
            first.Fill.BackgroundColor.SetColor("FFC7CE");

            SpreadsheetConditionalFormattingStyle second = rules.AddUniqueValues("B1:B9").Style;
            second.Fill.BackgroundColor.SetColor("FFC7CE");
            second.Font.SetBold(true);
        });

        Assert.Single(DifferentialFormats(written));
    }

    [Fact]
    public void A_property_set_to_false_is_a_different_format_from_one_never_set()
    {
        // Bold off and bold untouched are different instructions; sharing an entry would turn one
        // into the other.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;

            SpreadsheetConditionalFormattingStyle first = rules.AddDuplicateValues("A1:A9").Style;
            first.Fill.BackgroundColor.SetColor("FFC7CE");

            SpreadsheetConditionalFormattingStyle second = rules.AddUniqueValues("B1:B9").Style;
            second.Fill.BackgroundColor.SetColor("FFC7CE");
            second.Font.SetBold(false);
        });

        Assert.Equal(2, DifferentialFormats(written).Count);
    }

    [Fact]
    public void Rules_on_different_sheets_share_differential_formats_too()
    {
        // The style table belongs to the workbook, not to a sheet.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            foreach (string name in new[] { "North", "South", "East" })
            {
                SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add(name);
                sheet.Cells[1, 1].Value = 1;
                Highlight(sheet.ConditionalFormatting.AddDuplicateValues("A1:A9").Style);
            }
        });

        Assert.Single(DifferentialFormats(written));
        Assert.Equal(
            [0U, 0U, 0U],
            new[] { "North", "South", "East" }.Select(name => Rules(written, name).Single().FormatId!.Value));
    }

    [Fact]
    public void The_differential_format_count_matches_its_entries()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            rules.AddDuplicateValues("A1:A9").Style.Font.SetBold(true);
            rules.AddUniqueValues("A1:A9").Style.Font.SetItalic(true);
            rules.AddUniqueValues("B1:B9").Style.Font.SetItalic(true);
        });

        DifferentialFormats formats = written.Stylesheet.DifferentialFormats!;

        Assert.Equal(2U, formats.Count!.Value);
        Assert.Equal(2, formats.Elements<DifferentialFormat>().Count());
    }

    [Fact]
    public void Differential_formats_leave_the_cell_style_table_as_it_was()
    {
        // A dxf has its own table. Writing one must not move a cell's style index or add a
        // cellXfs entry nobody asked for.
        static void Build(SpreadsheetPackage package, bool withRules)
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[1, 1].Value = "heading";
            sheet.Cells[1, 1].Style.Font.SetBold(true);
            sheet.Cells[2, 1].Value = 1.5D;
            sheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor("EEEEEE");
            if (withRules)
            {
                SpreadsheetConditionalFormattingStyle style = sheet.ConditionalFormatting.AddDuplicateValues("A1:A9").Style;
                style.Font.SetBold(true);
                style.Fill.BackgroundColor.SetColor("EEEEEE");
                style.Border.BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.Black);
            }
        }

        using WrittenWorkbook without = WrittenWorkbook.From(package => Build(package, withRules: false));
        using WrittenWorkbook with = WrittenWorkbook.From(package => Build(package, withRules: true));

        Assert.Equal(without.Stylesheet.CellFormats!.OuterXml, with.Stylesheet.CellFormats!.OuterXml);
        Assert.Equal(without.Stylesheet.Fonts!.OuterXml, with.Stylesheet.Fonts!.OuterXml);
        Assert.Equal(without.Stylesheet.Fills!.OuterXml, with.Stylesheet.Fills!.OuterXml);
        Assert.Equal(without.Stylesheet.Borders!.OuterXml, with.Stylesheet.Borders!.OuterXml);
        Assert.Equal(
            without.Cell("Data", "A2")!.StyleIndex!.Value,
            with.Cell("Data", "A2")!.StyleIndex!.Value);
    }

    // ---------------------------------------------------------------- priority, grouping, addresses

    [Fact]
    public void Rules_take_priority_in_the_order_they_were_added_across_the_whole_sheet()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            rules.AddDuplicateValues("A1:A9");
            rules.AddTwoColorScale("B1:B9");
            rules.AddUniqueValues("A1:A9");
            rules.AddFormula("C1:C9").Formula.ExcelFormula = "C1>0";
        });

        // Grouped by address, the document order is A1:A9's two rules, then B1:B9's, then
        // C1:C9's — but the priorities still record the order the rules were added in.
        Assert.Equal([1, 3, 2, 4], Rules(written).Select(rule => rule.Priority!.Value));
        Assert.Equal(
            [ConditionalFormatValues.DuplicateValues, ConditionalFormatValues.ColorScale,
             ConditionalFormatValues.UniqueValues, ConditionalFormatValues.Expression],
            Rules(written).OrderBy(rule => rule.Priority!.Value).Select(rule => rule.Type!.Value));
    }

    [Fact]
    public void Rules_on_the_same_address_share_one_element_and_others_get_their_own()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
            rules.AddDuplicateValues("A1:A9");
            rules.AddTwoColorScale("B1:B9");
            rules.AddUniqueValues("$A$1:$A$9");
        });

        ConditionalFormatting[] blocks = [.. Blocks(written)];

        Assert.Equal(["A1:A9", "B1:B9"], blocks.Select(block => block.SequenceOfReferences!.InnerText));
        Assert.Equal(2, blocks[0].Elements<ConditionalFormattingRule>().Count());
        Assert.Single(blocks[1].Elements<ConditionalFormattingRule>());
    }

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("$B$2:$B$9", "B2:B9")]
    [InlineData("Sheet1!A1:C3", "A1:C3")]
    [InlineData("'Some Sheet'!$A$1:$C$3", "A1:C3")]
    [InlineData("b2:b9   e2:e9", "B2:B9 E2:E9")]
    [InlineData("C:C", "C:C")]
    [InlineData("2:4", "2:4")]
    public void A_conditional_format_address_is_normalised_rather_than_copied(string given, string expected)
    {
        using SpreadsheetPackage package = new();

        Assert.Equal(expected, Sheet(package).ConditionalFormatting.AddDuplicateValues(given).Address);
    }

    [Fact]
    public void A_rule_can_cover_more_than_one_rectangle()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).ConditionalFormatting.AddDuplicateValues("B1:B9 E1:E9"));

        Assert.Equal("B1:B9 E1:E9", Blocks(written).Single().SequenceOfReferences!.InnerText);
    }

    [Theory]
    [InlineData("not an address")]
    [InlineData("A0")]
    [InlineData("C1:A1")]
    [InlineData("   ")]
    public void A_conditional_format_address_that_is_not_one_fails_where_it_was_written(string address)
    {
        // The same refusals every other address in the library gets: malformed text is a
        // FormatException, and a well-formed address outside the grid an ArgumentException.
        using SpreadsheetPackage package = new();

        Exception refused = Assert.ThrowsAny<Exception>(() => Sheet(package).ConditionalFormatting.AddDuplicateValues(address));
        Assert.True(refused is ArgumentException or FormatException, refused.GetType().Name);
    }

    [Fact]
    public void A_bad_address_adds_no_rule()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = Sheet(package);

        Assert.ThrowsAny<Exception>(() => sheet.ConditionalFormatting.AddDuplicateValues("A0"));
        Assert.Empty(sheet.ConditionalFormatting);
    }

    [Fact]
    public void A_null_address_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;

        Assert.Throws<ArgumentNullException>(() => rules.AddCellValue(null!, SpreadsheetConditionalFormattingOperator.Equal));
        Assert.Throws<ArgumentNullException>(() => rules.AddFormula(null!));
        Assert.Throws<ArgumentNullException>(() => rules.AddDuplicateValues(null!));
        Assert.Throws<ArgumentNullException>(() => rules.AddUniqueValues(null!));
        Assert.Throws<ArgumentNullException>(() => rules.AddTwoColorScale(null!));
        Assert.Throws<ArgumentNullException>(() => rules.AddThreeColorScale(null!));
    }

    [Fact]
    public void Every_kind_of_rule_added_through_a_range_takes_the_ranges_address()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = Sheet(package);
        SpreadsheetRangeConditionalFormatting range = sheet.Cells["B2:D9"].ConditionalFormatting;

        range.AddCellValue(SpreadsheetConditionalFormattingOperator.Equal);
        range.AddFormula();
        range.AddDuplicateValues();
        range.AddUniqueValues();
        range.AddTwoColorScale();
        range.AddThreeColorScale();

        Assert.Equal(6, sheet.ConditionalFormatting.Count());
        Assert.All(sheet.ConditionalFormatting, rule => Assert.Equal("B2:D9", rule.Address));
    }

    [Fact]
    public void A_whole_column_can_be_formatted_through_a_range()
    {
        // A data validation through a range refuses this. A conditional format is one sqref
        // whatever it covers, and a column people keep appending to is the usual thing to format.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            Sheet(package).Cells["C:C"].ConditionalFormatting.AddThreeColorScale());

        Assert.Equal("C:C", Blocks(written).Single().SequenceOfReferences!.InnerText);
        Assert.Empty(written.SchemaErrors());
    }

    [Fact]
    public void The_rules_can_be_enumerated_in_priority_order_before_they_are_written()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetConditionalFormattingCollection rules = Sheet(package).ConditionalFormatting;
        rules.AddTwoColorScale("A1:A9");
        rules.AddCellValue("B1", SpreadsheetConditionalFormattingOperator.Equal);
        rules.AddFormula("C1");

        Assert.Equal(["A1:A9", "B1", "C1"], rules.Select(rule => rule.Address));
        Assert.IsType<SpreadsheetColorScaleConditionalFormattingRule>(rules.First());
        Assert.IsType<SpreadsheetFormulaConditionalFormattingRule>(rules.Last());
    }

    [Fact]
    public void A_sheet_with_no_rules_has_no_conditional_formatting_element()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package => Sheet(package));

        Assert.Empty(Blocks(written));
        Assert.Empty(DifferentialFormats(written));
    }

    // ---------------------------------------------------------------- the package around it

    [Fact]
    public void Adding_a_rule_after_the_workbook_was_written_is_refused()
    {
        // The rule would never reach the file, so the call that adds it is the one that should
        // say so.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = Sheet(package);
        package.GetAsByteArray();

        Assert.Throws<InvalidOperationException>(() => sheet.ConditionalFormatting.AddDuplicateValues("A1:A9"));
        Assert.Throws<InvalidOperationException>(() => sheet.Cells["A1:A9"].ConditionalFormatting.AddTwoColorScale());
    }

    [Fact]
    public void A_workbook_opened_for_reading_refuses_a_rule()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            Sheet(package);
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        Assert.Throws<InvalidOperationException>(() =>
            reopened.Workbook.Worksheets["Data"]!.ConditionalFormatting.AddDuplicateValues("A1:A9"));
    }

    [Fact]
    public void A_workbook_with_conditional_formats_reads_back_through_the_library()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[2, 1].Value = 42D;
            sheet.ConditionalFormatting.AddThreeColorScale("A1:A9");
            Highlight(sheet.ConditionalFormatting.AddDuplicateValues("A1:A9").Style);
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetWorksheet read = reopened.Workbook.Worksheets["Data"]!;

        Assert.Equal("x", read.Cells[1, 1].Value);
        Assert.Equal(42D, read.Cells[2, 1].Value);
    }

    // ---------------------------------------------------------------- schema

    [Fact]
    public void Conditional_formatting_sits_after_merged_cells_and_before_data_validation()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells["A4:B4"].Merge = true;
            sheet.DataValidations.AddAnyValidation("A2");
            sheet.ConditionalFormatting.AddDuplicateValues("A1:A9");
            sheet.ConditionalFormatting.AddTwoColorScale("B1:B9");
        });

        Assert.Equal<string[]>(
            ["dimension", "sheetData", "mergeCells", "conditionalFormatting", "conditionalFormatting", "dataValidations"],
            [.. written.ChildOrder("Data")]);
    }

    [Fact]
    public void Without_data_validation_conditional_formatting_still_precedes_a_drawing_and_tables()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[1, 1].Value = "Name";
            sheet.Cells[2, 1].Value = "Ada";
            sheet.Drawings.AddPicture("logo", new MemoryStream(TestImages.Png(20, 20)));
            sheet.Tables.Add(sheet.Cells["A1:A2"], "People");
            sheet.ConditionalFormatting.AddDuplicateValues("A2:A9");
        });

        Assert.Equal<string[]>(
            ["dimension", "sheetData", "conditionalFormatting", "drawing", "tableParts"],
            [.. written.ChildOrder("Data")]);
    }

    [Fact]
    public void Every_rule_type_and_every_kind_of_format_passes_the_schema_validator()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            SpreadsheetConditionalFormattingCollection rules = sheet.ConditionalFormatting;

            foreach (SpreadsheetConditionalFormattingOperator comparison in Enum.GetValues<SpreadsheetConditionalFormattingOperator>())
            {
                SpreadsheetCellValueConditionalFormattingRule rule = rules.AddCellValue("A1:A20", comparison);
                rule.Formula.Value = 10;
                if (comparison is SpreadsheetConditionalFormattingOperator.Between or SpreadsheetConditionalFormattingOperator.NotBetween)
                {
                    rule.Formula2.ExcelFormula = "$Z$1";
                }

                rule.Style.Font.SetBold(comparison is SpreadsheetConditionalFormattingOperator.Equal);
            }

            SpreadsheetCellValueConditionalFormattingRule text = rules.AddCellValue("B1:B20", SpreadsheetConditionalFormattingOperator.Equal);
            text.Formula.Value = "Overdue";
            text.StopIfTrue = true;
            SpreadsheetConditionalFormattingStyle everything = text.Style;
            everything.Font.SetBold(true);
            everything.Font.SetItalic(false);
            everything.Font.SetUnderline(true);
            everything.Font.SetStrikethrough(true);
            everything.Font.Color.SetColor("9C0006");
            everything.Fill.BackgroundColor.SetColor("FFC7CE");
            everything.Border.BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.DarkRed);
            everything.Border.SetBottom(SpreadsheetBorderStyle.Medium, Drawing.Color.Black);
            everything.NumberFormat.SetFormat("\"Due \"0");

            rules.AddCellValue("C1:C20", SpreadsheetConditionalFormattingOperator.LessThan).Formula.Value = new DateTime(2026, 1, 1);
            rules.AddFormula("A2:F20").Formula.ExcelFormula = "=AND($D2<TODAY(),$E2=\"\")";
            rules.AddDuplicateValues("D1:D20").Style.Fill.BackgroundColor.SetColor("FFEB9C");
            rules.AddUniqueValues("E1:E20 G1:G20").Style.Font.Color.SetColor(Drawing.Color.Green);
            rules.AddDuplicateValues("H1:H20").StopIfTrue = true;
            rules.AddTwoColorScale("I:I");

            SpreadsheetColorScaleConditionalFormattingRule scale = rules.AddThreeColorScale("J1:J20");
            scale.LowValue.SetNumber(0D);
            scale.MiddleValue!.SetPercent(50D);
            scale.HighValue.SetFormula("MAX($J$1:$J$20)");

            SpreadsheetColorScaleConditionalFormattingRule percentiles = rules.AddTwoColorScale("K1:K20");
            percentiles.LowValue.SetPercentile(5D);
            percentiles.HighValue.SetPercentile(95D);
        });

        Assert.Empty(written.SchemaErrors());
    }

    [Fact]
    public void Conditional_formatting_alongside_everything_else_a_sheet_can_hold_passes_the_schema_validator()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Invoices");
            sheet.SetCodeName("Invoices");
            sheet.DefaultColumnWidth = 12D;
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[1, 2].Value = "Due";
            sheet.Cells[1, 3].Value = "Amount";
            sheet.Cells["A1:C1"].Style.Font.SetBold(true);
            for (int row = 2; row <= 6; row++)
            {
                sheet.Cells[row, 1].Value = $"INV-{row}";
                sheet.Cells[row, 2].Value = new DateTime(2026, 9, row);
                sheet.Cells[row, 3].Value = row * 100.5D;
            }

            sheet.Column(3).Style.NumberFormat.SetFormat("#,##0.00");
            sheet.Column(1).SetWidth(20D);
            sheet.View.FreezePanes(2, 1);
            sheet.Cells["A8:C8"].Merge = true;
            sheet.DataValidations.AddDecimalValidation("C2:C6").Formula.Value = 0M;
            sheet.Drawings.AddPicture("logo", new MemoryStream(TestImages.Png(120, 40)));
            sheet.Tables.Add(sheet.Cells["A1:C6"], "Invoices");
            sheet.Protect(lockEverything: false, password: "not-a-secret");

            // The same cells carry a cell style, a validation, a table and two conditional
            // formats at once.
            SpreadsheetCellValueConditionalFormattingRule large =
                sheet.Cells["C2:C6"].ConditionalFormatting.AddCellValue(SpreadsheetConditionalFormattingOperator.GreaterThanOrEqual);
            large.Formula.Value = 500;
            large.Style.NumberFormat.SetFormat("#,##0.00 \"!\"");
            sheet.Cells["C2:C6"].ConditionalFormatting.AddTwoColorScale();
            sheet.Cells["A2:C6"].ConditionalFormatting.AddFormula().Formula.ExcelFormula = "$B2<DATE(2026,9,4)";
            sheet.Cells["A2:A6"].ConditionalFormatting.AddDuplicateValues().Style.Font.SetBold(false);

            SpreadsheetWorksheet second = package.Workbook.Worksheets.Add("Summary");
            second.Cells[1, 1].Value = 1;
            Highlight(second.ConditionalFormatting.AddUniqueValues("A1:A9").Style);
        });

        Assert.Empty(written.SchemaErrors());
        Assert.Equal<string[]>(
            [
                "sheetPr",
                "dimension",
                "sheetViews",
                "sheetFormatPr",
                "cols",
                "sheetData",
                "sheetProtection",
                "mergeCells",
                "conditionalFormatting",
                "conditionalFormatting",
                "conditionalFormatting",
                "dataValidations",
                "drawing",
                "tableParts"
            ],
            [.. written.ChildOrder("Invoices")]);
    }

    [Fact]
    public void A_workbook_without_conditional_formats_still_passes_the_schema_validator()
    {
        // The control for the two above: an empty dxfs is in every stylesheet this library writes,
        // and it has to have been valid before anything was put in it.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = Sheet(package);
            sheet.Cells[1, 1].Style.Font.SetBold(true);
        });

        Assert.Empty(written.SchemaErrors());
    }

    // ---------------------------------------------------------------- helpers

    private static SpreadsheetWorksheet Sheet(SpreadsheetPackage package)
    {
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].Value = "x";
        return sheet;
    }

    private static void Highlight(SpreadsheetConditionalFormattingStyle style)
    {
        style.Font.Color.SetColor("9C0006");
        style.Fill.BackgroundColor.SetColor("FFC7CE");
    }

    private static IEnumerable<ConditionalFormatting> Blocks(WrittenWorkbook written, string sheet = "Data") =>
        written.Sheet(sheet).Elements<ConditionalFormatting>();

    private static IEnumerable<ConditionalFormattingRule> Rules(WrittenWorkbook written, string sheet = "Data") =>
        Blocks(written, sheet).SelectMany(block => block.Elements<ConditionalFormattingRule>());

    private static IReadOnlyList<DifferentialFormat> DifferentialFormats(WrittenWorkbook written) =>
        [.. written.Stylesheet.DifferentialFormats?.Elements<DifferentialFormat>() ?? []];

    private static DifferentialFormat DxfOf(WrittenWorkbook written, ConditionalFormattingRule rule) =>
        DifferentialFormats(written)[(int)rule.FormatId!.Value];

    private static string[] Formulas(ConditionalFormattingRule rule) =>
        [.. rule.Elements<Formula>().Select(formula => formula.Text)];

    private static IEnumerable<(string Type, string? Value)> Anchors(ColorScale scale) =>
        scale.Elements<ConditionalFormatValueObject>().Select(anchor => (anchor.Type!.InnerText ?? string.Empty, anchor.Val?.Value));

    private static string[] Colors(ColorScale scale) =>
        [.. scale.Elements<Color>().Select(color => color.Rgb!.Value!)];
}
