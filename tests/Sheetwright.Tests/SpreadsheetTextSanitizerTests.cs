using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The guard between arbitrary user text and an XML document.
/// </summary>
/// <remarks>
/// A cell value comes from somewhere — a form answer, an imported file, a name
/// somebody pasted — and XML 1.0 cannot carry every character .NET can hold. An
/// unescapable control character does not fail on write; it produces a package
/// that opens as "unreadable content", which is discovered by the person the
/// export was for.
/// </remarks>
public sealed class SpreadsheetTextSanitizerTests
{
    [Fact]
    public void Ordinary_text_is_returned_unchanged()
    {
        Assert.Equal("Ordinary", SpreadsheetTextSanitizer.Sanitize("Ordinary"));
    }

    [Fact]
    public void Text_that_needs_nothing_done_to_it_is_not_copied()
    {
        // The fast path matters: this runs on every cell of a large export.
        string original = "Ordinary";

        Assert.Same(original, SpreadsheetTextSanitizer.Sanitize(original));
    }

    [Theory]
    [InlineData("\u0000")]
    [InlineData("\u0001")]
    [InlineData("\u001F")]
    [InlineData("\uFFFE")]
    public void A_character_XML_cannot_carry_is_dropped(string forbidden)
    {
        Assert.Equal("ab", SpreadsheetTextSanitizer.Sanitize($"a{forbidden}b"));
    }

    [Theory]
    [InlineData("\t")]
    [InlineData("\n")]
    [InlineData("\r")]
    public void Whitespace_XML_can_carry_is_kept(string permitted)
    {
        Assert.Equal($"a{permitted}b", SpreadsheetTextSanitizer.Sanitize($"a{permitted}b"));
    }

    [Fact]
    public void An_emoji_survives_because_a_surrogate_pair_is_one_character()
    {
        // Each half of the pair fails IsXmlChar on its own. Testing the halves
        // separately is the whole point: the naive loop deletes every emoji.
        Assert.Equal("a\U0001F600b", SpreadsheetTextSanitizer.Sanitize("a\U0001F600b"));
    }

    [Fact]
    public void A_surrogate_with_no_partner_is_dropped()
    {
        Assert.Equal("ab", SpreadsheetTextSanitizer.Sanitize("a\uD83Db"));
    }

    [Fact]
    public void A_trailing_high_surrogate_is_dropped_rather_than_read_past_the_end()
    {
        Assert.Equal("a", SpreadsheetTextSanitizer.Sanitize("a\uD83D"));
    }

    [Fact]
    public void Null_stays_null_where_the_caller_allows_one()
    {
        Assert.Null(SpreadsheetTextSanitizer.SanitizeNullable(null));
        Assert.Equal("ab", SpreadsheetTextSanitizer.SanitizeNullable("a\u0000b"));
    }
}
