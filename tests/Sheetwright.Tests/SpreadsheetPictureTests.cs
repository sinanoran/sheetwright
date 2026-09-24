using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Sheetwright.Tests;

/// <summary>
/// Pictures: the size read out of a header, and the bytes read back out again.
/// </summary>
/// <remarks>
/// A drawing anchor needs a width and a height in EMUs before the image is ever
/// decoded, so the library parses five image headers itself. That is five small
/// binary readers, each with a fallback, and a wrong one does not throw — it
/// produces a logo squashed into a 180 by 60 box, which is a bug reported months
/// later by somebody who assumed it was meant to look like that.
/// </remarks>
public sealed class SpreadsheetPictureTests
{
    private const long EmusPerPixel = 9525L;

    public static TheoryData<string, byte[], int, int> Images() => new()
    {
        { "Png", TestImages.Png(120, 40), 120, 40 },
        { "Gif", TestImages.Gif(64, 32), 64, 32 },
        { "Bmp", TestImages.Bmp(200, 100), 200, 100 },
        { "Jpeg", TestImages.Jpeg(320, 240), 320, 240 },
        { "Tiff", TestImages.Tiff(300, 200), 300, 200 }
    };

    [Theory]
    [MemberData(nameof(Images))]
    public void An_image_is_anchored_at_the_size_its_header_declares(
        string format,
        byte[] image,
        int width,
        int height)
    {
        _ = format;

        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("logo", new MemoryStream(image));
        });

        Xdr.OneCellAnchor anchor = DrawingOf(written).Elements<Xdr.OneCellAnchor>().Single();

        Assert.Equal(width * EmusPerPixel, anchor.Extent!.Cx!.Value);
        Assert.Equal(height * EmusPerPixel, anchor.Extent!.Cy!.Value);
    }

    [Theory]
    [MemberData(nameof(Images))]
    public void An_image_comes_back_with_its_bytes_and_its_format(
        string format,
        byte[] image,
        int width,
        int height)
    {
        _ = (width, height);

        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("logo", new MemoryStream(image));
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetPicture picture = reopened.Workbook.Worksheets["Data"]!.Drawings.Single();

        Assert.Equal("logo", picture.Name);
        Assert.Equal(format, picture.Format);
        Assert.Equal(image, picture.GetData());
    }

    [Fact]
    public void The_bytes_handed_out_are_a_copy_and_not_the_original_array()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("logo", new MemoryStream(TestImages.Png(10, 10)));
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetPicture picture = reopened.Workbook.Worksheets["Data"]!.Drawings.Single();

        byte[] first = picture.GetData();
        first[0] = 0;

        Assert.NotEqual(first, picture.GetData());
    }

    [Fact]
    public void A_picture_reports_the_row_it_is_anchored_to()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("logo", new MemoryStream(TestImages.Png(10, 10)));
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        // Zero-based in the file, because that is what the anchor holds.
        Assert.Equal(0, reopened.Workbook.Worksheets["Data"]!.Drawings.Single().From.Row);
    }

    [Fact]
    public void Two_pictures_get_two_anchors_and_two_identifiers()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("first", new MemoryStream(TestImages.Png(10, 10)));
            sheet.Drawings.AddPicture("second", new MemoryStream(TestImages.Png(20, 20)));
        });

        Xdr.Picture[] pictures = [.. DrawingOf(written).Descendants<Xdr.Picture>()];
        uint[] ids = [.. pictures.Select(p => p.NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value)];

        Assert.Equal(2, pictures.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void Moving_a_picture_changes_its_offset_and_not_its_size()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            SpreadsheetPictureWriter picture = sheet.Drawings.AddPicture(
                "logo",
                new MemoryStream(TestImages.Png(120, 40)));
            picture.SetPosition(pixelTop: 5, pixelLeft: 7);
        });

        Xdr.OneCellAnchor anchor = DrawingOf(written).Elements<Xdr.OneCellAnchor>().Single();

        Assert.Equal("0", anchor.FromMarker!.ColumnId!.Text);
        Assert.Equal("0", anchor.FromMarker!.RowId!.Text);
        Assert.Equal((7 * EmusPerPixel).ToString(), anchor.FromMarker!.ColumnOffset!.Text);
        Assert.Equal((5 * EmusPerPixel).ToString(), anchor.FromMarker!.RowOffset!.Text);
        Assert.Equal(120 * EmusPerPixel, anchor.Extent!.Cx!.Value);
    }

    [Fact]
    public void A_stream_that_cannot_be_rewound_is_buffered_rather_than_refused()
    {
        // The header has to be read twice — once to identify the format, once to
        // size it — so a network stream would fail on the second read.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Drawings.AddPicture("logo", new ForwardOnlyStream(TestImages.Png(120, 40)));
        });

        Xdr.OneCellAnchor anchor = DrawingOf(written).Elements<Xdr.OneCellAnchor>().Single();

        Assert.Equal(120 * EmusPerPixel, anchor.Extent!.Cx!.Value);
    }

    [Fact]
    public void Something_that_is_not_an_image_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<InvalidDataException>(
            () => sheet.Drawings.AddPicture("logo", new MemoryStream([1, 2, 3, 4, 5, 6, 7, 8])));
    }

    [Fact]
    public void A_null_image_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => sheet.Drawings.AddPicture("logo", null!));
        Assert.Throws<ArgumentNullException>(() => sheet.Drawings.AddPicture(null!, new MemoryStream()));
    }

    [Fact]
    public void A_sheet_with_no_pictures_enumerates_to_nothing()
    {
        using SpreadsheetPackage package = new();

        Assert.Empty(package.Workbook.Worksheets.Add("Data").Drawings);
    }

    private static Xdr.WorksheetDrawing DrawingOf(WrittenWorkbook written)
    {
        WorksheetPart part = written.WorkbookPart.WorksheetParts.Single();
        return part.DrawingsPart!.WorksheetDrawing!;
    }

    /// <summary>A stream that can be read once and never seeked.</summary>
    private sealed class ForwardOnlyStream(byte[] content) : Stream
    {
        private readonly MemoryStream _inner = new(content);

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
