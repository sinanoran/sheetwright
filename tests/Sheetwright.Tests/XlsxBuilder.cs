using System.IO.Compression;
using System.Text;

namespace Sheetwright.Tests;

/// <summary>
/// A workbook assembled by hand, part by part, for the files no writer produces.
/// </summary>
/// <remarks>
/// The reader is for files from other programs and from people who mean harm, and Sheetwright's own writer
/// never produces inline strings, strict namespaces, the 1904 date system, a zip that lies about its size or a
/// tag with a million attributes. These tests need the bytes, so they write the XML themselves.
/// </remarks>
internal sealed class XlsxBuilder
{
    internal const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly List<(string Name, byte[] Bytes)> _parts = [];

    /// <summary>A one-sheet workbook with the given <c>sheetData</c> rows.</summary>
    internal static XlsxBuilder Workbook(
        string rows,
        string? sharedStrings = null,
        string? styles = null,
        bool date1904 = false,
        string sheetName = "Sheet1",
        string sheetAttributes = "")
    {
        XlsxBuilder builder = new();
        builder.Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
        builder.Add(
            "_rels/.rels",
            $"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        builder.Add(
            "xl/workbook.xml",
            $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><workbookPr{(date1904 ? " date1904=\"1\"" : string.Empty)}/>"
            + $"<sheets><sheet name=\"{sheetName}\" sheetId=\"1\" {sheetAttributes} r:id=\"rId1\"/></sheets></workbook>");

        StringBuilder rels = new($"<Relationships xmlns=\"{PackageRel}\">");
        rels.Append($"<Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
        if (sharedStrings is not null)
        {
            rels.Append($"<Relationship Id=\"rId2\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        }

        if (styles is not null)
        {
            rels.Append($"<Relationship Id=\"rId3\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/>");
        }

        rels.Append("</Relationships>");
        builder.Add("xl/_rels/workbook.xml.rels", rels.ToString());
        builder.Add("xl/worksheets/sheet1.xml", Sheet(rows));

        if (sharedStrings is not null)
        {
            builder.Add("xl/sharedStrings.xml", $"<sst xmlns=\"{Main}\">{sharedStrings}</sst>");
        }

        if (styles is not null)
        {
            builder.Add("xl/styles.xml", styles);
        }

        return builder;
    }

    internal static string Sheet(string rows) => $"<worksheet xmlns=\"{Main}\"><sheetData>{rows}</sheetData></worksheet>";

    /// <summary>A style part with the given custom formats and one <c>xf</c> per format id, in order.</summary>
    internal static string Styles(string numFmts, params int[] xfFormatIds)
    {
        string xfs = string.Concat(xfFormatIds.Select(id => $"<xf numFmtId=\"{id}\"/>"));
        return $"<styleSheet xmlns=\"{Main}\">{(numFmts.Length > 0 ? $"<numFmts>{numFmts}</numFmts>" : string.Empty)}"
            + $"<cellXfs count=\"{xfFormatIds.Length}\">{xfs}</cellXfs></styleSheet>";
    }

    internal XlsxBuilder Add(string name, string xml) => Add(name, Encoding.UTF8.GetBytes(xml));

    internal XlsxBuilder Add(string name, byte[] bytes)
    {
        _parts.Add((name, bytes));
        return this;
    }

    internal XlsxBuilder Replace(string name, string xml)
    {
        int index = _parts.FindIndex(part => part.Name == name);
        _parts[index] = (name, Encoding.UTF8.GetBytes(xml));
        return this;
    }

    internal XlsxBuilder Remove(string name)
    {
        _parts.RemoveAll(part => part.Name == name);
        return this;
    }

    internal MemoryStream ToStream(CompressionLevel level = CompressionLevel.Optimal)
    {
        MemoryStream memory = new();
        using (ZipArchive zip = new(memory, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach ((string name, byte[] bytes) in _parts)
            {
                ZipArchiveEntry entry = zip.CreateEntry(name, level);
                using Stream stream = entry.Open();
                stream.Write(bytes);
            }
        }

        memory.Position = 0;
        return memory;
    }

    /// <summary>The cell XML for a row of the common cell types, for tests that do not care about the markup.</summary>
    internal static string Row(int number, params string[] cells)
    {
        return $"<row r=\"{number}\">{string.Concat(cells)}</row>";
    }

    internal static string Number(string reference, string value, int? style = null)
    {
        return $"<c r=\"{reference}\"{(style is int s ? $" s=\"{s}\"" : string.Empty)}><v>{value}</v></c>";
    }

    internal static string Shared(string reference, int index) => $"<c r=\"{reference}\" t=\"s\"><v>{index}</v></c>";

    internal static string Inline(string reference, string text) => $"<c r=\"{reference}\" t=\"inlineStr\"><is><t>{text}</t></is></c>";
}
