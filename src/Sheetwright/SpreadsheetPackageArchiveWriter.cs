using System.IO.Compression;
using System.Text;
using System.Xml;


namespace Sheetwright;

internal static class SpreadsheetPackageArchiveWriter
{
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    internal static void Write(Stream destination, IEnumerable<SpreadsheetMemoryPackagePart> parts)
    {
        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        if (parts is null)
        {
            throw new ArgumentNullException(nameof(parts));
        }

        SpreadsheetMemoryPackagePart[] packageParts = parts.OrderBy(part => part.Uri.OriginalString).ToArray();
        destination.SetLength(0);
        destination.Position = 0;

        using (ZipArchive archive = new(destination, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteContentTypes(archive, packageParts);

            foreach (SpreadsheetMemoryPackagePart part in packageParts)
            {
                ZipArchiveEntry entry = archive.CreateEntry(
                    part.Uri.OriginalString.TrimStart('/'),
                    CompressionLevel.Optimal);
                using Stream entryStream = entry.Open();
                part.CopyContentTo(entryStream);
            }
        }

        destination.Position = 0;
    }

    private static void WriteContentTypes(
        ZipArchive archive,
        IEnumerable<SpreadsheetMemoryPackagePart> parts)
    {
        ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        using XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false
        });

        writer.WriteStartDocument();
        writer.WriteStartElement("Types", ContentTypesNamespace);
        foreach (SpreadsheetMemoryPackagePart part in parts)
        {
            writer.WriteStartElement("Override", ContentTypesNamespace);
            writer.WriteAttributeString("PartName", part.Uri.OriginalString);
            writer.WriteAttributeString("ContentType", part.ContentType);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }
}
