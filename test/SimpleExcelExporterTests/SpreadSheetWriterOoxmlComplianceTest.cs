namespace SimpleExcelExporter.Tests
{
  using System.IO;
  using System.IO.Compression;
  using System.Linq;
  using System.Xml.Linq;
  using NUnit.Framework;
  using SimpleExcelExporter.Tests.Preparators.Definitions;

  [TestFixture]
  public class SpreadSheetWriterOoxmlComplianceTest
  {
    private const string SpreadsheetMlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    [Test]
    public void RowsAndCells_HaveReferenceAttributes()
    {
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      var rows = sheetXml.Descendants(ns + "row").ToList();
      Assert.That(rows, Is.Not.Empty, "Expected at least one <row> element");

      uint expectedRowIndex = 1U;
      foreach (var row in rows)
      {
        var rowRef = row.Attribute("r");
        Assert.That(rowRef, Is.Not.Null, $"Row at position {expectedRowIndex} is missing the 'r' attribute");
        Assert.That(rowRef!.Value, Is.EqualTo(expectedRowIndex.ToString()), "Row 'r' attribute should be 1-indexed and contiguous");

        var columnIndex = 1;
        foreach (var cell in row.Elements(ns + "c"))
        {
          var cellRef = cell.Attribute("r");
          Assert.That(cellRef, Is.Not.Null, $"Cell at row {expectedRowIndex} column {columnIndex} is missing the 'r' attribute");
          var expected = $"{ColumnReferenceHelper.ToLetters(columnIndex)}{expectedRowIndex}";
          Assert.That(cellRef!.Value, Is.EqualTo(expected), $"Cell reference mismatch at row {expectedRowIndex} column {columnIndex}");
          columnIndex++;
        }

        expectedRowIndex++;
      }
    }

    [Test]
    public void SheetDimension_MatchesUsedRange()
    {
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      // Fixture FirstFirstWithCollections: 7 columns (A..G), 1 header row + 3 data rows = A1:G4
      var dimension = sheetXml.Descendants(ns + "dimension").SingleOrDefault();
      Assert.That(dimension, Is.Not.Null, "Expected a <dimension> element in the worksheet");
      Assert.That(dimension!.Attribute("ref")?.Value, Is.EqualTo("A1:G4"));
    }

    [Test]
    public void ContentTypes_DefaultXmlExtension_IsGenericApplicationXml()
    {
      var contentTypesXml = LoadContentTypesXml();
      var ns = XNamespace.Get(ContentTypesNamespace);

      var defaultXml = contentTypesXml.Descendants(ns + "Default").SingleOrDefault(d => (string?)d.Attribute("Extension") == "xml");
      Assert.That(defaultXml, Is.Not.Null, "Expected a <Default Extension='xml'> entry");
      Assert.That(defaultXml!.Attribute("ContentType")?.Value, Is.EqualTo("application/xml"), "Default xml extension must be generic application/xml");

      var overrides = contentTypesXml.Descendants(ns + "Override").ToList();
      var overrideForCore = overrides.SingleOrDefault(o => (string?)o.Attribute("PartName") == "/docProps/core.xml");
      Assert.That(overrideForCore, Is.Not.Null, "Expected a <Override> for /docProps/core.xml");
      Assert.That(overrideForCore!.Attribute("ContentType")?.Value, Is.EqualTo("application/vnd.openxmlformats-package.core-properties+xml"));

      var overrideForWorkbook = overrides.SingleOrDefault(o => (string?)o.Attribute("PartName") == "/xl/workbook.xml");
      Assert.That(overrideForWorkbook, Is.Not.Null, "Expected a <Override> for /xl/workbook.xml");

      var overrideForSheet1 = overrides.SingleOrDefault(o => (string?)o.Attribute("PartName") == "/xl/worksheets/sheet1.xml");
      Assert.That(overrideForSheet1, Is.Not.Null, "Expected a <Override> for /xl/worksheets/sheet1.xml");
    }

    private static XDocument LoadContentTypesXml()
    {
      using var archive = GenerateAndOpenXlsxArchive();
      var entry = archive.GetEntry("[Content_Types].xml");
      Assert.That(entry, Is.Not.Null, "Missing [Content_Types].xml in the generated package");
      using var stream = entry!.Open();
      return XDocument.Load(stream);
    }

    private static XDocument LoadSheetXml()
    {
      using var archive = GenerateAndOpenXlsxArchive();
      var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
      Assert.That(entry, Is.Not.Null, "Missing xl/worksheets/sheet1.xml in the generated package");
      using var stream = entry!.Open();
      return XDocument.Load(stream);
    }

    private static ZipArchive GenerateAndOpenXlsxArchive()
    {
      var workbook = WorkbookDfnPreparator.FirstFirstWithCollections();
      var memory = new MemoryStream();
      var writer = new SpreadsheetWriter(memory, workbook);
      writer.Write();
      memory.Position = 0;
      return new ZipArchive(memory, ZipArchiveMode.Read, leaveOpen: false);
    }
  }
}
