namespace SimpleExcelExporter.Tests
{
  using System.IO;
  using System.IO.Compression;
  using System.Linq;
  using System.Text;
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

      // Fixture FirstFirstWithCollections: header has 7 columns (A..G) but row3 has 8 cells (A..H).
      // maxColumnCount = max(7, 7, 7, 8) = 8 → used range is A1:H4.
      var dimension = sheetXml.Descendants(ns + "dimension").SingleOrDefault();
      Assert.That(dimension, Is.Not.Null, "Expected a <dimension> element in the worksheet");
      Assert.That(dimension!.Attribute("ref")?.Value, Is.EqualTo("A1:H4"));
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

    [Test]
    public void StringCells_UseInlineString_NotStr()
    {
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      var strCells = sheetXml.Descendants(ns + "c")
        .Where(c => (string?)c.Attribute("t") == "str")
        .ToList();
      Assert.That(strCells, Is.Empty, "No cell should use t=\"str\" (reserved for formula results)");

      var inlineStrCells = sheetXml.Descendants(ns + "c")
        .Where(c => (string?)c.Attribute("t") == "inlineStr")
        .ToList();
      Assert.That(inlineStrCells, Is.Not.Empty, "String cells should use t=\"inlineStr\"");

      // Verify inlineStr cells: non-empty ones must have <is><t>...</t></is>, none should have <v>
      foreach (var cell in inlineStrCells)
      {
        Assert.That(cell.Element(ns + "v"), Is.Null, $"inlineStr cell {cell.Attribute("r")?.Value} must NOT have <v> element");
        var inlineString = cell.Element(ns + "is");
        if (inlineString != null)
        {
          var text = inlineString.Element(ns + "t");
          Assert.That(text, Is.Not.Null, $"inlineStr cell {cell.Attribute("r")?.Value} must have <is><t> element");
        }
      }
    }

    [Test]
    public void EmptyInlineStrCell_IsSelfClosingWithoutIsChild()
    {
      // Fixture FirstFirstWithCollections: row3 column G is new CellDfn(string.Empty, CellDataType.String),
      // which becomes cell G4 (row 4 after the header row). Apple Numbers requires empty inlineStr
      // cells to be self-closing with no <is> child (not <c t="inlineStr"><is/></c> or similar).
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      var emptyCell = sheetXml.Descendants(ns + "c")
        .SingleOrDefault(c => (string?)c.Attribute("r") == "G4");
      Assert.That(emptyCell, Is.Not.Null, "Expected cell G4 (the empty string cell from the fixture)");
      Assert.That(emptyCell!.Attribute("t")?.Value, Is.EqualTo("inlineStr"), "G4 should declare t=\"inlineStr\"");
      Assert.That(emptyCell.Elements().Any(), Is.False, "Empty inlineStr cell must have no child element (no <is>, no <v>)");
      Assert.That(emptyCell.Value, Is.Empty, "Empty inlineStr cell must have no text content");

      using var archive = GenerateAndOpenXlsxArchive();
      var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
      using var stream = entry!.Open();
      using var reader = new StreamReader(stream, Encoding.UTF8);
      var content = reader.ReadToEnd();
      Assert.That(
        content,
        Does.Match("<c [^>]*r=\"G4\"[^>]*/>"),
        "Empty inlineStr cell G4 must be emitted as a self-closing element");
    }

    [Test]
    public void Worksheet_UsesDefaultNamespace()
    {
      using var archive = GenerateAndOpenXlsxArchive();
      var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
      Assert.That(entry, Is.Not.Null, "Missing xl/worksheets/sheet1.xml");
      using var stream = entry!.Open();
      using var reader = new StreamReader(stream, Encoding.UTF8);
      var content = reader.ReadToEnd();

      Assert.That(content, Does.Not.Contain("xmlns:x="), "Worksheet should use default namespace, not x: prefix");
      Assert.That(content, Does.Contain("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""), "Worksheet should declare SpreadsheetML namespace as default");
      Assert.That(content, Does.Not.Contain("<x:"), "Worksheet should not use x: prefix on any element");
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
