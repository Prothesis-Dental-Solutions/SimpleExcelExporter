namespace SimpleExcelExporter.Tests
{
  using System.Globalization;
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

    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    [Test]
    public void RowsAndCells_HaveReferenceAttributes()
    {
      // Every emitted row carries r="N" and every emitted cell carries r="A1" with the correct
      // row number. Because empty cells are omitted, column indices are not contiguous — we only
      // verify that each cell's reference matches its actual position and that positions are
      // strictly increasing within a row.
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

        var lastColumnIndex = 0;
        foreach (var cell in row.Elements(ns + "c"))
        {
          var cellRef = cell.Attribute("r");
          Assert.That(cellRef, Is.Not.Null, $"Cell in row {expectedRowIndex} is missing the 'r' attribute");

          var letters = new string(cellRef!.Value.TakeWhile(char.IsLetter).ToArray());
          var rowPart = cellRef.Value[letters.Length ..];
          Assert.That(letters, Is.Not.Empty, $"Cell reference '{cellRef.Value}' has no column letters");
          Assert.That(rowPart, Is.EqualTo(expectedRowIndex.ToString()), $"Cell reference '{cellRef.Value}' does not match row {expectedRowIndex}");

          var columnIndex = LettersToColumnIndex(letters);
          Assert.That(columnIndex, Is.GreaterThan(lastColumnIndex), "Cells must appear in strictly increasing column order within a row");
          lastColumnIndex = columnIndex;
        }

        expectedRowIndex++;
      }
    }

    private static int LettersToColumnIndex(string letters)
    {
      var result = 0;
      foreach (var c in letters)
      {
        result = (result * 26) + (c - 'A' + 1);
      }

      return result;
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
    public void StringCells_UseSharedStringsOrEmptyInlineStr()
    {
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      // t="str" is reserved for formula results — never used by this writer.
      var strCells = sheetXml.Descendants(ns + "c")
        .Where(c => (string?)c.Attribute("t") == "str")
        .ToList();
      Assert.That(strCells, Is.Empty, "No cell should use t=\"str\" (reserved for formula results)");

      // Non-empty strings go through the shared strings table: t="s", <v>index</v>, no <is>.
      var sharedCells = sheetXml.Descendants(ns + "c")
        .Where(c => (string?)c.Attribute("t") == "s")
        .ToList();
      Assert.That(sharedCells, Is.Not.Empty, "Non-empty string cells should reference the shared strings table with t=\"s\"");

      foreach (var cell in sharedCells)
      {
        Assert.That(cell.Element(ns + "is"), Is.Null, $"t=\"s\" cell {cell.Attribute("r")?.Value} must NOT have <is> element");
        var value = cell.Element(ns + "v");
        Assert.That(value, Is.Not.Null, $"t=\"s\" cell {cell.Attribute("r")?.Value} must have <v> element with index");
        Assert.That(int.TryParse(value!.Value, out _), Is.True, $"t=\"s\" cell {cell.Attribute("r")?.Value} value must be an integer index");
      }

      // Empty strings remain as self-closing t="inlineStr" — see EmptyInlineStrCell_IsSelfClosingWithoutIsChild.
      var inlineStrCells = sheetXml.Descendants(ns + "c")
        .Where(c => (string?)c.Attribute("t") == "inlineStr")
        .ToList();
      foreach (var cell in inlineStrCells)
      {
        Assert.That(cell.Elements().Any(), Is.False, $"inlineStr cell {cell.Attribute("r")?.Value} must be empty (self-closing)");
      }
    }

    [Test]
    public void SharedStringsTable_IsConsistentWithCellReferences()
    {
      // Load sharedStrings.xml and verify every t="s" cell's index is within bounds, count/uniqueCount are sane.
      using var archive = GenerateAndOpenXlsxArchive();
      var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
      Assert.That(sharedEntry, Is.Not.Null, "Missing xl/sharedStrings.xml");
      using var sharedStream = sharedEntry!.Open();
      var sharedDoc = XDocument.Load(sharedStream);
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      var sst = sharedDoc.Root!;
      Assert.That(sst.Name, Is.EqualTo(ns + "sst"), "Root element of sharedStrings.xml must be <sst>");

      var items = sst.Elements(ns + "si").ToList();
      var uniqueCount = int.Parse(sst.Attribute("uniqueCount")!.Value, CultureInfo.InvariantCulture);
      Assert.That(uniqueCount, Is.EqualTo(items.Count), "uniqueCount attribute must match the number of <si> children");

      // Verify every <si><t> has non-empty text (empty strings go inline, not to the table).
      foreach (var si in items)
      {
        var text = si.Element(ns + "t");
        Assert.That(text, Is.Not.Null, "<si> must contain a <t> element");
        Assert.That(string.IsNullOrEmpty(text!.Value), Is.False, "<si><t> must not be empty — empty strings belong in inline cells");
      }

      // Verify every t="s" cell references a valid index.
      var sheetXml = LoadSheetXml();
      var sharedCells = sheetXml.Descendants(ns + "c").Where(c => (string?)c.Attribute("t") == "s").ToList();
      Assert.That(sharedCells, Is.Not.Empty);
      foreach (var cell in sharedCells)
      {
        var index = int.Parse(cell.Element(ns + "v")!.Value, CultureInfo.InvariantCulture);
        Assert.That(index, Is.InRange(0, items.Count - 1), $"Cell {cell.Attribute("r")?.Value} references index {index}, out of range");
      }
    }

    [Test]
    public void ContentTypes_AllDefaultsAppearBeforeAnyOverride()
    {
      // Strict OOXML parsers expect every <Default> entry to appear before the first <Override>.
      // The schema technically allows interleaving, but tools like Apple Numbers reject it.
      var contentTypesXml = LoadContentTypesXml();
      var ns = XNamespace.Get(ContentTypesNamespace);

      var children = contentTypesXml.Root!.Elements().ToList();
      var lastDefaultIndex = -1;
      var firstOverrideIndex = -1;
      for (var i = 0; i < children.Count; i++)
      {
        if (children[i].Name == ns + "Default")
        {
          lastDefaultIndex = i;
        }
        else if (children[i].Name == ns + "Override" && firstOverrideIndex == -1)
        {
          firstOverrideIndex = i;
        }
      }

      Assert.That(lastDefaultIndex, Is.GreaterThanOrEqualTo(0), "Expected at least one <Default> element");
      Assert.That(firstOverrideIndex, Is.GreaterThanOrEqualTo(0), "Expected at least one <Override> element");
      Assert.That(
        lastDefaultIndex,
        Is.LessThan(firstOverrideIndex),
        "All <Default> elements must appear before any <Override> element");
    }

    [Test]
    public void RelationshipTargets_AreRelative()
    {
      // Apple Numbers rejects .rels files whose <Relationship Target="..."> starts with '/'.
      // All targets must be relative paths. Applies to both _rels/.rels (package-level) and
      // xl/_rels/workbook.xml.rels (workbook-level).
      var relsNs = XNamespace.Get(PackageRelationshipsNamespace);
      using var archive = GenerateAndOpenXlsxArchive();

      foreach (var relsPath in new[] { "_rels/.rels", "xl/_rels/workbook.xml.rels" })
      {
        var entry = archive.GetEntry(relsPath);
        Assert.That(entry, Is.Not.Null, $"Missing {relsPath} in the generated package");
        using var stream = entry!.Open();
        var doc = XDocument.Load(stream);

        var relationships = doc.Descendants(relsNs + "Relationship").ToList();
        Assert.That(relationships, Is.Not.Empty, $"{relsPath} should contain at least one <Relationship>");

        foreach (var rel in relationships)
        {
          var target = rel.Attribute("Target")?.Value;
          var id = rel.Attribute("Id")?.Value;
          Assert.That(target, Is.Not.Null.And.Not.Empty, $"{relsPath}: Relationship {id} must have a non-empty Target");
          Assert.That(
            target,
            Does.Not.StartWith("/"),
            $"{relsPath}: Relationship {id} target '{target}' must be relative (no leading '/')");
        }
      }
    }

    [Test]
    public void EmptyCell_IsOmittedFromOutput()
    {
      // Fixture FirstFirstWithCollections: row3 column G is new CellDfn(string.Empty, CellDataType.String).
      // The library omits cells with no content entirely — OOXML readers infer empty positions
      // from the 'r' attribute of surrounding cells. This is also what Excel emits natively and is
      // accepted by Apple Numbers.
      var sheetXml = LoadSheetXml();
      var ns = XNamespace.Get(SpreadsheetMlNamespace);

      var cellG4 = sheetXml.Descendants(ns + "c")
        .SingleOrDefault(c => (string?)c.Attribute("r") == "G4");
      Assert.That(cellG4, Is.Null, "Empty cell G4 must not appear in the output; its position is inferred from F4 and H4");

      // Sanity: the neighbouring cells should still be present to bracket the missing G4.
      var cellF4 = sheetXml.Descendants(ns + "c").SingleOrDefault(c => (string?)c.Attribute("r") == "F4");
      var cellH4 = sheetXml.Descendants(ns + "c").SingleOrDefault(c => (string?)c.Attribute("r") == "H4");
      Assert.That(cellF4, Is.Not.Null, "Expected cell F4 to be present (non-empty in the fixture)");
      Assert.That(cellH4, Is.Not.Null, "Expected cell H4 to be present (non-empty in the fixture)");
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
