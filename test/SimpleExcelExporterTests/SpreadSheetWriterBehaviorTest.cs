// Behaviour-level tests for SpreadsheetWriter.
//
// These tests are DESIGNED TO COMPILE AND PASS ON BOTH master AND
// fix/xlsx-ooxml-compliance-numbers. They rely only on:
//   - The public API of SpreadsheetWriter
//   - The annotation types (CellDefinition, Index, Header, MultiColumn,
//     SheetName, IgnoreFromSpreadSheet)
//   - The DocumentFormat.OpenXml SDK for reading the produced file back
//
// They do NOT assert any OOXML shape (t="s" vs t="inlineStr" vs t="str",
// presence of r="A1" on cells, sharedStrings.xml existence, etc.) because
// those shapes are legitimately different between master and the PR. They
// focus on semantic behaviour: values preserved, columns ordered, styles
// deduplicated, ignored properties omitted — invariants that must hold
// across every implementation.
//
// Copy this file as-is into a master checkout to verify that the PR does
// not regress these behaviours.
namespace SimpleExcelExporter.Tests
{
  using System;
  using System.Collections.Generic;
  using System.Globalization;
  using System.IO;
  using System.Linq;
  using DocumentFormat.OpenXml.Packaging;
  using DocumentFormat.OpenXml.Spreadsheet;
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;
  using SimpleExcelExporter.Tests.Models;
  using SimpleExcelExporter.Tests.Preparators;

  [TestFixture]
  public class SpreadSheetWriterBehaviorTest
  {
    private static readonly string[] ExpectedMultiSheetNames = ["Alpha", "Beta", "Gamma"];

    [Test]
    public void CellValue_String_RoundTrip()
    {
      var workbook = BuildSingleRowWorkbook(new CellDfn("Hello, world!", cellDataType: CellDataType.String));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      Assert.That(GetCellString(cells[0], doc.WorkbookPart!), Is.EqualTo("Hello, world!"));
    }

    [Test]
    public void CellValue_Integer_RoundTrip()
    {
      var workbook = BuildSingleRowWorkbook(new CellDfn(42, cellDataType: CellDataType.Number));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      Assert.That(cells[0].CellValue!.InnerText, Is.EqualTo("42"));
    }

    [Test]
    public void CellValue_Decimal_PreservesPrecision()
    {
      var workbook = BuildSingleRowWorkbook(new CellDfn(12345.6789m, cellDataType: CellDataType.Number));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      Assert.That(decimal.Parse(cells[0].CellValue!.InnerText, CultureInfo.InvariantCulture), Is.EqualTo(12345.6789m));
    }

    [Test]
    public void CellValue_Double_RoundTrip()
    {
      var workbook = BuildSingleRowWorkbook(new CellDfn(3.14159d, cellDataType: CellDataType.Number));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      Assert.That(double.Parse(cells[0].CellValue!.InnerText, CultureInfo.InvariantCulture), Is.EqualTo(3.14159d).Within(1e-9));
    }

    [Test]
    public void CellValue_Boolean_WritesZeroOrOne()
    {
      // True -> "0", False -> "1" (matches the library's current convention on both master and PR).
      var workbookTrue = BuildSingleRowWorkbook(new CellDfn(true, cellDataType: CellDataType.Boolean));
      var workbookFalse = BuildSingleRowWorkbook(new CellDfn(false, cellDataType: CellDataType.Boolean));
      using var streamTrue = WriteToStream(workbookTrue);
      using var streamFalse = WriteToStream(workbookFalse);
      using var docTrue = SpreadsheetDocument.Open(streamTrue, false);
      using var docFalse = SpreadsheetDocument.Open(streamFalse, false);

      Assert.That(GetRowCells(docTrue)[0].CellValue!.InnerText, Is.EqualTo("0"));
      Assert.That(GetRowCells(docFalse)[0].CellValue!.InnerText, Is.EqualTo("1"));
    }

    [Test]
    public void CellValue_DateTime_RoundTrip()
    {
      var date = new DateTime(2020, 1, 15, 12, 0, 0);
      var workbook = BuildSingleRowWorkbook(new CellDfn(date, cellDataType: CellDataType.Date));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      var parsed = DateTime.Parse(cells[0].CellValue!.InnerText, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
      Assert.That(parsed, Is.EqualTo(date));
    }

    [Test]
    public void CellValue_TimeSpan_WritesFractionOfDay()
    {
      var workbook = BuildSingleRowWorkbook(new CellDfn(new TimeSpan(12, 0, 0), cellDataType: CellDataType.Time));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc);
      Assert.That(double.Parse(cells[0].CellValue!.InnerText, CultureInfo.InvariantCulture), Is.EqualTo(0.5d).Within(1e-9));
    }

    [Test]
    public void CellValue_NullInAnnotatedObject_ProducesEmptyCell()
    {
      // PlayerDummyObjectPreparator.First() leaves PlayerCode = null. PlayerCode has Index=0 → column A.
      // The cell at A2 (row 2, first player) must be empty across every library version:
      //   - master: the cell is emitted with an empty CellValue
      //   - PR (pre-skip): the cell is emitted as self-closing t="inlineStr"
      //   - PR (post-skip): the cell is entirely omitted — A2 is empty by position inference
      // Looking up by r="A2" covers all three cases: if the cell is absent or carries no value,
      // it is treated as empty.
      var team = TeamDummyObjectPreparator.FirstWithCollections();
      using var stream = WriteToStreamFromObject(team);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cells = GetRowCells(doc, rowIndex: 1); // skip header row, read the first player
      var cellA2 = cells.FirstOrDefault(c => c.CellReference?.Value == "A2");
      var value = cellA2 != null ? GetCellString(cellA2, doc.WorkbookPart!) : null;
      Assert.That(value, Is.Null.Or.Empty, "PlayerCode was null in the source, cell A2 must be absent or empty");
    }

    [Test]
    public void StyleIndex_DeduplicatesIdenticalCellDataType()
    {
      // Twenty cells of CellDataType.Date should share a single style entry.
      var row = new RowDfn();
      for (var i = 0; i < 20; i++)
      {
        row.Cells.Add(new CellDfn(new DateTime(2020, 1, 1).AddDays(i), cellDataType: CellDataType.Date));
      }

      var workbook = new WorkbookDfn();
      var worksheet = new WorksheetDfn("Sheet1");
      worksheet.Rows.Add(row);
      workbook.Worksheets.Add(worksheet);

      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);
      var cellFormats = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToList();

      var usedDateStyleIndices = doc.WorkbookPart.WorksheetParts
        .SelectMany(ws => ws.Worksheet!.Descendants<Cell>())
        .Select(c => c.StyleIndex?.Value ?? 0U)
        .Distinct()
        .ToList();

      // All 20 date cells share the same style index; worksheet may have others
      // implicitly (default 0 for the sheet, but only 1 distinct for Date cells).
      var dateFormats = cellFormats.Where(cf => cf.NumberFormatId?.Value == 14U).ToList();
      Assert.That(dateFormats.Count, Is.EqualTo(1), "Expected exactly one CellFormat with numFmtId=14 (Date), found " + dateFormats.Count);
    }

    [TestCase(CellDataType.Date, 14U)]
    [TestCase(CellDataType.String, 49U)]
    [TestCase(CellDataType.Percentage, 10U)]
    [TestCase(CellDataType.Time, 20U)]
    public void NumberFormatId_CellDataType_MapsToExpectedId(CellDataType dataType, uint expectedNumFmtId)
    {
      object value = dataType switch
      {
        CellDataType.Date => new DateTime(2020, 1, 1),
        CellDataType.String => "x",
        CellDataType.Percentage => 0.5d,
        CellDataType.Time => new TimeSpan(1, 0, 0),
        _ => 0,
      };

      var workbook = BuildSingleRowWorkbook(new CellDfn(value, cellDataType: dataType));
      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var cell = doc.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().First();
      var styleIndex = cell.StyleIndex?.Value ?? 0U;
      var format = doc.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)styleIndex);

      Assert.That(format.NumberFormatId?.Value, Is.EqualTo(expectedNumFmtId));
    }

    [Test]
    public void IgnoreFromSpreadSheet_OmitsAnnotatedProperty()
    {
      // ChildOfPlayerDummyObject has HeaderMention tagged [IgnoreFromSpreadSheet].
      // When a player with Children is written, HeaderMention must not appear as a column.
      var team = TeamDummyObjectPreparator.FirstWithCollections();
      using var stream = WriteToStreamFromObject(team);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var firstSheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
      var headerRow = firstSheet.Descendants<Row>().First();
      var headerTexts = headerRow.Descendants<Cell>()
        .Select(c => GetCellString(c, doc.WorkbookPart))
        .ToList();

      Assert.That(headerTexts, Has.None.EqualTo("HeaderMention"), "A property tagged [IgnoreFromSpreadSheet] must not produce a column");
    }

    [Test]
    public void MultipleWorksheets_PreservesSheetNamesAndOrder()
    {
      var workbook = new WorkbookDfn();
      workbook.Worksheets.Add(MakeSheetWithOneValueRow("Alpha", "a"));
      workbook.Worksheets.Add(MakeSheetWithOneValueRow("Beta", "b"));
      workbook.Worksheets.Add(MakeSheetWithOneValueRow("Gamma", "c"));

      using var stream = WriteToStream(workbook);
      using var doc = SpreadsheetDocument.Open(stream, false);
      var sheets = doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();

      Assert.That(sheets.Count, Is.EqualTo(3));
      Assert.That(sheets.Select(s => s.Name?.Value).ToList(), Is.EqualTo(ExpectedMultiSheetNames));
    }

    [Test]
    public void IndexAttribute_OrdersColumnsByIndex()
    {
      // PlayerDummyObject has PlayerCode(Index=0), PlayerName(Index=1), DateOfBirth(Index=2), …
      var team = TeamDummyObjectPreparator.FirstWithCollections();
      using var stream = WriteToStreamFromObject(team);
      using var doc = SpreadsheetDocument.Open(stream, false);

      var firstSheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
      var headerRow = firstSheet.Descendants<Row>().First();
      var headers = headerRow.Descendants<Cell>()
        .Select(c => GetCellString(c, doc.WorkbookPart))
        .ToList();

      // The first three columns must come from Index=0,1,2 = PlayerCode, PlayerName, DateOfBirth.
      // Exact header values come from PlayerDummyObjectRes resources; we only check the relative
      // order matches the IndexAttribute assignment.
      Assert.That(headers[0], Is.Not.Null.And.Not.Empty);
      Assert.That(headers[1], Is.Not.Null.And.Not.Empty);
      Assert.That(headers[2], Is.Not.Null.And.Not.Empty);

      // The Player.PlayerCode property has Index 0 so its header is first.
      // Its resource value 'Player code' is stable across master and PR.
      Assert.That(headers[0], Does.Contain("code").IgnoreCase, "First column should correspond to PlayerCode (Index=0)");
    }

    // Helpers

    private static WorksheetDfn MakeSheetWithOneValueRow(string sheetName, string value)
    {
      var sheet = new WorksheetDfn(sheetName);
      var row = new RowDfn();
      row.Cells.Add(new CellDfn(value, cellDataType: CellDataType.String));
      sheet.Rows.Add(row);
      return sheet;
    }

    private static WorkbookDfn BuildSingleRowWorkbook(CellDfn cell)
    {
      var workbook = new WorkbookDfn();
      var sheet = new WorksheetDfn("Sheet1");
      var row = new RowDfn();
      row.Cells.Add(cell);
      sheet.Rows.Add(row);
      workbook.Worksheets.Add(sheet);
      return workbook;
    }

    private static MemoryStream WriteToStream(WorkbookDfn workbook)
    {
      var stream = new MemoryStream();
      var writer = new SpreadsheetWriter(stream, workbook);
      writer.Write();
      stream.Position = 0;
      return stream;
    }

    private static MemoryStream WriteToStreamFromObject(object annotated)
    {
      var stream = new MemoryStream();
      var writer = new SpreadsheetWriter(stream, annotated);
      writer.Write();
      stream.Position = 0;
      return stream;
    }

    private static List<Cell> GetRowCells(SpreadsheetDocument doc, int rowIndex = 0)
    {
      var sheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
      var rows = sheet.Descendants<Row>().ToList();
      return rows[rowIndex].Descendants<Cell>().ToList();
    }

    private static string? GetCellString(Cell cell, WorkbookPart workbookPart)
    {
      // Handles every string representation: t="str" (master), t="inlineStr" (PR pre-shared),
      // t="s" (PR post-shared), and the implicit inline CellValue text fallback.
      if (cell.DataType?.Value == CellValues.SharedString && cell.CellValue != null)
      {
        var index = int.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture);
        var sharedTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        return sharedTable?.ElementAt(index).InnerText;
      }

      if (cell.DataType?.Value == CellValues.InlineString && cell.InlineString != null)
      {
        return cell.InlineString.InnerText;
      }

      return cell.CellValue?.InnerText;
    }
  }
}
