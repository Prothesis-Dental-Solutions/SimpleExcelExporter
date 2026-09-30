namespace SimpleExcelExporter.Tests
{
  using System;
  using System.Collections.Generic;
  using System.IO;
  using System.Linq;
  using DocumentFormat.OpenXml;
  using DocumentFormat.OpenXml.Packaging;
  using DocumentFormat.OpenXml.Spreadsheet;
  using DocumentFormat.OpenXml.Validation;
  using NUnit.Framework;
  using SimpleExcelExporter.Definitions;
  using SimpleExcelExporter.Resources;
  using SimpleExcelExporter.Tests.Preparators;
  using SimpleExcelExporter.Tests.Preparators.Definitions;

  [TestFixture]
  public class SpreadSheetWriterTest
  {
    [Test]
    public void WriteTest()
    {
      // Prepare an empty workbook
      var workBookDfn = WorkbookDfnPreparator.First();

      // Act && Check
      // ReSharper disable once ObjectCreationAsStatement
      Assert.Throws<DefinitionException>((Action)(() => new SpreadsheetWriter(new MemoryStream(), workBookDfn)));

      // Prepare a non-empty workbook
      workBookDfn = WorkbookDfnPreparator.FirstFirstWithCollections();
      using var memoryStream = new MemoryStream();

      // Act
      var spreadsheetWriter = new SpreadsheetWriter(memoryStream, workBookDfn);
      spreadsheetWriter.Write();

      // Check
      Assert.That(memoryStream.Length, Is.Not.Zero);
      Validate(memoryStream, 1, 4, 7);

      // Prepare an object
      var team = TeamDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, team);
      spreadsheetWriter.Write();

      // Check
      Assert.That(memoryStream.Length, Is.Not.Zero);
      Validate(memoryStream, 1, 1, 1);

      // Prepare an object
      team = TeamDummyObjectPreparator.FirstWithCollections();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, team);
      spreadsheetWriter.Write();

      // Check
      Assert.That(memoryStream.Length, Is.Not.Zero);
      // expected 1 sheet, 6 rows (1 header + 5 players + 2 children of player), 20 cells
      Validate(memoryStream, 1, 8, 20);

      // Prepare an empty object - two properties with the same index column
      var teamWithSameColumnIndex = TeamWithSameColumnIndexDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, teamWithSameColumnIndex);
      spreadsheetWriter.Write();

      // Check
      Assert.That(memoryStream.Length, Is.Not.Zero);
      Validate(memoryStream, 1, 1, 1);

      // Prepare an object - with same index column
      teamWithSameColumnIndex = TeamWithSameColumnIndexDummyObjectPreparator.FirstWithCollections();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, teamWithSameColumnIndex);
      spreadsheetWriter.Write();

      // Check
      Assert.That(memoryStream.Length, Is.Not.Zero);
      // expected 1 sheet, 4 rows (1 header + 3 players), 3 cells in the header row.
      // PlayerWithSameColumnIndexDummyObject.FourthColumn has no [Header] attribute, so its
      // empty header cell is omitted from the output (the library skips cells with no content).
      Validate(memoryStream, 1, 4, 3);

      // Prepare with object - with same sheet
      var teamWithSameSheetName = TeamWithSameSheetNameDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      // ReSharper disable once ObjectCreationAsStatement
      Assert.Throws<DefinitionException>((Action)(() => new SpreadsheetWriter(memoryStream, teamWithSameSheetName)));
    }

    [Test]
    public void SpreadsheetExportSheetNameLengthTest()
    {
      // Prepare a non-empty workbook
      const string tooLongSheetName = "Name with something bigger than 31 characters.";
      var workBookDfn = WorkbookDfnPreparator.First();
      workBookDfn.Worksheets.Add(new WorksheetDfn(tooLongSheetName));
      using var memoryStream = new MemoryStream();
      var expected = new SimpleExcelExporterException(string.Format(MessageRes.SheetNameLengthTooLong, tooLongSheetName));

      // Act
      // ReSharper disable once ObjectCreationAsStatement
      var simpleExcelExporterException = Assert.Throws<SimpleExcelExporterException>((Action)(() => new SpreadsheetWriter(memoryStream, workBookDfn)));

      // Check
      Assert.That(simpleExcelExporterException, Is.Not.Null);
      Assert.That(simpleExcelExporterException!.Message, Is.EqualTo(expected.Message));
    }

    private static readonly List<string> ExpectedErrors =
    [
      "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
    ];

    private static void Validate(
      Stream memoryStream,
      int expectedSheetsCount,
      int expectedRowsCount,
      int expectedCellsCount)
    {
      using var spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, true);
      var validator = new OpenXmlValidator();
      var errors = validator.Validate(spreadsheetDocument).Where(validationError => !ExpectedErrors.Contains(validationError.Description));
      Assert.That(errors, Is.Empty);

      var fileFormat = validator.FileFormat;
      Assert.That(fileFormat, Is.EqualTo(FileFormatVersions.Office2007));

      var workbookPart = spreadsheetDocument.WorkbookPart;
      var worksheetsPart = workbookPart!.WorksheetParts.First();
      var sheetData = worksheetsPart.Worksheet!.GetFirstChild<SheetData>();
      var rows = sheetData!.Descendants<Row>().ToList();
      var cells = rows[0].Descendants<Cell>();

      Assert.That(workbookPart.Workbook, Is.Not.Null);
      Assert.That(workbookPart.Workbook!.Sheets, Is.Not.Null);
      using (Assert.EnterMultipleScope())
      {
        Assert.That(workbookPart.Workbook.Sheets!.Count(), Is.EqualTo(expectedSheetsCount));
        Assert.That(rows, Has.Count.EqualTo(expectedRowsCount));
        Assert.That(cells.Count(), Is.EqualTo(expectedCellsCount));
      }
    }
  }
}
