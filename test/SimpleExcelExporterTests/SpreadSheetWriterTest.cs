namespace SimpleExcelExporter.Tests
{
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
      Assert.Throws<DefinitionException>(() => new SpreadsheetWriter(new MemoryStream(), workBookDfn));

      // Prepare a non-empty workbook
      workBookDfn = WorkbookDfnPreparator.FirstFirstWithCollections();
      using var memoryStream = new MemoryStream();

      // Act
      var spreadsheetWriter = new SpreadsheetWriter(memoryStream, workBookDfn);
      spreadsheetWriter.Write();

      // Check
      Assert.AreNotEqual(memoryStream.Length, 0);
      Validate(memoryStream, 1, 3, 7);

      // Prepare an object
      var team = TeamDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, team);
      spreadsheetWriter.Write();

      // Check
      Assert.AreNotEqual(memoryStream.Length, 0);
      Validate(memoryStream, 1, 1, 1);

      // Prepare an object
      team = TeamDummyObjectPreparator.FirstWithCollections();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, team);
      spreadsheetWriter.Write();

      // Check
      Assert.AreNotEqual(memoryStream.Length, 0);
      // expected 1 sheet, 5 rows (1 header + 4 players + 2 children of player), 15 cells
      Validate(memoryStream, 1, 5, 15);

      // Prepare an empty object - two properties with the same index column
      var teamWithSameColumnIndex = TeamWithSameColumnIndexDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, teamWithSameColumnIndex);
      spreadsheetWriter.Write();

      // Check
      Assert.AreNotEqual(memoryStream.Length, 0);
      Validate(memoryStream, 1, 1, 1);

      // Prepare an object - with same index column
      teamWithSameColumnIndex = TeamWithSameColumnIndexDummyObjectPreparator.FirstWithCollections();
      memoryStream.SetLength(0);

      // Act
      spreadsheetWriter = new SpreadsheetWriter(memoryStream, teamWithSameColumnIndex);
      spreadsheetWriter.Write();

      // Check
      Assert.AreNotEqual(memoryStream.Length, 0);
      // expected 1 sheet, 4 rows (1 header + 3 players), 3 cells
      Validate(memoryStream, 1, 4, 4);

      // Prepare with object - with same sheet
      var teamWithSameSheetName = TeamWithSameSheetNameDummyObjectPreparator.First();
      memoryStream.SetLength(0);

      // Act
      // ReSharper disable once ObjectCreationAsStatement
      Assert.Throws<DefinitionException>(() => new SpreadsheetWriter(memoryStream, teamWithSameSheetName));
    }

    [Test]
    public void SpreadsheetExportSheetNameLengthTest()
    {
      // Prepare a non-empty workbook
      var tooLongSheetName = "Name with something bigger than 31 characters.";
      var workBookDfn = WorkbookDfnPreparator.First();
      workBookDfn.Worksheets.Add(new WorksheetDfn(tooLongSheetName));
      using var memoryStream = new MemoryStream();
      var expected = new SimpleExcelExporterException(string.Format(MessageRes.SheetNameLengthTooLong, tooLongSheetName));

      // Act
      // ReSharper disable once ObjectCreationAsStatement
      SimpleExcelExporterException simpleExcelExporterException = Assert.Throws<SimpleExcelExporterException>(() => new SpreadsheetWriter(memoryStream, workBookDfn));

      // Check
      Assert.IsNotNull(simpleExcelExporterException);
      Assert.AreEqual(expected.Message, simpleExcelExporterException.Message);
    }

    private static readonly List<string> ExpectedErrors = new()
    {
      "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
    };

    private static void Validate(
      Stream memoryStream,
      int expectedSheetsCount,
      int expectedRowsCount,
      int expectedCellsCount)
    {
      using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(memoryStream, true);
      OpenXmlValidator validator = new OpenXmlValidator();
      var errors = validator.Validate(spreadsheetDocument).Where(validationError => !ExpectedErrors.Contains(validationError.Description));
      Assert.IsEmpty(errors);

      var fileFormat = validator.FileFormat;
      Assert.AreEqual(fileFormat, FileFormatVersions.Office2007);

      var workbookPart = spreadsheetDocument.WorkbookPart;
      var worksheetsPart = workbookPart!.WorksheetParts.First();
      var sheetData = worksheetsPart.Worksheet.GetFirstChild<SheetData>();
      var rows = sheetData!.Descendants<Row>().ToList();
      var cells = rows.First().Descendants<Cell>();

      Assert.IsNotNull(workbookPart.Workbook);
      Assert.IsNotNull(workbookPart.Workbook.Sheets);
      Assert.AreEqual(expectedSheetsCount, workbookPart.Workbook.Sheets!.Count());
      Assert.AreEqual(expectedRowsCount, rows.Count);
      Assert.AreEqual(expectedCellsCount, cells.Count());
    }
  }
}
