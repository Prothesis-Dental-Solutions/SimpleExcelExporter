// ReSharper disable LocalizableElement

namespace ConsoleApp
{
  using System;
  using System.Diagnostics;
  using System.IO;
  using SimpleExcelExporter;
  using SimpleExcelExporter.Definitions;

  public static class Program
  {
    public static void Main()
    {
      // First test : try to create empty excel file
      var n = DateTime.Now;
      var tempDi = new DirectoryInfo($"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");
      tempDi.Create();

      GenerateSpreadSheetFromWorkbookDfn(tempDi);
      GenerateSpreadSheetFromAnnotatedDataEmpty(tempDi);
      GenerateSpreadSheetFromAnnotatedData(tempDi);
      GenerateBigSpreadsheetFromAnnotatedData(tempDi);
      GenerateBigSpreadsheetFromWorkBookDfn(tempDi);
    }

    private static void GenerateBigSpreadsheetFromAnnotatedData(DirectoryInfo tempDi)
    {
      Console.WriteLine("GenerateBigSpreadsheetFromAnnotatedData");
      using var memoryStream = new MemoryStream();
      using var streamWriter = new StreamWriter(memoryStream);
      var team = new Team();
      Random rnd = new Random();
      RandomDateTime randomDate = new RandomDateTime();
      Console.WriteLine("Generating the players...");
      Stopwatch stopwatch = new Stopwatch();
      stopwatch.Start();

      var nullPlayer = new Player
      {
        PlayerCode = "Code0",
        PlayerName = null,
        Size = null,
        DateOfBirth = null,
        IsActiveFlag = null,
        NumberOfVictory = null,
        FieldGoalPercentage = null,
        Salary = null,
      };
      team.Players.Add(nullPlayer);

      for (int i = 1; i < 1000000; i++)
      {
        var player = new Player
        {
          PlayerCode = $"Code{i}",
          PlayerName = $"Player{i}",
          Size = rnd.Next(18, 100),
          DateOfBirth = randomDate.Next(),
          IsActiveFlag = Convert.ToBoolean(rnd.Next(0, 100) % 2),
          NumberOfVictory = rnd.Next(0, 100),
          FieldGoalPercentage = Convert.ToDouble(rnd.Next(0, 100)) / 100,
          Salary = Convert.ToDecimal(rnd.Next(2000, 1000000) + 0.12654984m),
        };
        team.Players.Add(player);
      }

      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      Console.WriteLine("Instantiating the SpreadsheetWriter...");
      stopwatch.Reset();
      stopwatch.Start();
      SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, team);
      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      Console.WriteLine("Writing the Excel file...");
      stopwatch.Reset();
      stopwatch.Start();
      spreadsheetWriter.Write();
      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithData3.xlsx"), FileMode.Create, FileAccess.Write);
      memoryStream.WriteTo(file);
    }

    private static void GenerateBigSpreadsheetFromWorkBookDfn(DirectoryInfo tempDi)
    {
      Console.WriteLine("GenerateBigSpreadsheetFromWorkBookDfn");
      using var memoryStream = new MemoryStream();
      using var streamWriter = new StreamWriter(memoryStream);
      var workbookDfn = new WorkbookDfn();
      var worksheetDfn = new WorksheetDfn("Team");
      workbookDfn.Worksheets.Add(worksheetDfn);
      Random rnd = new Random();
      RandomDateTime randomDate = new RandomDateTime();
      Console.WriteLine("Generating the players...");
      Stopwatch stopwatch = new Stopwatch();
      stopwatch.Start();
      for (int i = 0; i < 1000000; i++)
      {
        var rowDfn = new RowDfn
        {
          Cells =
          {
            new CellDfn($"Code{i}"),
            new CellDfn($"Player{i}"),
            new CellDfn(rnd.Next(18, 100), cellDataType: CellDataType.Number),
            new CellDfn(randomDate.Next(), cellDataType: CellDataType.Date),
            new CellDfn(Convert.ToBoolean(rnd.Next(0, 100) % 2), cellDataType: CellDataType.Boolean),
            new CellDfn(rnd.Next(0, 100), cellDataType: CellDataType.Number),
            new CellDfn(Convert.ToDouble(rnd.Next(0, 100)) / 100, cellDataType: CellDataType.Percentage),
          },
        };
        worksheetDfn.Rows.Add(rowDfn);
      }

      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      Console.WriteLine("Instantiating the SpreadsheetWriter...");
      stopwatch.Reset();
      stopwatch.Start();
      SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, workbookDfn);
      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      Console.WriteLine("Writing the Excel file...");
      stopwatch.Reset();
      stopwatch.Start();
      spreadsheetWriter.Write();
      stopwatch.Stop();
      Console.WriteLine($"Done in {stopwatch.Elapsed.Seconds} seconds !");

      using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithData4.xlsx"), FileMode.Create, FileAccess.Write);
      memoryStream.WriteTo(file);
    }

    private static void GenerateSpreadSheetFromAnnotatedData(DirectoryInfo tempDi)
    {
      Console.WriteLine("GenerateSpreadSheetFromAnnotatedData");
      using var memoryStream = new MemoryStream();
      using var streamWriter = new StreamWriter(memoryStream);
      var team = new Team
      {
        Players =
        {
          new Player
          {
            PlayerCode = null,
            PlayerName = "Alexandre",
            Size = 1.93d,
            DateOfBirth = new DateTime(1974, 02, 01),
            IsActiveFlag = true,
            NumberOfVictory = 45,
            FieldGoalPercentage = 0.1111,
            Salary = 2000.5m,
          },
          new Player
          {
            PlayerCode = "02",
            PlayerName = "Elina",
            Size = 1.72d,
            DateOfBirth = new DateTime(1990, 10, 13),
            IsActiveFlag = true,
            NumberOfVictory = 52,
            FieldGoalPercentage = 0.222,
            Salary = 2141.5452m,
          },
          new Player
          {
            PlayerCode = "03",
            PlayerName = "Franck",
            Size = 1.85d,
            DateOfBirth = new DateTime(1976, 3, 1),
            IsActiveFlag = true,
            NumberOfVictory = 80,
            FieldGoalPercentage = 0.33,
            Salary = 2111.5452m,
          },
          new Player
          {
            PlayerCode = "04",
            PlayerName = "Yann",
            Size = 1.79d,
            DateOfBirth = new DateTime(1979, 3, 1),
            IsActiveFlag = false,
            NumberOfVictory = 35,
            FieldGoalPercentage = 0.4,
            Salary = 2845.719m,
          },
        },
      };

      SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, team);
      spreadsheetWriter.Write();

      using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithData2.xlsx"), FileMode.Create, FileAccess.Write);
      memoryStream.WriteTo(file);
    }

    private static void GenerateSpreadSheetFromAnnotatedDataEmpty(DirectoryInfo tempDi)
    {
      Console.WriteLine("GenerateSpreadSheetFromAnnotatedDataEmpty");
      using var memoryStream = new MemoryStream();
      using var streamWriter = new StreamWriter(memoryStream);
      var team = new Team();

      SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, team);
      spreadsheetWriter.Write();

      using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithDataEmpty.xlsx"), FileMode.Create, FileAccess.Write);
      memoryStream.WriteTo(file);
    }

    private static void GenerateSpreadSheetFromWorkbookDfn(DirectoryInfo tempDi)
    {
      Console.WriteLine("GenerateSpreadSheetFromWorkbookDfn");
      using var memoryStream = new MemoryStream();
      using var streamWriter = new StreamWriter(memoryStream);
      var workbookDfn = new WorkbookDfn();

      // First sheet
      var worksheet1Dfn = new WorksheetDfn("MyFirstSheet");
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Name|\b|\n|\t|\r|<|>|&|'|\"|"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Age"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Rate"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Postal code"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("DateTime"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Field goal percentage"));
      workbookDfn.Worksheets.Add(worksheet1Dfn);
      var row1 = new RowDfn();
      row1.Cells.Add(new CellDfn("Eric", cellDataType: CellDataType.String));
      row1.Cells.Add(new CellDfn(50, cellDataType: CellDataType.Number));
      row1.Cells.Add(new CellDfn(45.00M, cellDataType: CellDataType.Number));
      row1.Cells.Add(new CellDfn("01090", cellDataType: CellDataType.String));
      row1.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row1.Cells.Add(new CellDfn(0.0111, cellDataType: CellDataType.Percentage));
      worksheet1Dfn.Rows.Add(row1);
      var row2 = new RowDfn();
      row2.Cells.Add(new CellDfn("Bob", cellDataType: CellDataType.String));
      row2.Cells.Add(new CellDfn(42, cellDataType: CellDataType.Number));
      row2.Cells.Add(new CellDfn(78.00M, cellDataType: CellDataType.Number));
      row2.Cells.Add(new CellDfn("01080", cellDataType: CellDataType.String));
      row2.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row2.Cells.Add(new CellDfn(0.0222, cellDataType: CellDataType.Percentage));
      worksheet1Dfn.Rows.Add(row2);

      // Second sheet
      var worksheet2Dfn = new WorksheetDfn("MySecondSheet");
      worksheet2Dfn.ColumnHeadings.Cells.Add(new CellDfn("Name"));
      worksheet2Dfn.ColumnHeadings.Cells.Add(new CellDfn("Age"));
      worksheet2Dfn.ColumnHeadings.Cells.Add(new CellDfn("Rate"));
      worksheet2Dfn.ColumnHeadings.Cells.Add(new CellDfn("Postal Code"));
      worksheet2Dfn.ColumnHeadings.Cells.Add(new CellDfn("Field goal percentage"));
      workbookDfn.Worksheets.Add(worksheet2Dfn);

      // Third sheet
      var worksheet3Dfn = new WorksheetDfn("MyThirdSheet");
      workbookDfn.Worksheets.Add(worksheet3Dfn);
      var row31 = new RowDfn();
      row31.Cells.Add(new CellDfn("Eric", cellDataType: CellDataType.String));
      row31.Cells.Add(new CellDfn(50, cellDataType: CellDataType.Number));
      row31.Cells.Add(new CellDfn(45.00M, cellDataType: CellDataType.Number));
      row31.Cells.Add(new CellDfn("01090", cellDataType: CellDataType.String));
      row31.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row31.Cells.Add(new CellDfn(true, cellDataType: CellDataType.Boolean));
      row31.Cells.Add(new CellDfn(0.11, cellDataType: CellDataType.Percentage));
      worksheet3Dfn.Rows.Add(row31);
      var row32 = new RowDfn();
      row32.Cells.Add(new CellDfn("Bob", cellDataType: CellDataType.String));
      row32.Cells.Add(new CellDfn(42, cellDataType: CellDataType.Number));
      row32.Cells.Add(new CellDfn(78.00M, cellDataType: CellDataType.Number));
      row32.Cells.Add(new CellDfn("01080", cellDataType: CellDataType.String));
      row32.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row32.Cells.Add(new CellDfn(false, cellDataType: CellDataType.Boolean));
      row31.Cells.Add(new CellDfn(22, cellDataType: CellDataType.Percentage));
      worksheet3Dfn.Rows.Add(row32);

      SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, workbookDfn);
      spreadsheetWriter.Write();

      using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithData1.xlsx"), FileMode.Create, FileAccess.Write);
      memoryStream.WriteTo(file);
    }
  }
}
