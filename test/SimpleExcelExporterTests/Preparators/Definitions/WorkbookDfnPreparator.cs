namespace SimpleExcelExporter.Tests.Preparators.Definitions
{
  using System;
  using SimpleExcelExporter.Definitions;

  /// <summary>
  /// This class is able to prepare dummy <see cref="WorkbookDfn"/> for test purposes.
  /// </summary>
  public static class WorkbookDfnPreparator
  {
    public static WorkbookDfn First() => new();

    public static WorkbookDfn FirstFirstWithCollections()
    {
      var workbookDfn = new WorkbookDfn();

      // First sheet
      var worksheet1Dfn = new WorksheetDfn("MyFirstSheet");
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Name"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Age"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Rate"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("Postal code"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("DateTime"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("FieldGoalPercentage"));
      worksheet1Dfn.ColumnHeadings.Cells.Add(new CellDfn("PracticeTime"));
      workbookDfn.Worksheets.Add(worksheet1Dfn);
      var row1 = new RowDfn();
      row1.Cells.Add(new CellDfn("Eric", cellDataType: CellDataType.String));
      row1.Cells.Add(new CellDfn(50, cellDataType: CellDataType.Number));
      row1.Cells.Add(new CellDfn(45.00M, cellDataType: CellDataType.Number));
      row1.Cells.Add(new CellDfn("01090", cellDataType: CellDataType.String));
      row1.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row1.Cells.Add(new CellDfn(0.1111, cellDataType: CellDataType.Percentage));
      row1.Cells.Add(new CellDfn(new TimeSpan(9, 1, 0), cellDataType: CellDataType.Time));
      worksheet1Dfn.Rows.Add(row1);
      var row2 = new RowDfn();
      row2.Cells.Add(new CellDfn("Bob", cellDataType: CellDataType.String));
      row2.Cells.Add(new CellDfn(42, cellDataType: CellDataType.Number));
      row2.Cells.Add(new CellDfn(78.00M, cellDataType: CellDataType.Number));
      row2.Cells.Add(new CellDfn("01080", cellDataType: CellDataType.String));
      row2.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row2.Cells.Add(new CellDfn(0.2222, cellDataType: CellDataType.Percentage));
      row2.Cells.Add(new CellDfn(new TimeSpan(9, 2, 0), cellDataType: CellDataType.Time));
      worksheet1Dfn.Rows.Add(row2);
      var row3 = new RowDfn();
      row3.Cells.Add(new CellDfn("Bob", cellDataType: CellDataType.String));
      row3.Cells.Add(new CellDfn(42, cellDataType: CellDataType.Number));
      row3.Cells.Add(new CellDfn(78.00M, cellDataType: CellDataType.Number));
      row3.Cells.Add(new CellDfn("01080", cellDataType: CellDataType.String));
      row3.Cells.Add(new CellDfn(DateTime.Now, cellDataType: CellDataType.Date));
      row3.Cells.Add(new CellDfn(0.2222, cellDataType: CellDataType.Percentage));
      row3.Cells.Add(new CellDfn(string.Empty, cellDataType: CellDataType.String));
      row3.Cells.Add(new CellDfn(new TimeSpan(9, 2, 0), cellDataType: CellDataType.Time));
      worksheet1Dfn.Rows.Add(row3);

      return workbookDfn;
    }
  }
}
