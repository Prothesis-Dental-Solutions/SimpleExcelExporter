namespace SimpleExcelExporter.Definitions
{
  using System.Collections.Generic;

  public class WorkbookDfn
  {
    public ICollection<WorksheetDfn> Worksheets { get; } = new HashSet<WorksheetDfn>();
  }
}
