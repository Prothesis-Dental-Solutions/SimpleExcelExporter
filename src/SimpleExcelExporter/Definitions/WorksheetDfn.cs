namespace SimpleExcelExporter.Definitions
{
  using System.Collections.Generic;

  public class WorksheetDfn
  {
    public WorksheetDfn(string name)
    {
      Name = name;
    }

    public string Name { get; }

    public RowDfn ColumnHeadings { get; } = new();

    public ICollection<RowDfn> Rows { get; } = new HashSet<RowDfn>();
  }
}
