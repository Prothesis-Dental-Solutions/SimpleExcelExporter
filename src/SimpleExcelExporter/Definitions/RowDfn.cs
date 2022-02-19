namespace SimpleExcelExporter.Definitions
{
  using System.Collections.Generic;
  using System.Linq;

  public class RowDfn
  {
    public ICollection<CellDfn> Cells { get; private set; } = new HashSet<CellDfn>();

    public void OrderCells()
    {
      Cells = Cells.OrderBy(c => c.Index).ToHashSet();
    }
  }
}
