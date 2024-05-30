namespace SimpleExcelExporter.Definitions
{
  using System.Collections.Generic;
  using System.Linq;

  public class RowDfn
  {
    private static readonly CellDfnComparer CellDfnComparer = new CellDfnComparer();

    public ICollection<CellDfn> Cells { get; private set; } = new HashSet<CellDfn>();

    public void OrderCells()
    {
      Cells = Cells.OrderBy(c => c, CellDfnComparer).ToHashSet();
    }
  }
}
