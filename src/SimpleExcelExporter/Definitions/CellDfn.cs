namespace SimpleExcelExporter.Definitions
{
  using System;

  public class CellDfn
  {
    public CellDfn(
      object value,
      CellDataType cellDataType = CellDataType.String,
      int index = 0)
    {
      CellDataType = cellDataType;
      Index = index;
      Value = value;
    }

    public CellDataType CellDataType { get; }

    /// <summary>
    /// Gets or set the value of the cell
    /// Value can be:
    /// - string
    /// - bool
    /// - DateTime
    /// - int32, int64, uint, double, float, etc.
    /// </summary>
    public object? Value { get; }

    public int Index { get; }

    public int GetStyleHashCode()
    {
      return HashCode.Combine((int)CellDataType);
    }
  }
}
