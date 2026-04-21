namespace SimpleExcelExporter
{
  using System;
  using System.Text;

  /// <summary>
  /// Converts 1-indexed column numbers to their A1 spreadsheet notation.
  /// </summary>
  public static class ColumnReferenceHelper
  {
    /// <summary>
    /// Converts a 1-indexed column number to its A1 letter notation.
    /// </summary>
    /// <param name="columnIndex">The 1-indexed column number (1 → "A", 27 → "AA").</param>
    /// <returns>The column letters.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="columnIndex"/> is less than 1.</exception>
    public static string ToLetters(int columnIndex)
    {
      if (columnIndex < 1)
      {
        throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, "Column index must be 1 or greater.");
      }

      var builder = new StringBuilder();
      var remaining = columnIndex;
      while (remaining > 0)
      {
        var modulo = (remaining - 1) % 26;
        builder.Insert(0, (char)('A' + modulo));
        remaining = (remaining - 1) / 26;
      }

      return builder.ToString();
    }
  }
}
