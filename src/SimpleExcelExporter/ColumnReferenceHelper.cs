namespace SimpleExcelExporter
{
  using System;
  using System.Text;

  /// <summary>
  /// Converts 1-indexed column numbers to their A1 spreadsheet notation.
  /// </summary>
  public static class ColumnReferenceHelper
  {
    // Precomputed A1 letters for columns 1..CacheSize. Hot path on the XLSX writer:
    // every cell goes through ToLetters, so the repeated StringBuilder allocations
    // measurably weighed on the Write phase.
    // 16384 is Excel's hard column cap — it covers every workbook that can exist.
    // Pre-allocates ~450 KB of interned strings once per process, trading that
    // static footprint for zero per-cell allocation on the write hot path regardless
    // of how wide the sheet is.
    private const int CacheSize = 16384;

    private static readonly string[] Cache = BuildCache();

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

      if (columnIndex <= CacheSize)
      {
        return Cache[columnIndex - 1];
      }

      return Compute(columnIndex);
    }

    private static string[] BuildCache()
    {
      var cache = new string[CacheSize];
      for (var i = 1; i <= CacheSize; i++)
      {
        cache[i - 1] = Compute(i);
      }

      return cache;
    }

    private static string Compute(int columnIndex)
    {
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
