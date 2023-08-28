namespace SimpleExcelExporter.Definitions
{
  public enum CellDataType
  {
    /// <summary>
    /// Boolean data type
    /// </summary>
    Boolean,

    /// <summary>
    /// Date data type
    /// </summary>
    Date,

    /// <summary>
    /// Number data type
    /// </summary>
    Number,

    /// <summary>
    /// Percentage data type
    /// </summary>
    Percentage,

    /// <summary>
    /// String data type
    /// </summary>
#pragma warning disable CA1720 // Identifier contains type name
    String,
#pragma warning restore CA1720 // Identifier contains type name

    /// <summary>
    /// Time data type
    /// </summary>
    Time,

  }
}
