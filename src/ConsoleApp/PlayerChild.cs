namespace ConsoleApp
{
  using System;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class PlayerChild
  {
    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerChildRes), "ChildFirstName")]
    [Index(1)]
    public string? FirstName { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerChildRes), "ChildAge")]
    [Index(2)]
    public int? Age { get; set; }
  }
}
