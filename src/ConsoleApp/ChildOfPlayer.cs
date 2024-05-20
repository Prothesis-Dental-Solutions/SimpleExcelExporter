namespace ConsoleApp
{
  using System;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class ChildOfPlayer
  {
    [CellDefinition(CellDataType.String)]
    [Header(typeof(ChildOfPlayerRes), "ChildFirstName")]
    [Index(1)]
    public string? FirstName { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(ChildOfPlayerRes), "ChildAge")]
    [Index(2)]
    public int? Age { get; set; }
  }
}
