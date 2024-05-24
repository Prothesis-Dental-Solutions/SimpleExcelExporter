namespace ConsoleApp
{
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class Child
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
