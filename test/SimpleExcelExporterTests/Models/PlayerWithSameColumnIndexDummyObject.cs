namespace SimpleExcelExporter.Tests.Models
{
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class PlayerWithSameColumnIndexDummyObject
  {
    [CellDefinition(CellDataType.Date)]
    [Header(typeof(PlayerWithSameColumnIndexDummyObjectRes), "FirstColumnName")]
    [Index(0)]
    public string? FirstColumn { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerWithSameColumnIndexDummyObjectRes), "SecondColumnName")]
    [Index(0)]
    public string? SecondColumn { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerWithSameColumnIndexDummyObjectRes), "ThirdColumnName")]
    [Index(1)]
    public string? ThirdColumn { get; set; }

    public string? FourthColumn { get; set; }
  }
}
