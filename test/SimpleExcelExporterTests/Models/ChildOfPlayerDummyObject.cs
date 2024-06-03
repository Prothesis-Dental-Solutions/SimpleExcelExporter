namespace SimpleExcelExporter.Tests.Models
{
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class ChildOfPlayerDummyObject
  {
    [CellDefinition(CellDataType.String)]
    [Header(typeof(ChildOfPlayerDummyObjectRes), "ChildFirstNameColumnName")]
    [Index(1)]
    public string? FirstName { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(ChildOfPlayerDummyObjectRes), "ChildAgeColumnName")]
    [Index(3)]
    public int? Age { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(ChildOfPlayerDummyObjectRes), "ChildGender", nameof(HeaderMention))]
    [Index(2)]
    public string? Gender { get; set; }

    [IgnoreFromSpreadSheet]
    public string? HeaderMention { get; set; }
  }
}
