namespace SimpleExcelExporter.Tests.Models
{
  using System;
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
    [Index(2)]
    public int? Age { get; set; }
  }
}
