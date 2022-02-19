namespace SimpleExcelExporter.Tests.Models
{
  using System;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  public class PlayerDummyObject
  {
    [CellDefinition(CellDataType.Date)]
    [Header(typeof(PlayerDummyObjectRes), "DateOfBirthColumnName")]
    [Index(2)]
    public DateTime? DateOfBirth { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerDummyObjectRes), "PlayerCodeColumnName")]
    [Index(0)]
    public string? PlayerCode { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerDummyObjectRes), "PlayerNameColumnName")]
    [Index(1)]
    public string? PlayerName { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "NumberOfVictoryColumnName")]
    [Index(3)]
    public int? NumberOfVictory { get; set; }

    [CellDefinition(CellDataType.Boolean)]
    [Header(typeof(PlayerDummyObjectRes), "IsActiveFlagColumnName")]
    [Index(4)]
    public bool? IsActiveFlag { get; set; }

    [CellDefinition(CellDataType.Boolean)]
    [Header(typeof(PlayerDummyObjectRes), "SizeColumnName")]
    [Index(5)]
    public double? Size { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "SalaryColumnName")]
    [Index(5)]
    public decimal? Salary { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "ByteColumnName")]
    [Index(6)]
    public byte? ByteColumn { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerDummyObjectRes), "DateTimeOffsetColumnName")]
    [Index(7)]
    public DateTimeOffset? DateTimeOffsetColumn { get; set; }
  }
}
