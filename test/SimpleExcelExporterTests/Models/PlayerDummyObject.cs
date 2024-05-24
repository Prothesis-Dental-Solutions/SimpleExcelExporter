namespace SimpleExcelExporter.Tests.Models
{
  using System;
  using System.Collections.Generic;
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

    [CellDefinition(CellDataType.Time)]
    [Header(typeof(PlayerDummyObjectRes), "PracticeTimeColumnName")]
    [Index(10)]
    public TimeSpan? PracticeTime { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "NumberOfVictoryColumnName")]
    [Index(3)]
    public int? NumberOfVictory { get; set; }

    [CellDefinition(CellDataType.Boolean)]
    [Header(typeof(PlayerDummyObjectRes), "IsActiveFlagColumnName")]
    [Index(4)]
    public bool? IsActiveFlag { get; set; }

    [CellDefinition(CellDataType.Percentage)]
    [Header(typeof(PlayerDummyObjectRes), "FieldGoalPercentageColumnName")]
    [Index(11)]
    public double? FieldGoalPercentage { get; set; }

    [CellDefinition(CellDataType.Boolean)]
    [Header(typeof(PlayerDummyObjectRes), "SizeColumnName")]
    [Index(5)]
    public double? Size { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "SalaryColumnName")]
    [Index(6)]
    public decimal? Salary { get; set; }

    [MultiColumn]
    [Index(9)]
    public ICollection<ChildOfPlayerDummyObject>? Children { get; set; }

    [CellDefinition(CellDataType.Number)]
    [Header(typeof(PlayerDummyObjectRes), "ByteColumnName")]
    [Index(7)]
    public byte? ByteColumn { get; set; }

    [CellDefinition(CellDataType.String)]
    [Header(typeof(PlayerDummyObjectRes), "DateTimeOffsetColumnName")]
    [Index(8)]
    public DateTimeOffset? DateTimeOffsetColumn { get; set; }
  }
}
