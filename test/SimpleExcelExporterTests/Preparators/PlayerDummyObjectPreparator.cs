namespace SimpleExcelExporter.Tests.Preparators
{
  using System;
  using System.Collections.Generic;
  using SimpleExcelExporter.Tests.Models;

  public static class PlayerDummyObjectPreparator
  {
    public static PlayerDummyObject First() => new()
    {
      PlayerCode = null,
      PlayerName = "Player\bName1<a href=\"https://www.google.com\" /> &lt;b /&gt; \r\n\t",
      PracticeTime = new TimeSpan(9, 1, 0),
      Size = 1.93d,
      DateOfBirth = new DateTime(1974, 02, 01),
      IsActiveFlag = true,
      NumberOfVictory = 45,
      Salary = 2000.5m,
      ByteColumn = 1,
      DateTimeOffsetColumn = new DateTimeOffset(new DateTime(1974, 02, 01)),
      FieldGoalPercentage = 0.0111d,
    };

    public static PlayerDummyObject Second() => new()
    {
      PlayerCode = "02",
      PlayerName = "PlayerName2",
      PracticeTime = new TimeSpan(9, 2, 0),
      Size = 1.72d,
      DateOfBirth = new DateTime(1990, 10, 13),
      IsActiveFlag = true,
      NumberOfVictory = 52,
      Salary = 2141.5452m,
      ByteColumn = 2,
      DateTimeOffsetColumn = new DateTimeOffset(new DateTime(1990, 10, 13)),
      FieldGoalPercentage = 0.0222d,
      Children = new List<ChildOfPlayerDummyObject> {
        ChildOfPlayerDummyObjectPreparator.Third(),
        ChildOfPlayerDummyObjectPreparator.Fourth(),
      }
    };

    public static PlayerDummyObject Third() => new()
    {
      PlayerCode = "03",
      PlayerName = "PlayerName3",
      PracticeTime = new TimeSpan(9, 3, 0),
      Size = 1.85d,
      DateOfBirth = new DateTime(1976, 3, 1),
      IsActiveFlag = true,
      NumberOfVictory = 80,
      Salary = 2111.5452m,
      ByteColumn = 3,
      DateTimeOffsetColumn = new DateTimeOffset(new DateTime(1976, 3, 1)),
      FieldGoalPercentage = 0.0333d,
      Children = new List<ChildOfPlayerDummyObject> {
        ChildOfPlayerDummyObjectPreparator.First(),
        ChildOfPlayerDummyObjectPreparator.Fourth(),
      }
    };

    public static PlayerDummyObject Fourth() => new()
    {
      PlayerCode = "04",
      PlayerName = "PlayerName4",
      PracticeTime = new TimeSpan(9, 4, 0),
      Size = 1.79d,
      DateOfBirth = new DateTime(1979, 3, 1),
      IsActiveFlag = false,
      NumberOfVictory = 35,
      Salary = 2845.719m,
      ByteColumn = 4,
      DateTimeOffsetColumn = new DateTimeOffset(new DateTime(1979, 3, 1)),
      FieldGoalPercentage = 0.0444d,
      Children = new List<ChildOfPlayerDummyObject> {
        ChildOfPlayerDummyObjectPreparator.Second(),
        ChildOfPlayerDummyObjectPreparator.Third(),
      }
    };

    public static PlayerDummyObject Fifth() => new()
    {
      Children = null,
    };

    public static PlayerDummyObject Sixth() => new();

    public static PlayerDummyObject Seventh() => new()
    {
      PlayerCode = null,
      PlayerName = "Player\bName1<a href=\"https://www.google.com\" /> &lt;b /&gt; \r\n\t",
      PracticeTime = new TimeSpan(9, 1, 0),
      Size = 1.93d,
      DateOfBirth = new DateTime(1974, 02, 01),
      IsActiveFlag = true,
      NumberOfVictory = 45,
      Salary = 2000.5m,
      ByteColumn = 1,
      DateTimeOffsetColumn = new DateTimeOffset(new DateTime(1974, 02, 01)),
      FieldGoalPercentage = 0.0111d,
      Children = new List<ChildOfPlayerDummyObject> {
        ChildOfPlayerDummyObjectPreparator.First(),
        ChildOfPlayerDummyObjectPreparator.Second(),
        ChildOfPlayerDummyObjectPreparator.Third(),
      }
    };
  }
}
