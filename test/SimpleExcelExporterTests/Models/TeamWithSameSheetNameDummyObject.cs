namespace SimpleExcelExporter.Tests.Models
{
  using System.Collections.Generic;
  using SimpleExcelExporter.Annotations;

  public class TeamWithSameSheetNameDummyObject
  {
    private ICollection<PlayerDummyObject>? _firstPlayerDummyObjects;

    [SheetName(typeof(TeamWithSameSheetNameDummyObjectRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamWithSameSheetNameDummyObjectRes), "EmptyResultMessage")]
    public ICollection<PlayerDummyObject> FirstPlayerDummyObjects
    {
      get => _firstPlayerDummyObjects ??= new HashSet<PlayerDummyObject>();
    }

    private ICollection<PlayerDummyObject>? _secondPlayerDummyObjects;

    [SheetName(typeof(TeamWithSameSheetNameDummyObjectRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamWithSameSheetNameDummyObjectRes), "EmptyResultMessage")]
    public ICollection<PlayerDummyObject> SecondPlayerDummyObjects
    {
      get => _secondPlayerDummyObjects ??= new HashSet<PlayerDummyObject>();
    }

  }
}
