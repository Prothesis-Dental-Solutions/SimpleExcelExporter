namespace SimpleExcelExporter.Tests.Models
{
  using System.Collections.Generic;
  using SimpleExcelExporter.Annotations;

  public class TeamWithSameColumnIndexDummyObject
  {
    private ICollection<PlayerWithSameColumnIndexDummyObject>? _playerWithSameColumnIndexDummyObjects;

    [SheetName(typeof(TeamWithSameColumnIndexDummyObjectRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamWithSameColumnIndexDummyObjectRes), "EmptyResultMessage")]
    public ICollection<PlayerWithSameColumnIndexDummyObject> PlayerWithSameColumnIndexDummyObjects
    {
      get => _playerWithSameColumnIndexDummyObjects ??= new HashSet<PlayerWithSameColumnIndexDummyObject>();
    }
  }
}
