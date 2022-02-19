namespace SimpleExcelExporter.Tests.Models
{
  using System.Collections.Generic;
  using SimpleExcelExporter.Annotations;

  public class TeamDummyObject
  {
    private ICollection<PlayerDummyObject>? _players;

    [SheetName(typeof(TeamDummyObjectRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamDummyObjectRes), "EmptyResultMessage")]
    public ICollection<PlayerDummyObject> Players
    {
      get => _players ??= new HashSet<PlayerDummyObject>();
    }
  }
}
