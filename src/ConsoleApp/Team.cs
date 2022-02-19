namespace ConsoleApp
{
  using System.Collections.Generic;
  using SimpleExcelExporter.Annotations;

  public class Team
  {
    private ICollection<Player>? _players;

    [SheetName(typeof(TeamRes), "SheetName")]
    [EmptyResultMessage(typeof(TeamRes), "EmptyResultMessage")]
    public ICollection<Player> Players
    {
      get => _players ??= new HashSet<Player>();
    }
  }
}
