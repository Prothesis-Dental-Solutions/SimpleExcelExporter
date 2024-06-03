namespace ConsoleApp;

using System.Collections.Generic;
using SimpleExcelExporter.Annotations;
using SimpleExcelExporter.Definitions;

public class Person
{
  [CellDefinition(CellDataType.String)]
  [Header(typeof(PlayerRes), "PlayerNameColumnName")]
  [Index(1)]
  public string? Name { get; set; }

  [MultiColumn]
  [Index(2)]
  public ICollection<Person> Children { get; } = new List<Person>();
}
