namespace ConsoleApp;

using System.Collections.Generic;
using SimpleExcelExporter.Annotations;

public class Group
{
  private ICollection<Person>? _persons;

  [SheetName(typeof(TeamRes), "SheetName")]
  [EmptyResultMessage(typeof(TeamRes), "EmptyResultMessage")]
  public ICollection<Person> Persons
  {
    get => _persons ??= new HashSet<Person>();
  }
}
