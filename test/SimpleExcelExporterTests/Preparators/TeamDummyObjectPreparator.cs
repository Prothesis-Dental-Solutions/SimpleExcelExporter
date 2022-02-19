namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class TeamDummyObjectPreparator
  {
    public static TeamDummyObject First() => new TeamDummyObject();

    public static TeamDummyObject FirstWithCollections() => new TeamDummyObject
    {
      Players =
        {
          PlayerDummyObjectPreparator.First(),
          PlayerDummyObjectPreparator.Second(),
          PlayerDummyObjectPreparator.Third(),
          PlayerDummyObjectPreparator.Fourth(),
        },
    };
  }
}
