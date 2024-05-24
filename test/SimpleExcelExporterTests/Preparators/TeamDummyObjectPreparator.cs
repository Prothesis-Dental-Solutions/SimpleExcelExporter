namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class TeamDummyObjectPreparator
  {
    public static TeamDummyObject First() => new();

    public static TeamDummyObject FirstWithCollections() => new()
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
