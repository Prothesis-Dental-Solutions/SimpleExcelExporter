namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class TeamWithSameColumnIndexDummyObjectPreparator
  {
    public static TeamWithSameColumnIndexDummyObject First() => new();

    public static TeamWithSameColumnIndexDummyObject FirstWithCollections() => new()
    {
      PlayerWithSameColumnIndexDummyObjects =
        {
          PlayerWithSameColumnIndexDummyObjectPreparator.First(),
          PlayerWithSameColumnIndexDummyObjectPreparator.Second(),
          PlayerWithSameColumnIndexDummyObjectPreparator.Third(),
        },
    };
  }
}
