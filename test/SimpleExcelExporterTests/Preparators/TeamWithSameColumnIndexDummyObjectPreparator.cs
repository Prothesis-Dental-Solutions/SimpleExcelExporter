namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class TeamWithSameColumnIndexDummyObjectPreparator
  {
    public static TeamWithSameColumnIndexDummyObject First() => new TeamWithSameColumnIndexDummyObject();

    public static TeamWithSameColumnIndexDummyObject FirstWithCollections() => new TeamWithSameColumnIndexDummyObject
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
