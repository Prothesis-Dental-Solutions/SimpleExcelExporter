namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class ChildOfPlayerDummyObjectPreparator
  {
    public static ChildOfPlayerDummyObject First() => new()
    {
      Age = 11,
      FirstName = "FirstName 1"
    };

    public static ChildOfPlayerDummyObject Second() => new()
    {
      Age = 12,
      FirstName = "FirstName 2"
    };

    public static ChildOfPlayerDummyObject Third() => new()
    {
      Age = 13,
      FirstName = "FirstName 3"
    };

    public static ChildOfPlayerDummyObject Fourth() => new()
    {
      Age = 14,
      FirstName = "FirstName 4"
    };
  }
}
