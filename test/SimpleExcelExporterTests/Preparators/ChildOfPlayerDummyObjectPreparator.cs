namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class ChildOfPlayerDummyObjectPreparator
  {
    public static ChildOfPlayerDummyObject First() => new()
    {
      Age = 11,
      FirstName = "FirstName 1",
      HeaderMention = "Old"
    };

    public static ChildOfPlayerDummyObject Second() => new()
    {
      Age = 12,
      FirstName = "FirstName 2",
      Gender = null
    };

    public static ChildOfPlayerDummyObject Third() => new()
    {
      Age = 13,
      FirstName = "FirstName 3",
      Gender = "Female"
    };

    public static ChildOfPlayerDummyObject Fourth() => new()
    {
      Age = 14,
      FirstName = "FirstName 4",
      Gender = "Male"
    };
  }
}
