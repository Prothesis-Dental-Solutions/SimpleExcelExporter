namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class TeamWithSameSheetNameDummyObjectPreparator
  {
    public static TeamWithSameSheetNameDummyObject First() => new()
    {

      FirstPlayerDummyObjects =
        {
          PlayerDummyObjectPreparator.First(),
        },
      SecondPlayerDummyObjects =
        {
          PlayerDummyObjectPreparator.Second(),
        },
    };
  }
}
