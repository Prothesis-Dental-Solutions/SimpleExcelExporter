namespace SimpleExcelExporter.Tests.Preparators
{
  using SimpleExcelExporter.Tests.Models;

  public static class PlayerWithSameColumnIndexDummyObjectPreparator
  {
    public static PlayerWithSameColumnIndexDummyObject First() => new()
    {
      FirstColumn = "01/01/2001",
      SecondColumn = "SecondColumn1",
      ThirdColumn = "ThirdColumn1",
      FourthColumn = "FourthColumn1",
    };

    public static PlayerWithSameColumnIndexDummyObject Second() => new()
    {
      FirstColumn = "02/02/2002",
      SecondColumn = "SecondColumn2",
      ThirdColumn = "ThirdColumn2",
      FourthColumn = "FourthColumn2",
    };

    public static PlayerWithSameColumnIndexDummyObject Third() => new()
    {
      FirstColumn = "03/03/2003",
      SecondColumn = "SecondColumn3",
      ThirdColumn = "ThirdColumn3",
      FourthColumn = "FourthColumn3",
    };
  }
}
