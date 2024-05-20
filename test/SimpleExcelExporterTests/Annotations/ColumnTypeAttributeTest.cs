namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  [TestFixture]
  public class ColumnTypeAttributeTest
  {
    [Test]
    public void Test()
    {
      // Prepare
      var columnTypeAttribute = new ColumnTypeAttribute(ColumnType.Basic);

      // Act & Check
      Assert.IsNotNull(columnTypeAttribute);
      Assert.AreEqual(ColumnType.Basic, columnTypeAttribute.ColumnType);
    }
  }
}
