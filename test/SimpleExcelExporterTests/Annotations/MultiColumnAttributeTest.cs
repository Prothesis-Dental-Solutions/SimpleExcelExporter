namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;

  [TestFixture]
  public class MultiColumnAttributeTest
  {
    [Test]
    public void Test()
    {
      // Prepare
      var columnTypeAttribute = new MultiColumnAttribute();

      // Act & Check
      Assert.IsNotNull(columnTypeAttribute);
    }
  }
}
