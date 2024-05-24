namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;

  [TestFixture]
  public class IgnoreFromSpreadSheetAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var ignoreFromSpreadSheetAttribute = new IgnoreFromSpreadSheetAttribute(true);

      // Act & Check
      Assert.IsNotNull(ignoreFromSpreadSheetAttribute);
      Assert.AreEqual(true, ignoreFromSpreadSheetAttribute.IgnoreFlag);
    }
  }
}
