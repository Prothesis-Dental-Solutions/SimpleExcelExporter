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
      Assert.That(ignoreFromSpreadSheetAttribute, Is.Not.Null);
      Assert.That(true, Is.EqualTo(ignoreFromSpreadSheetAttribute.IgnoreFlag));
    }
  }
}
