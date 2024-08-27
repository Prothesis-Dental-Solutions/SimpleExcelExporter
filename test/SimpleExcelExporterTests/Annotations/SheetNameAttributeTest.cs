namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Tests.Models;

  [TestFixture]
  public class SheetNameAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var resourceType = typeof(PlayerDummyObjectRes);

      var sheetNameAttribute = new SheetNameAttribute(resourceType, "PlayerNameColumnName");

      // Act & Check
      Assert.That(sheetNameAttribute, Is.Not.Null);
      Assert.That(PlayerDummyObjectRes.PlayerNameColumnName, Is.EqualTo(sheetNameAttribute.Text));
    }
  }
}
