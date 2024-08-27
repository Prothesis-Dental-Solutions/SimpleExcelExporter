namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Tests.Models;

  [TestFixture]
  public class HeaderAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var resourceType = typeof(PlayerDummyObjectRes);

      var headerAttribute = new HeaderAttribute(resourceType, "PlayerNameColumnName");

      // Act & Check
      Assert.That(headerAttribute, Is.Not.Null);
      Assert.That(PlayerDummyObjectRes.PlayerNameColumnName, Is.EqualTo(headerAttribute.Text));
      Assert.That("PlayerNameColumnName", Is.EqualTo(headerAttribute.ResourceName));
      Assert.That(typeof(PlayerDummyObjectRes), Is.EqualTo(headerAttribute.ResourceType));
    }
  }
}
