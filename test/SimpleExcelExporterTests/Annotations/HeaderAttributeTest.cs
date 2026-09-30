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
      using (Assert.EnterMultipleScope())
      {
        Assert.That(headerAttribute.Text, Is.EqualTo(PlayerDummyObjectRes.PlayerNameColumnName));
        Assert.That(headerAttribute.ResourceName, Is.EqualTo("PlayerNameColumnName"));
        Assert.That(headerAttribute.ResourceType, Is.EqualTo(typeof(PlayerDummyObjectRes)));
      }
    }
  }
}
