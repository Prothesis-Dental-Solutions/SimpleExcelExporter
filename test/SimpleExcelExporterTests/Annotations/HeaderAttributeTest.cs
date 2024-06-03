namespace SimpleExcelExporter.Tests.Annotations
{
  using System;
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
      Type resourceType = typeof(PlayerDummyObjectRes);

      var headerAttribute = new HeaderAttribute(resourceType, "PlayerNameColumnName");

      // Act & Check
      Assert.IsNotNull(headerAttribute);
      Assert.AreEqual(PlayerDummyObjectRes.PlayerNameColumnName, headerAttribute.Text);
      Assert.AreEqual("PlayerNameColumnName", headerAttribute.ResourceName);
      Assert.AreEqual(typeof(PlayerDummyObjectRes), headerAttribute.ResourceType);
    }
  }
}
