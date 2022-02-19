namespace SimpleExcelExporter.Tests.Annotations
{
  using System;
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Tests.Models;

  [TestFixture]
  public class SheetNameAttributeTest
  {
    [Test]
    public void Test()
    {
      // Prepare
      Type resourceType = typeof(PlayerDummyObjectRes);

      var sheetNameAttribute = new SheetNameAttribute(resourceType, "PlayerNameColumnName");

      // Act & Check
      Assert.IsNotNull(sheetNameAttribute);
      Assert.AreEqual(PlayerDummyObjectRes.PlayerNameColumnName, sheetNameAttribute.Text);
    }
  }
}
