namespace SimpleExcelExporter.Tests.Annotations
{
  using System;
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Resources;

  [TestFixture]
  public class EmptyResultMessageAttributeTest
  {
    [Test]
    public void Test()
    {
      // Prepare
      Type resourceType = typeof(MessageRes);

      var emptyResultMessageAttribute = new EmptyResultMessageAttribute(resourceType, "EmptyMessageDefault");

      // Act & Check
      Assert.IsNotNull(emptyResultMessageAttribute);
      Assert.AreEqual(MessageRes.EmptyMessageDefault, emptyResultMessageAttribute.Text);
    }
  }
}
