namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Resources;

  [TestFixture]
  public class EmptyResultMessageAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var resourceType = typeof(MessageRes);

      var emptyResultMessageAttribute = new EmptyResultMessageAttribute(resourceType, "EmptyMessageDefault");

      // Act & Check
      Assert.That(emptyResultMessageAttribute, Is.Not.Null);
      Assert.That(MessageRes.EmptyMessageDefault, Is.EqualTo(emptyResultMessageAttribute.Text));
    }
  }
}
