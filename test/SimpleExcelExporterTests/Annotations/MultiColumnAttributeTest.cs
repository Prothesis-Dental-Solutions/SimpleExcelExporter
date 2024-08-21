namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;

  [TestFixture]
  public class MultiColumnAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var multiColumnAttribute = new MultiColumnAttribute();

      // Act & Check
      Assert.That(multiColumnAttribute, Is.Not.Null);
      Assert.That(multiColumnAttribute.MinimalNumberOfElement, Is.EqualTo(0));
      Assert.That(multiColumnAttribute.MaxNumberOfElement, Is.EqualTo(0));
    }


    [Test]
    public void ConstructorTest_MinimalNumberOfElement()
    {
      // Prepare
      var multiColumnAttribute = new MultiColumnAttribute(6);

      // Act & Check
      Assert.That(multiColumnAttribute, Is.Not.Null);
      Assert.That(multiColumnAttribute.MinimalNumberOfElement, Is.EqualTo(6));
      Assert.That(multiColumnAttribute.MaxNumberOfElement, Is.EqualTo(6));
    }
  }
}
