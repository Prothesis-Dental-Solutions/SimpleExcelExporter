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
      Assert.IsNotNull(multiColumnAttribute);
      Assert.AreEqual(multiColumnAttribute.MinimalNumberOfElement, 0);
      Assert.AreEqual(multiColumnAttribute.MaxNumberOfElement, 0);
    }


    [Test]
    public void ConstructorTest_MinimalNumberOfElement()
    {
      // Prepare
      var multiColumnAttribute = new MultiColumnAttribute(6);

      // Act & Check
      Assert.IsNotNull(multiColumnAttribute);
      Assert.AreEqual(multiColumnAttribute.MinimalNumberOfElement, 6);
      Assert.AreEqual(multiColumnAttribute.MaxNumberOfElement, 6);
    }
  }
}
