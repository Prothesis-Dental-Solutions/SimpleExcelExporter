
namespace SimpleExcelExporter.Tests.Annotations
{
  using System;
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;

  public class IndexAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      //Prepare && Act & Check
      // ReSharper disable once ObjectCreationAsStatement
      Assert.Throws<InvalidOperationException>(() => new IndexAttribute(-1));

      // Prepare
      var expectedIndex = 1;
      var indexAttribute = new IndexAttribute(expectedIndex);

      //Act && Check
      Assert.That(indexAttribute.Index, Is.Not.Null);
      Assert.That(indexAttribute.Index, Is.EqualTo(expectedIndex));
    }
  }
}
