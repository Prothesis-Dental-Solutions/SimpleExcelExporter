
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
      int expectedIndex = 1;
      var indexAttribute = new IndexAttribute(expectedIndex);

      //Act && Check
      Assert.IsNotNull(indexAttribute.Index);
      Assert.AreEqual(indexAttribute.Index, expectedIndex);

    }
  }
}
