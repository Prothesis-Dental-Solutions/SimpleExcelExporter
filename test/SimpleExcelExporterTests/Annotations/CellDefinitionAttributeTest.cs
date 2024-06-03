namespace SimpleExcelExporter.Tests.Annotations
{
  using NUnit.Framework;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;

  [TestFixture]
  public class CellDefinitionAttributeTest
  {
    [Test]
    public void ConstructorTest()
    {
      // Prepare
      var cellDefinitionAttribute = new CellDefinitionAttribute(CellDataType.Boolean);

      // Act & Check
      Assert.IsNotNull(cellDefinitionAttribute);
      Assert.AreEqual(CellDataType.Boolean, cellDefinitionAttribute.CellDataType);
    }
  }
}
