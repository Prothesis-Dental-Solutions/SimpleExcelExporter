namespace SimpleExcelExporter.Tests
{
  using NUnit.Framework;

  [TestFixture]
  public class XmlStringHelperTest
  {
    [Test]
    public void SanitizeTest()
    {
      // Act
      string value = XmlStringHelper.Sanitize("|\b|\n|\t|\r|<|>|&|'|\"|");

      // Check
      Assert.AreEqual("| |\n|\t|\r|<|>|&|'|\"|", value);
    }
  }
}
