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
      var value = XmlStringHelper.Sanitize("|\b|\n|\t|\r|<|>|&|'|\"|");

      // Check
      Assert.That(value, Is.EqualTo("| |\n|\t|\r|<|>|&|'|\"|"));
    }
  }
}
