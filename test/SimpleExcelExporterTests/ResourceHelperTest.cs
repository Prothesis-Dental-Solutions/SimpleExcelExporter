namespace SimpleExcelExporter.Tests
{
  using System;
  using NUnit.Framework;
  using SimpleExcelExporter.Resources;

  [TestFixture]
  public class ResourceHelperTest
  {
    [Test]
    public void GetResourceLookupTest()
    {
      // Prepare
      var resourceType = typeof(MessageRes);

      // Act && Check -- ResourceType invalid
      Assert.Throws<InvalidOperationException>(() => ResourceHelper.GetResourceLookup(typeof(string), string.Empty));

      // Act && Check -- Resource name does not exist
      Assert.Throws<InvalidOperationException>(() => ResourceHelper.GetResourceLookup(resourceType, "PropertyDoesNotExist"));

      // Act 
      var resultValue = ResourceHelper.GetResourceLookup(resourceType, "EmptyMessageDefault");

      // Check
      Assert.That(resultValue, Is.EqualTo(MessageRes.EmptyMessageDefault));
    }
  }
}
