namespace SimpleExcelExporter.Tests
{
  using System;
  using NUnit.Framework;

  [TestFixture]
  public class ColumnReferenceHelperTest
  {
    [TestCase(1, "A")]
    [TestCase(2, "B")]
    [TestCase(25, "Y")]
    [TestCase(26, "Z")]
    [TestCase(27, "AA")]
    [TestCase(28, "AB")]
    [TestCase(52, "AZ")]
    [TestCase(53, "BA")]
    [TestCase(702, "ZZ")]
    [TestCase(703, "AAA")]
    public void ToLetters_ConvertsOneIndexedColumnNumberToA1Letters(int columnIndex, string expected)
    {
      var result = ColumnReferenceHelper.ToLetters(columnIndex);

      Assert.That(result, Is.EqualTo(expected));
    }

    [Test]
    public void ToLetters_ThrowsForZeroOrNegative()
    {
      Assert.Throws<ArgumentOutOfRangeException>((Action)(() => ColumnReferenceHelper.ToLetters(0)));
      Assert.Throws<ArgumentOutOfRangeException>((Action)(() => ColumnReferenceHelper.ToLetters(-1)));
    }
  }
}
