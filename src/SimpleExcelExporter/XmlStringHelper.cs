namespace SimpleExcelExporter
{
  using System.Linq;
  using System.Xml;

  public static class XmlStringHelper
  {
    public static string Sanitize(string input)
    {
      return string.Concat(input.Where(c => XmlConvert.IsXmlChar(c)));
    }
  }
}
