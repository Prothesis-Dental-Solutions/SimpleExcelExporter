namespace SimpleExcelExporter
{
  using System.Text;
  using System.Xml;

  public static class XmlStringHelper
  {
    public static string Sanitize(string input)
    {
      StringBuilder sb = new StringBuilder();

      foreach (char c in input)
      {
        if (XmlConvert.IsXmlChar(c))
        {
          sb.Append(c);
        }
        else
        {
          sb.Append(' ');
        }
      }

      return sb.ToString();
    }
  }
}
