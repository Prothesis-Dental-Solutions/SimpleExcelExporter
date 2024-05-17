namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class HeaderAttribute : ResourceBaseAttribute
  {
    public HeaderAttribute(Type resourceType, string resourceName, string? textToAddToHeader = null)
      : base(resourceType, resourceName)
    {
      TextToAddToHeader = textToAddToHeader;
    }

    public string? TextToAddToHeader { get; }
  }
}
