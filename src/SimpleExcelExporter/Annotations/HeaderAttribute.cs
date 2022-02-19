namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class HeaderAttribute : ResourceBaseAttribute
  {
    public HeaderAttribute(Type resourceType, string resourceName)
      : base(resourceType, resourceName)
    {
    }
  }
}
