namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public sealed class SheetNameAttribute : ResourceBaseAttribute
  {
    public SheetNameAttribute(Type resourceType, string resourceName)
      : base(resourceType, resourceName)
    {
    }
  }
}
