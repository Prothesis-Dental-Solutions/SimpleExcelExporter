namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class EmptyResultMessageAttribute : ResourceBaseAttribute
  {
    public EmptyResultMessageAttribute(Type resourceType, string emptyResultMessage)
      : base(resourceType, emptyResultMessage)
    {
    }
  }
}
