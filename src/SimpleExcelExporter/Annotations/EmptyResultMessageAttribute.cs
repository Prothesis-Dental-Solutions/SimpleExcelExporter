namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public sealed class EmptyResultMessageAttribute : ResourceBaseAttribute
  {
    public EmptyResultMessageAttribute(Type resourceType, string emptyResultMessage)
      : base(resourceType, emptyResultMessage)
    {
      EmptyResultMessage = emptyResultMessage;
    }

    public string EmptyResultMessage { get; private init; }
  }
}
