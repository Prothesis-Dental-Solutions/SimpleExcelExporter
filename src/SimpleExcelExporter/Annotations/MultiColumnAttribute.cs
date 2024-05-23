namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public sealed class MultiColumnAttribute : Attribute
  {
    public MultiColumnAttribute(int maxNumberOfElement = 0)
    {
      MaxNumberOfElement = maxNumberOfElement;
    }

    public int MaxNumberOfElement { get; set; }
  }
}
