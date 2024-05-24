namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public sealed class MultiColumnAttribute : Attribute
  {
    public MultiColumnAttribute(int minimalNumberOfElement = 0)
    {
      MaxNumberOfElement = minimalNumberOfElement;
      MinimalNumberOfElement = minimalNumberOfElement;
    }

    public int MaxNumberOfElement { get; set; }

    public int MinimalNumberOfElement { get; }
  }
}
