namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class IndexAttribute : Attribute
  {
    public IndexAttribute(int index)
    {
      if (index < 0)
      {
        throw new InvalidOperationException("Order shouldn't be negative.");
      }

      Index = index;
    }

    public int Index { get; }
  }
}
