namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class IgnoreFromSpreadSheetAttribute : Attribute
  {
    public IgnoreFromSpreadSheetAttribute(bool ignoreFlag = true)
    {
      IgnoreFlag = ignoreFlag;
    }

    public bool IgnoreFlag { get; }
  }
}
