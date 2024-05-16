namespace SimpleExcelExporter.Annotations
{
  using System;
  using SimpleExcelExporter.Definitions;

  [AttributeUsage(AttributeTargets.Property)]
  public class ColumnTypeAttribute : Attribute
  {
    public ColumnTypeAttribute(
      ColumnType columnType)
    {
      ColumnType = columnType;
    }

    public ColumnType ColumnType { get; }
  }
}
