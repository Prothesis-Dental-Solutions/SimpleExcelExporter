namespace SimpleExcelExporter.Annotations
{
  using System;

  [AttributeUsage(AttributeTargets.Property)]
  public class ColumnTypeAttribute : Attribute
  {
    public ColumnTypeAttribute(
      ColumnType columnType)
    {
      ColumnType = columnType;
    }

    public ColumnType ColumnType { get; }

    // TODO Yanal - mettre ici le nombre d'éléments en facultatif et sera rempli par un premier parcours de tous les objets.
    public int MaxNumberOfElement { get; set; }
  }
}
