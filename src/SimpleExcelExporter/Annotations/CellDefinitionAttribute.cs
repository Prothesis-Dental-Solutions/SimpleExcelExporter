﻿namespace SimpleExcelExporter.Annotations
{
  using System;
  using SimpleExcelExporter.Definitions;

  [AttributeUsage(AttributeTargets.Property)]
  public sealed class CellDefinitionAttribute : Attribute
  {
    public CellDefinitionAttribute(
      CellDataType cellDataType)
    {
      CellDataType = cellDataType;
    }

    public CellDataType CellDataType { get; }
  }
}
