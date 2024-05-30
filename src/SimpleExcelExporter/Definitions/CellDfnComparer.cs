namespace SimpleExcelExporter.Definitions;

using System;
using System.Collections.Generic;

public class CellDfnComparer : IComparer<CellDfn>
{
  public int Compare(CellDfn? x, CellDfn? y)
  {
    if (x?.Index == null)
    {
      return -1;
    }

    if (y?.Index == null)
    {
      return 1;
    }

    if (x == y || x.Index == y.Index)
    {
      return 0;
    }

    int xCount = x.Index.Count;
    int yCount = y.Index.Count;

    if (xCount == 0 && yCount == 0)
    {
      return 0;
    }

    if (xCount == 0)
    {
      return -1;
    }

    if (yCount == 0)
    {
      return 1;
    }

    int minCount = Math.Min(xCount, yCount);

    for (int i = 0; i < minCount; i++)
    {
      if (x.Index[i] < y.Index[i])
      {
        return -1;
      }
      else if (x.Index[i] > y.Index[i])
      {
        return 1;
      }
    }

    return 0;
  }
}
