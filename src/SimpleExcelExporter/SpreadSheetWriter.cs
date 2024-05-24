﻿namespace SimpleExcelExporter
{
  using System;
  using System.Collections.Generic;
  using System.Globalization;
  using System.IO;
  using System.Linq;
  using System.Reflection;
  using System.Xml;
  using DocumentFormat.OpenXml;
  using DocumentFormat.OpenXml.Packaging;
  using DocumentFormat.OpenXml.Spreadsheet;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;
  using SimpleExcelExporter.Resources;

  public class SpreadsheetWriter
  {
    private readonly IDictionary<string, Attribute?> _cachedAttributes = new Dictionary<string, Attribute?>();

    private readonly ISet<string> _headers = new HashSet<string>();

    private readonly Stream _stream;

    private readonly Stylesheet _stylesheet;

    private readonly WorkbookDfn _workbookDfn;

    public SpreadsheetWriter(Stream stream, WorkbookDfn workbookDfn)
    {
      _stream = stream;
      _stylesheet = new Stylesheet();
      _workbookDfn = workbookDfn;
      OrderWorkBookDfn();
      Validate();
    }

    public SpreadsheetWriter(Stream stream, object team)
    {
      _stream = stream;
      _stylesheet = new Stylesheet();
      _workbookDfn = BuildWorkbook(team);
      OrderWorkBookDfn();
      Validate();
    }

    private IDictionary<int, uint> Table { get; } = new Dictionary<int, uint>();

    public void Write()
    {
      using var document = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);

      // Adding core file properties is mandatory to avoid a problem with Google Spreadsheet transforming from XLSX to XLSM.
      // cf. https://stackoverflow.com/questions/70319867/avoid-google-spreadsheet-to-convert-an-xlsx-file-created-by-open-xml-sdk-to-xlsm
      // cf. https://github.com/OfficeDev/Open-XML-SDK/issues/1093
      // cf. https://issuetracker.google.com/issues/210875597
      var coreFilePropPart = document.AddCoreFilePropertiesPart();
      using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
      {
        writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
        writer.Flush();
      }

      CreatePartsForExcel(document);
    }

    private static void AddHeaderCellToWorkSheet(WorksheetDfn worksheetDfn, string text, decimal index)
    {
      CellDfn headerCellDfn = new CellDfn(text, index: index);
      worksheetDfn.ColumnHeadings.Cells.Add(headerCellDfn);
    }

    private static void CreateCellToRow(
      object? player,
      RowDfn rowDfn,
      CellDefinitionAttribute? cellDefinitionAttribute,
      PropertyInfo playerTypePropertyInfo,
      decimal index)
    {
      if (player == null)
      {
        CreateEmptyCellToRow(rowDfn, cellDefinitionAttribute, index);
      }
      else
      {
        CellDfn cellDfn;
        if (cellDefinitionAttribute != null)
        {
          cellDfn = new CellDfn(playerTypePropertyInfo.GetValue(player) ?? string.Empty, cellDefinitionAttribute.CellDataType, index);
        }
        else
        {
          cellDfn = new CellDfn(playerTypePropertyInfo.GetValue(player) ?? string.Empty, index: index);
        }

        rowDfn.Cells.Add(cellDfn);
      }
    }

    private static void CreateEmptyCellToRow(
      RowDfn rowDfn,
      CellDefinitionAttribute? cellDefinitionAttribute,
      decimal index)
    {
      CellDfn cellDfn;
      if (cellDefinitionAttribute != null)
      {
        cellDfn = new CellDfn(string.Empty, cellDefinitionAttribute.CellDataType, index);
      }
      else
      {
        cellDfn = new CellDfn(string.Empty, index: index);
      }

      rowDfn.Cells.Add(cellDfn);
    }

    private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart, SheetData sheetData)
    {
      var worksheet = new Worksheet();
      var sheetDimension = new SheetDimension { Reference = "A1" };

      var sheetViews = new SheetViews();

      var sheetView = new SheetView { TabSelected = true, WorkbookViewId = 0U };
      var selection = new Selection { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" } };

      sheetView.AppendChild(selection);

      sheetViews.AppendChild(sheetView);
      var sheetFormatProperties = new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 15D };

      var pageMargins = new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
      worksheet.AppendChild(sheetDimension);
      worksheet.AppendChild(sheetViews);
      worksheet.AppendChild(sheetFormatProperties);
      worksheet.AppendChild(sheetData);
      worksheet.AppendChild(pageMargins);
      worksheetPart.Worksheet = worksheet;
    }

    private void AddCellsToRowFromObjectPropertyInfos(
      object? player,
      PropertyInfo[] playerTypePropertyInfos,
      RowDfn rowDfn,
      int iteration,
      int deep,
      decimal parentIndex)
    {
      foreach (var playerTypePropertyInfo in playerTypePropertyInfos)
      {
        CellDefinitionAttribute? cellDefinitionAttribute = GetAttributeFrom<CellDefinitionAttribute>(playerTypePropertyInfo);
        IndexAttribute? indexAttribute = GetAttributeFrom<IndexAttribute>(playerTypePropertyInfo);
        IgnoreFromSpreadSheetAttribute? ignoreFromSpreadSheetAttribute = GetAttributeFrom<IgnoreFromSpreadSheetAttribute>(playerTypePropertyInfo);
        MultiColumnAttribute? multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);

        // TODO - Yanal - relire ce code
        // Index management
        decimal index = parentIndex;
        int power = (int)Math.Pow(10, deep);
        int iterationIncrement = 0;
        if (iteration > 0)
        {
          index += decimal.Divide(iteration, power);
          iterationIncrement = 1;
        }

        if (indexAttribute != null)
        {
          power = (int)Math.Pow(10, deep + iterationIncrement);
          index += decimal.Divide(indexAttribute.Index, power);
        }

        if (multiColumnAttribute != null)
        {
          // Retrieve child object
          if (player != null && playerTypePropertyInfo.GetValue(player) is IEnumerable<object> childPlayersEnumerable)
          {
            object[] childPlayers = childPlayersEnumerable.ToArray();
            int maxNumberOfElement = multiColumnAttribute.MaxNumberOfElement;
            int childDeep = deep + maxNumberOfElement.ToString().Length;
            int currentIteration = 1;
            Type? childPlayerType = null;
            PropertyInfo[]? childPlayerTypePropertyInfos = null;
            foreach (object? childPlayer in childPlayers)
            {
              childPlayerType = childPlayer.GetType();
              childPlayerTypePropertyInfos = childPlayerType.GetProperties();
              AddCellsToRowFromObjectPropertyInfos(childPlayer, childPlayerTypePropertyInfos, rowDfn, currentIteration, childDeep + iterationIncrement, index);
              currentIteration++;
            }

            // Add empty cells if needed
            int numberOfEmptyCellToAdd = maxNumberOfElement - childPlayers.Length;
            if (childPlayerType != null && childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
            {
              for (int i = 0; i < numberOfEmptyCellToAdd; i++)
              {
                AddCellsToRowFromObjectPropertyInfos(null, childPlayerTypePropertyInfos, rowDfn, currentIteration, childDeep + iterationIncrement, index);
                currentIteration++;
              }
            }
          }
          else
          {
            // Add empty cells if needed
            int maxNumberOfElement = multiColumnAttribute.MaxNumberOfElement;
            int childDeep = deep + maxNumberOfElement.ToString().Length;
            int numberOfEmptyCellToAdd = maxNumberOfElement;
            int currentIteration = 1;
            Type? childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.FirstOrDefault();
            if (childPlayerType != null)
            {
              PropertyInfo[] childPlayerTypePropertyInfos = childPlayerType.GetProperties(BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
              for (int i = 0; i < numberOfEmptyCellToAdd; i++)
              {
                AddCellsToRowFromObjectPropertyInfos(null, childPlayerTypePropertyInfos, rowDfn, currentIteration, childDeep + iterationIncrement, index);
                currentIteration++;
              }
            }
            else
            {
              CreateEmptyCellToRow(rowDfn, cellDefinitionAttribute, index);
            }
          }
        }
        else
        {
          if (ignoreFromSpreadSheetAttribute?.IgnoreFlag != true)
          {
            CreateCellToRow(player, rowDfn, cellDefinitionAttribute, playerTypePropertyInfo, index);
          }
        }
      }
    }

    private void AddHeaderCellsToRowFromObjectPropertyInfos(
      WorksheetDfn worksheetDfn,
      object? player,
      Type playerType,
      PropertyInfo[] playerTypePropertyInfos,
      int iteration,
      int deep,
      decimal parentIndex)
    {
      foreach (var playerTypePropertyInfo in playerTypePropertyInfos)
      {
        IndexAttribute? indexAttribute = GetAttributeFrom<IndexAttribute>(playerTypePropertyInfo);
        IgnoreFromSpreadSheetAttribute? ignoreFromSpreadSheetAttribute = GetAttributeFrom<IgnoreFromSpreadSheetAttribute>(playerTypePropertyInfo);
        MultiColumnAttribute? multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);

        // TODO - Yanal - relire ce code
        // Index management
        decimal index = parentIndex;
        int power = (int)Math.Pow(10, deep);
        int iterationIncrement = 0;
        if (iteration > 0)
        {
          index += decimal.Divide(iteration, power);
          iterationIncrement = 1;
        }

        if (indexAttribute != null)
        {
          power = (int)Math.Pow(10, deep + iterationIncrement);
          index += decimal.Divide(indexAttribute.Index, power);
        }

        if (multiColumnAttribute != null)
        {
          // Retrieve child object
          if (player == null)
          {
            // Add empty cells if needed
            int maxNumberOfElement = multiColumnAttribute.MaxNumberOfElement;
            int childDeep = deep + maxNumberOfElement.ToString().Length;
            int numberOfEmptyCellToAdd = maxNumberOfElement;
            int currentIteration = 1;
            Type? childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.FirstOrDefault();
            if (childPlayerType != null)
            {
              PropertyInfo[] childPlayerTypePropertyInfos = childPlayerType.GetProperties(BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
              for (int i = 0; i < numberOfEmptyCellToAdd; i++)
              {
                AddHeaderCellsToRowFromObjectPropertyInfos(worksheetDfn, null, childPlayerType, childPlayerTypePropertyInfos, currentIteration, childDeep + iterationIncrement, index);
                currentIteration++;
              }
            }
            else
            {
              AddHeaderCellToWorkSheet(worksheetDfn, string.Empty, index);
            }
          }
          else if (playerTypePropertyInfo.GetValue(player) is IEnumerable<object> childPlayersEnumerable)
          {
            object[] childPlayers = childPlayersEnumerable.ToArray();
            int maxNumberOfElement = multiColumnAttribute.MaxNumberOfElement;
            int childDeep = deep + maxNumberOfElement.ToString().Length;
            int currentIteration = 1;
            Type? childPlayerType = null;
            PropertyInfo[]? childPlayerTypePropertyInfos = null;
            foreach (object? childPlayer in childPlayers)
            {
              childPlayerType = childPlayer.GetType();
              childPlayerTypePropertyInfos = childPlayerType.GetProperties();
              AddHeaderCellsToRowFromObjectPropertyInfos(worksheetDfn, childPlayer, childPlayerType, childPlayerTypePropertyInfos, currentIteration, childDeep + iterationIncrement, index);
              currentIteration++;
            }

            // Add empty cells if needed
            int numberOfEmptyCellToAdd = maxNumberOfElement - childPlayers.Length;
            if (childPlayerType != null && childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
            {
              for (int i = 0; i < numberOfEmptyCellToAdd; i++)
              {
                AddHeaderCellsToRowFromObjectPropertyInfos(worksheetDfn, null, childPlayerType, childPlayerTypePropertyInfos, currentIteration, childDeep + iterationIncrement, index);
                currentIteration++;
              }
            }
          }
          else if (playerTypePropertyInfo.GetValue(player) != null)
          {
            AddHeaderCellToWorkSheet(worksheetDfn, string.Empty, index);
          }
        }
        else
        {
          if (ignoreFromSpreadSheetAttribute?.IgnoreFlag != true)
          {
            string key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}_{index}";
            if (_headers.Add(key))
            {
              HeaderAttribute? headerAttribute = GetAttributeFrom<HeaderAttribute>(playerTypePropertyInfo);
              if (headerAttribute != null)
              {
                // TODO - et si le header n'est pas défini ?
                string text = headerAttribute.Text;
                if (headerAttribute.TextToAddToHeader != null)
                {
                  PropertyInfo? textToAddToHeaderPropertyInfo = playerType.GetProperty(headerAttribute.TextToAddToHeader);
                  if (textToAddToHeaderPropertyInfo != null && textToAddToHeaderPropertyInfo.GetValue(player, null) != null)
                  {
                    text = string.Format(text, textToAddToHeaderPropertyInfo.GetValue(player, null));
                  }
                }

                AddHeaderCellToWorkSheet(worksheetDfn, text, index);
              }
              else
              {
                AddHeaderCellToWorkSheet(worksheetDfn, string.Empty, index);
              }
            }
          }
        }
      }
    }

    private WorkbookDfn BuildWorkbook(object team)
    {
      var workbookDfn = new WorkbookDfn();
      var teamType = team.GetType();
      var teamTypePropertyInfos = teamType.GetProperties();
      var i = 1;

      foreach (var teamTypePropertyInfo in teamTypePropertyInfos)
      {
        var emptyExportMessage = MessageRes.EmptyMessageDefault;
        var emptyExportMessageAttribute = GetAttributeFrom<EmptyResultMessageAttribute>(teamTypePropertyInfo);
        if (emptyExportMessageAttribute != null && !string.IsNullOrEmpty(emptyExportMessageAttribute.Text))
        {
          emptyExportMessage = emptyExportMessageAttribute.Text;
        }

        var sheetNameAttribute = GetAttributeFrom<SheetNameAttribute>(teamTypePropertyInfo);
        var sheetName = $"Sheet{i}";
        if (sheetNameAttribute != null)
        {
          sheetName = sheetNameAttribute.Text;
        }

        var worksheetDfn = new WorksheetDfn(sheetName);

        if (teamTypePropertyInfo.GetValue(team) is IEnumerable<object?> playersEnumerable)
        {
          object?[] players = playersEnumerable as object?[] ?? playersEnumerable.ToArray();

          // Add data (header lines + data lines)
          if (!players.Any())
          {
            // Create fake cell with warning message
            RowDfn rowDfn = new RowDfn();
            worksheetDfn.Rows.Add(rowDfn);
            CellDfn cellDfn = new CellDfn(emptyExportMessage, 0);
            rowDfn.Cells.Add(cellDfn);
            workbookDfn.Worksheets.Add(worksheetDfn);
          }
          else
          {
            foreach (object? player in players)
            {
              if (player != null)
              {
                var playerType = player.GetType();
                PropertyInfo[] playerTypePropertyInfos = playerType.GetProperties();
                CalculateMaxNumberOfElement(player, playerTypePropertyInfos);
              }
            }

            // Headers
            foreach (object? player in players)
            {
              if (player != null)
              {
                var playerType = player.GetType();
                PropertyInfo[] playerTypePropertyInfos = playerType.GetProperties();
                AddHeaderCellsToRowFromObjectPropertyInfos(worksheetDfn, player, playerType, playerTypePropertyInfos, 0, 0, 0);
              }
            }

            // Rows
            foreach (object? player in players)
            {
              if (player != null)
              {
                var playerType = player.GetType();
                PropertyInfo[] playerTypePropertyInfos = playerType.GetProperties();
                RowDfn rowDfn = new RowDfn();
                worksheetDfn.Rows.Add(rowDfn);
                AddCellsToRowFromObjectPropertyInfos(player, playerTypePropertyInfos, rowDfn, 0, 0, 0);
              }
            }
          }
        }

        workbookDfn.Worksheets.Add(worksheetDfn);
        i++;
      }

      foreach (WorksheetDfn worksheet in workbookDfn.Worksheets)
      {
        worksheet.ColumnHeadings.OrderCells();

        foreach (RowDfn rowDfn in worksheet.Rows)
        {
          rowDfn.OrderCells();
        }
      }

      return workbookDfn;
    }

    private void CalculateMaxNumberOfElement(object? player, PropertyInfo[] playerTypePropertyInfos)
    {
      foreach (var playerTypePropertyInfo in playerTypePropertyInfos)
      {
        MultiColumnAttribute? multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);
        if (multiColumnAttribute != null)
        {
          if (playerTypePropertyInfo.GetValue(player) is IEnumerable<object> childPlayersEnumerable)
          {
            object[] childPlayers = childPlayersEnumerable.ToArray();
            int numberOfElement = childPlayers.Length;
            if (numberOfElement > multiColumnAttribute.MaxNumberOfElement)
            {
              multiColumnAttribute.MaxNumberOfElement = numberOfElement;
            }
          }
        }
      }
    }

    private Cell CreateCell(CellDfn cellDfn)
    {
      var stylIndex = CreateOrGetStylIndex(cellDfn);

      var cell = new Cell
      {
        StyleIndex = stylIndex,
      };

      if (cellDfn.Value == null)
      {
        cell.CellValue = new CellValue(string.Empty);
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
      }
      else
      {
        if (cellDfn.Value is DateTime dateTimeValue)
        {
          cell.CellValue = new CellValue(dateTimeValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else if (cellDfn.Value is DateTimeOffset dateTimeOffsetValue)
        {
          cell.CellValue = new CellValue(dateTimeOffsetValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else if (cellDfn.Value is bool boolValue)
        {
          int intValue = boolValue ? 0 : 1;
          cell.CellValue = new CellValue(intValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (cellDfn.Value is byte byteValue)
        {
          cell.CellValue = new CellValue(byteValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellDfn.Value is decimal decimalValue)
        {
          cell.CellValue = new CellValue(decimalValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellDfn.Value is double doubleValue)
        {
          cell.CellValue = new CellValue(doubleValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellDfn.Value is int intValue)
        {
          cell.CellValue = new CellValue(intValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellDfn.Value is string stringValue)
        {
          stringValue = XmlStringHelper.Sanitize(stringValue);
          cell.CellValue = new CellValue(stringValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }
        else if (cellDfn.Value is TimeSpan timeSpanValue)
        {
          // Excel saves time in seconds divided by maximum seconds of a day
          double cellValue = timeSpanValue.TotalSeconds / 86400; // 86400 = 24 * 60 *60
          cell.CellValue = new CellValue(cellValue.ToString(CultureInfo.InvariantCulture));
        }
        else
        {
          throw new NotSupportedException($"Type {cellDfn.Value.GetType()} is not supported as a Cell value");
        }
      }

      return cell;
    }

    private Row CreateHeaderRowForExcel(IEnumerable<CellDfn> columnHeadings)
    {
      var row = new Row();
      foreach (var cellDfn in columnHeadings)
      {
        row.AppendChild(CreateCell(cellDfn));
      }

      return row;
    }

    private uint CreateOrGetStylIndex(CellDfn cellDfn)
    {
      int styleHashCode = cellDfn.GetStyleHashCode();
      if (Table.TryGetValue(styleHashCode, out var stylIndex))
      {
        return stylIndex;
      }

      var cellFormat = new CellFormat
      {
        ApplyBorder = true,
        ApplyFont = true,
        ApplyNumberFormat = BooleanValue.FromBoolean(true),
        BorderId = 0U,
        FillId = 0U,
        FormatId = 0U,
        FontId = 0U,
      };

      // https://stackoverflow.com/questions/11781210/c-sharp-open-xml-2-0-numberformatid-range
      if (cellDfn.CellDataType == CellDataType.Date)
      {
        cellFormat.NumberFormatId = 14U; // d/m/yyyy
      }
      else if (cellDfn.CellDataType == CellDataType.String)
      {
        cellFormat.NumberFormatId = 49U; // @
      }
      else if (cellDfn.CellDataType == CellDataType.Percentage)
      {
        cellFormat.NumberFormatId = 10U;
      }
      else if (cellDfn.CellDataType == CellDataType.Time)
      {
        cellFormat.NumberFormatId = 20U; // H:mm
      }
      else
      {
        cellFormat.NumberFormatId = 0U;
      }

      var index = _stylesheet.CellFormats!.Count!.Value;
      _stylesheet.CellFormats!.Count!.Value++;
      _stylesheet.CellFormats.AppendChild(cellFormat);
      Table.Add(styleHashCode, index);

      return index;
    }

    private void CreatePartsForExcel(SpreadsheetDocument document)
    {
      var workbookPart = document.AddWorkbookPart();
      var workbook = new Workbook();
      workbookPart.Workbook = workbook;
      var sheets = new Sheets();
      workbook.AppendChild(sheets);

      var workbookStylesPart1 = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
      GenerateWorkbookStylesPartContent(workbookStylesPart1);

      // Thank you https://stackoverflow.com/questions/9120544/openxml-multiple-sheets
      uint count = 1U;
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheet = new Sheet { Name = worksheet.Name, SheetId = count, Id = workbookPart.GetIdOfPart(worksheetPart) };
        sheets.AppendChild(sheet);
        var sheetData = GenerateSheetDataForDetails(worksheet);
        GenerateWorksheetPartContent(worksheetPart, sheetData);
        count++;
      }
    }

    private Row GenerateRowForChildPartDetail(RowDfn rowDfn)
    {
      var row = new Row();

      foreach (var cellDfn in rowDfn.Cells)
      {
        row.AppendChild(CreateCell(cellDfn));
      }

      return row;
    }

    private SheetData GenerateSheetDataForDetails(WorksheetDfn worksheet)
    {
      var sheetData1 = new SheetData();
      if (worksheet.ColumnHeadings.Cells.Count > 0)
      {
        sheetData1.AppendChild(CreateHeaderRowForExcel(worksheet.ColumnHeadings.Cells));
      }

      foreach (var row in worksheet.Rows)
      {
        var partsRows = GenerateRowForChildPartDetail(row);
        sheetData1.AppendChild(partsRows);
      }

      return sheetData1;
    }

    private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
    {
      var fonts = new Fonts { Count = 1U };

      // Font 1
      var font = new Font
      {
        FontSize = new FontSize { Val = 11D },
        FontName = new FontName { Val = "Calibri" },
        FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
        FontScheme = new FontScheme { Val = FontSchemeValues.Minor },
      };

      fonts.AppendChild(font);

      // Default Fill
      var fills = new Fills { Count = 1U };
      var fill = new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } };
      fills.AppendChild(fill);

      // Default Border
      var borders = new Borders { Count = 1U };
      var border = new Border
      {
        LeftBorder = new LeftBorder(),
        RightBorder = new RightBorder(),
        TopBorder = new TopBorder(),
        BottomBorder = new BottomBorder(),
        DiagonalBorder = new DiagonalBorder(),
      };
      borders.AppendChild(border);

      // CellStyleFormats
      var cellStyleFormats = new CellStyleFormats { Count = 1U };
      var cellFormat = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
      cellStyleFormats.AppendChild(cellFormat);

      // CellFormats
      var cellFormats = new CellFormats { Count = 0U };

      _stylesheet.AppendChild(fonts);
      _stylesheet.AppendChild(fills);
      _stylesheet.AppendChild(borders);
      _stylesheet.AppendChild(cellStyleFormats);
      _stylesheet.AppendChild(cellFormats);

      workbookStylesPart.Stylesheet = _stylesheet;
    }

    private T? GetAttributeFrom<T>(PropertyInfo propertyInfo)
      where T : Attribute
    {
      string key = $"{propertyInfo.Module.MetadataToken}_{propertyInfo.MetadataToken}_{typeof(T).Name}"; // TODO Yanal - voir si on peut retirer _{typeof(T).Name}
      if (_cachedAttributes.TryGetValue(key, out var cachedAttribute))
      {
        return (T?)cachedAttribute;
      }

      var attrType = typeof(T);

      // property is expected to be not null because instance and property
      var attribute = (T?)propertyInfo.GetCustomAttributes(attrType, false).FirstOrDefault();
      _cachedAttributes.Add(key, attribute);
      return attribute;
    }

    private void OrderWorkBookDfn()
    {
      foreach (WorksheetDfn worksheet in _workbookDfn.Worksheets)
      {
        worksheet.ColumnHeadings.OrderCells();

        foreach (RowDfn rowDfn in worksheet.Rows)
        {
          rowDfn.OrderCells();
        }
      }
    }

    private void Validate()
    {
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        var count = _workbookDfn.Worksheets.Count(w => w.Name == worksheet.Name);
        if (count > 1)
        {
          throw new DefinitionException($"Only one worksheet could be named [{worksheet.Name}]");
        }

        if (worksheet.Name.Length > 31)
        {
          throw new SimpleExcelExporterException(string.Format(MessageRes.SheetNameLengthTooLong, worksheet.Name));
        }
      }

      if (!_workbookDfn.Worksheets.Any())
      {
        throw new DefinitionException("WorkBook could not be null or empty.");
      }
    }
  }
}
