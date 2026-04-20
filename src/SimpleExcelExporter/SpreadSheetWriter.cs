namespace SimpleExcelExporter
{
  using System;
  using System.Collections.Generic;
  using System.Globalization;
  using System.IO;
  using System.IO.Compression;
  using System.Linq;
  using System.Reflection;
  using System.Text;
  using System.Xml;
  using DocumentFormat.OpenXml;
  using DocumentFormat.OpenXml.Spreadsheet;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;
  using SimpleExcelExporter.Resources;

  public class SpreadsheetWriter
  {
    private static readonly Dictionary<CellValues, string> CellTypeAttributes = new()
    {
      [CellValues.Boolean] = "b",
      [CellValues.Date] = "d",
      [CellValues.Error] = "e",
      [CellValues.InlineString] = "inlineStr",
      [CellValues.Number] = "n",
      [CellValues.SharedString] = "s",
      [CellValues.String] = "str",
    };

    private readonly Dictionary<string, Attribute?> _cachedAttributes = [];

    private readonly Dictionary<string, (CellDfn, bool)> _headers = [];

    private readonly Dictionary<string, int> _multiColumnAttribute = [];

    private readonly Stream _stream;

    private readonly List<uint> _numberFormatIds = [];

    private readonly List<string> _sharedStrings = [];

    private readonly Dictionary<string, int> _sharedStringsIndex = [];

    private readonly WorkbookDfn _workbookDfn;

    private int _sharedStringsTotalCount;

    /// <summary>
    /// Initializes a new <see cref="SpreadsheetWriter"/> from an explicit <see cref="WorkbookDfn"/>.
    /// </summary>
    /// <param name="stream">The destination stream for the XLSX output. The caller keeps ownership and is responsible for disposal.</param>
    /// <param name="workbookDfn">The fully built workbook definition.</param>
    /// <remarks>
    /// The constructor orders the cells in every row and validates the workbook; call <see cref="Write"/>
    /// afterwards to produce the actual .xlsx file.
    /// </remarks>
    public SpreadsheetWriter(Stream stream, WorkbookDfn workbookDfn)
    {
      _stream = stream;
      _workbookDfn = workbookDfn;
      OrderWorkBookDfn();
      Validate();
    }

    /// <summary>
    /// Initializes a new <see cref="SpreadsheetWriter"/> from an annotated object graph.
    /// </summary>
    /// <param name="stream">The destination stream for the XLSX output. The caller keeps ownership and is responsible for disposal.</param>
    /// <param name="team">An object whose properties are decorated with SimpleExcelExporter attributes (<c>SheetName</c>, <c>Header</c>, <c>CellDefinition</c>, …).</param>
    /// <remarks>
    /// This constructor runs a reflection walk over <paramref name="team"/> to build the underlying
    /// <see cref="WorkbookDfn"/>, then orders and validates it. For large inputs this phase usually
    /// dominates the total runtime — more than the subsequent <see cref="Write"/> call.
    /// </remarks>
    public SpreadsheetWriter(Stream stream, object team)
    {
      _stream = stream;
      _workbookDfn = BuildWorkbook(team);
      OrderWorkBookDfn();
      Validate();
    }

    private Dictionary<int, uint> Table { get; } = [];

    public void Write()
    {
      // Adding core file properties is mandatory to avoid a problem with Google Spreadsheet transforming from XLSX to XLSM.
      // cf. https://stackoverflow.com/questions/70319867/avoid-google-spreadsheet-to-convert-an-xlsx-file-created-by-open-xml-sdk-to-xlsm
      // cf. https://github.com/OfficeDev/Open-XML-SDK/issues/1093
      // cf. https://issuetracker.google.com/issues/210875597
      using var archive = new ZipArchive(_stream, ZipArchiveMode.Create, leaveOpen: true);

      // WriteStyles and WriteSharedStrings must run AFTER WriteWorksheets so that
      // _numberFormatIds and _sharedStrings have been populated during cell creation.
      WriteContentTypes(archive);
      WritePackageRelationships(archive);
      WriteCoreProperties(archive);
      WriteWorkbookRelationships(archive);
      WriteWorkbook(archive);
      WriteWorksheets(archive);
      WriteStyles(archive);
      WriteSharedStrings(archive);
    }

    private static CellDfn AddHeaderCellToWorkSheet(WorksheetDfn worksheetDfn, string text, List<int> index)
    {
      var headerCellDfn = new CellDfn(text, index: index);
      worksheetDfn.ColumnHeadings.Cells.Add(headerCellDfn);
      return headerCellDfn;
    }

    private static void CreateCellToRow(
      object? player,
      RowDfn rowDfn,
      CellDefinitionAttribute? cellDefinitionAttribute,
      PropertyInfo playerTypePropertyInfo,
      List<int> index)
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
          cellDfn = new CellDfn(playerTypePropertyInfo.GetValue(player) ?? string.Empty, index, cellDefinitionAttribute.CellDataType);
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
      List<int> index)
    {
      CellDfn cellDfn;
      if (cellDefinitionAttribute != null)
      {
        cellDfn = new CellDfn(string.Empty, index, cellDefinitionAttribute.CellDataType);
      }
      else
      {
        cellDfn = new CellDfn(string.Empty, index: index);
      }

      rowDfn.Cells.Add(cellDfn);
    }

    private static List<int> ManageIndex(int iteration, List<int>? parentIndex, IndexAttribute? indexAttribute)
    {
      // Index management
      var index = parentIndex ?? [];
      if (iteration > 0)
      {
        index.Add(iteration);
      }

      if (indexAttribute != null)
      {
        index.Add(indexAttribute.Index);
      }

      return index;
    }

    private static string ToCellTypeAttribute(CellValues value) =>
      CellTypeAttributes.TryGetValue(value, out var attribute) ? attribute : value.ToString();

    private static void WriteCoreProperties(ZipArchive archive)
    {
      var entry = archive.CreateEntry("docProps/core.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      using var writer = new XmlTextWriter(stream, Encoding.UTF8);

      var nowIso = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
      writer.WriteRaw(
        $"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n" +
        "<cp:coreProperties " +
        "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
        "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<dc:creator>SimpleExcelExporter</dc:creator>" +
        $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{nowIso}</dcterms:created>" +
        $"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{nowIso}</dcterms:modified>" +
        "</cp:coreProperties>");
      writer.Flush();
    }

    private static void WriteOverride(XmlWriter writer, string partName, string contentType)
    {
      writer.WriteStartElement("Override", Ns.ContentTypes);
      writer.WriteAttributeString("PartName", partName);
      writer.WriteAttributeString("ContentType", contentType);
      writer.WriteEndElement();
    }

    private static void WritePackageRelationships(ZipArchive archive)
    {
      var entry = archive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement("Relationships", Ns.PackageRelationships);

      WriteRelationship(
        writer,
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "xl/workbook.xml");

      WriteRelationship(
        writer,
        "rId2",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml");

      writer.WriteEndElement();
      writer.WriteEndDocument();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
      writer.WriteStartElement("Relationship", Ns.PackageRelationships);
      writer.WriteAttributeString("Id", id);
      writer.WriteAttributeString("Type", type);
      writer.WriteAttributeString("Target", target);
      writer.WriteEndElement();
    }

    private void AddCellsToRowFromObjectPropertyInfos(
      object? player,
      PropertyInfo[] playerTypePropertyInfos,
      RowDfn rowDfn,
      int iteration,
      List<int>? parentIndex)
    {
      var objectQueue = new Queue<(object?, PropertyInfo[], int, List<int>?)>(); // Use a queue to manage child objects
      objectQueue.Enqueue((player, playerTypePropertyInfos, iteration, parentIndex));

      while (objectQueue.Count > 0)
      {
        var (currentPlayer, currentPlayerTypePropertyInfos, currentIteration, currentParentIndex) = objectQueue.Dequeue();

        foreach (var playerTypePropertyInfo in currentPlayerTypePropertyInfos)
        {
          var cellDefinitionAttribute = GetAttributeFrom<CellDefinitionAttribute>(playerTypePropertyInfo);
          var indexAttribute = GetAttributeFrom<IndexAttribute>(playerTypePropertyInfo);
          var ignoreFromSpreadSheetAttribute = GetAttributeFrom<IgnoreFromSpreadSheetAttribute>(playerTypePropertyInfo);
          var multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);
          var index = ManageIndex(currentIteration, currentParentIndex, indexAttribute);
          if (multiColumnAttribute != null)
          {
            if (currentPlayer != null && playerTypePropertyInfo.GetValue(currentPlayer) is IEnumerable<object?> childPlayers)
            {
              var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}";
              var maxNumberOfElement = _multiColumnAttribute[key];
              var childIteration = 1;
              var childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.Single();
              PropertyInfo[]? childPlayerTypePropertyInfos = null;
              foreach (var childPlayer in childPlayers)
              {
                childPlayerTypePropertyInfos = childPlayerType.GetProperties();
                objectQueue.Enqueue((childPlayer, childPlayerTypePropertyInfos, childIteration, index)); // Enqueue child object for later processing
                childIteration++;
              }

              // Add empty cells if needed
              var numberOfEmptyCellToAdd = maxNumberOfElement - childIteration + 1;
              if (childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
              {
                for (var i = 0; i < numberOfEmptyCellToAdd; i++)
                {
                  objectQueue.Enqueue((null, childPlayerTypePropertyInfos, childIteration, index)); // Enqueue child object for later processing
                  childIteration++;
                }
              }
            }
            else
            {
              // Add empty cells if needed
              var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}";
              var numberOfEmptyCellToAdd = _multiColumnAttribute[key];
              var childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.FirstOrDefault();
              if (childPlayerType != null)
              {
                var childPlayerTypePropertyInfos = childPlayerType.GetProperties(BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
                for (var i = 0; i < numberOfEmptyCellToAdd; i++)
                {
                  objectQueue.Enqueue((null, childPlayerTypePropertyInfos, i + 1, index)); // Enqueue child object for later processing
                }
              }
              else
              {
                CreateEmptyCellToRow(rowDfn, cellDefinitionAttribute, index);
              }
            }
          }
          else if (ignoreFromSpreadSheetAttribute?.IgnoreFlag != true)
          {
            CreateCellToRow(currentPlayer, rowDfn, cellDefinitionAttribute, playerTypePropertyInfo, index);
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
      List<int>? parentIndex)
    {
      var objectQueue = new Queue<(object?, Type, PropertyInfo[], int, List<int>?)>(); // Use a queue to manage child objects
      objectQueue.Enqueue((player, playerType, playerTypePropertyInfos, iteration, parentIndex));

      while (objectQueue.Count > 0)
      {
        var (currentPlayer, currentPlayerType, currentPlayerTypePropertyInfos, currentIteration, currentParentIndex) = objectQueue.Dequeue();

        foreach (var playerTypePropertyInfo in currentPlayerTypePropertyInfos)
        {
          var indexAttribute = GetAttributeFrom<IndexAttribute>(playerTypePropertyInfo);
          var ignoreFromSpreadSheetAttribute = GetAttributeFrom<IgnoreFromSpreadSheetAttribute>(playerTypePropertyInfo);
          var multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);
          var index = ManageIndex(currentIteration, currentParentIndex, indexAttribute);
          if (multiColumnAttribute != null)
          {
            if (currentPlayer != null && playerTypePropertyInfo.GetValue(currentPlayer) is IEnumerable<object?> childPlayers)
            {
              var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}";
              var maxNumberOfElement = _multiColumnAttribute[key];
              var childIteration = 1;
              var childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.Single();
              PropertyInfo[]? childPlayerTypePropertyInfos = null;
              foreach (var childPlayer in childPlayers)
              {
                childPlayerTypePropertyInfos = childPlayerType.GetProperties();
                objectQueue.Enqueue((childPlayer, childPlayerType, childPlayerTypePropertyInfos, childIteration, index)); // Enqueue child object for later processing
                childIteration++;
              }

              // Add empty cells if needed
              var numberOfEmptyCellToAdd = maxNumberOfElement - childIteration + 1;
              if (childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
              {
                for (var i = 0; i < numberOfEmptyCellToAdd; i++)
                {
                  objectQueue.Enqueue((null, childPlayerType, childPlayerTypePropertyInfos, childIteration, index)); // Enqueue child object for later processing
                  childIteration++;
                }
              }
            }
            else
            {
              // Add empty cells if needed
              var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}";
              var numberOfEmptyCellToAdd = _multiColumnAttribute[key];
              var childPlayerType = playerTypePropertyInfo.PropertyType.GenericTypeArguments.FirstOrDefault();
              if (childPlayerType != null)
              {
                var childPlayerTypePropertyInfos = childPlayerType.GetProperties(BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
                for (var i = 0; i < numberOfEmptyCellToAdd; i++)
                {
                  objectQueue.Enqueue((null, childPlayerType, childPlayerTypePropertyInfos, i + 1, index)); // Enqueue child object for later processing
                }
              }
              else
              {
                _ = AddHeaderCellToWorkSheet(worksheetDfn, string.Empty, index);
              }
            }
          }
          else if (ignoreFromSpreadSheetAttribute?.IgnoreFlag != true)
          {
            var text = BuildText(playerTypePropertyInfo, currentPlayerType, currentPlayer);

            var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}_{string.Join("_", index)}";
            if (!_headers.TryGetValue(key, out var value))
            {
              var headerCellDfn = AddHeaderCellToWorkSheet(worksheetDfn, string.IsNullOrEmpty(text) ? string.Empty : text, index);
              _headers.Add(key, (headerCellDfn, currentPlayer != null));
            }
            else
            {
              // If the currentPlayer was null when the existing headerCell was added in _headers, then we should update the text.
              var (headerCellDfn, textCorrectlySetFlag) = value;
              if (!textCorrectlySetFlag && currentPlayer != null)
              {
                _ = _headers.Remove(key);
                headerCellDfn.Value = text;
                _headers.Add(key, (headerCellDfn, true));
              }
            }
          }
        }
      }
    }

    private string? BuildText(PropertyInfo playerTypePropertyInfo, Type currentPlayerType, object? currentPlayer)
    {
      var headerAttribute = GetAttributeFrom<HeaderAttribute>(playerTypePropertyInfo);
      string? text = null;
      if (headerAttribute != null)
      {
        text = headerAttribute.Text;
        if (headerAttribute.TextToAddToHeader != null)
        {
          var textToAddToHeaderPropertyInfo = currentPlayerType.GetProperty(headerAttribute.TextToAddToHeader);
          if (currentPlayer != null)
          {
            if (textToAddToHeaderPropertyInfo?.GetValue(currentPlayer, null) != null)
            {
              text = string.Format(text, textToAddToHeaderPropertyInfo.GetValue(currentPlayer, null));
            }
          }
        }
      }

      return text;
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
          var players = playersEnumerable as object?[] ?? playersEnumerable.ToArray();

          // Add data (header lines + data lines)
          if (players.Length == 0)
          {
            // Create fake cell with warning message
            var rowDfn = new RowDfn();
            worksheetDfn.Rows.Add(rowDfn);
            var cellDfn = new CellDfn(emptyExportMessage);
            rowDfn.Cells.Add(cellDfn);
            workbookDfn.Worksheets.Add(worksheetDfn);
          }
          else
          {
            foreach (var player in players)
            {
              if (player != null)
              {
                CalculateMaxNumberOfElement(player);
              }
            }

            // Headers
            foreach (var player in players)
            {
              if (player != null)
              {
                var playerType = player.GetType();
                var playerTypePropertyInfos = playerType.GetProperties();
                AddHeaderCellsToRowFromObjectPropertyInfos(worksheetDfn, player, playerType, playerTypePropertyInfos, 0, null);
              }
            }

            // Rows
            foreach (var player in players)
            {
              if (player != null)
              {
                var playerType = player.GetType();
                var playerTypePropertyInfos = playerType.GetProperties();
                var rowDfn = new RowDfn();
                worksheetDfn.Rows.Add(rowDfn);
                AddCellsToRowFromObjectPropertyInfos(player, playerTypePropertyInfos, rowDfn, 0, null);
              }
            }
          }
        }

        workbookDfn.Worksheets.Add(worksheetDfn);
        i++;
      }

      foreach (var worksheet in workbookDfn.Worksheets)
      {
        worksheet.ColumnHeadings.OrderCells();

        foreach (var rowDfn in worksheet.Rows)
        {
          rowDfn.OrderCells();
        }
      }

      return workbookDfn;
    }

    private void CalculateMaxNumberOfElement(object player)
    {
      var objectQueue = new Queue<object>(); // Use a queue to manage child objects
      objectQueue.Enqueue(player);

      while (objectQueue.Count > 0)
      {
        var currentPlayer = objectQueue.Dequeue();
        var playerType = currentPlayer.GetType();
        var playerTypePropertyInfos = playerType.GetProperties();
        foreach (var playerTypePropertyInfo in playerTypePropertyInfos)
        {
          var multiColumnAttribute = GetAttributeFrom<MultiColumnAttribute>(playerTypePropertyInfo);
          if (multiColumnAttribute != null)
          {
            var key = $"{playerTypePropertyInfo.Module.MetadataToken}_{playerTypePropertyInfo.MetadataToken}";
            if (!_multiColumnAttribute.TryGetValue(key, out var value))
            {
              value = multiColumnAttribute.MaxNumberOfElement;
              _multiColumnAttribute.Add(key, value);
            }

            if (playerTypePropertyInfo.GetValue(currentPlayer) is IEnumerable<object?> childPlayers)
            {
              var numberOfElement = childPlayers.Count(x => x != null);
              if (numberOfElement > value)
              {
                _multiColumnAttribute[key] = numberOfElement;
              }

              objectQueue.Enqueue(childPlayers);
            }
          }
        }
      }
    }

    private Cell CreateCell(CellDfn cellDfn, uint rowIndex, int columnIndex)
    {
      var stylIndex = CreateOrGetStylIndex(cellDfn);

      var cell = new Cell
      {
        CellReference = $"{ColumnReferenceHelper.ToLetters(columnIndex)}{rowIndex}",
        StyleIndex = stylIndex,
      };

      if (cellDfn.Value == null)
      {
        cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
      }
      else if (cellDfn.Value is DateTime dateTimeValue)
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
        var intValue = boolValue ? 0 : 1;
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
        if (string.IsNullOrEmpty(stringValue))
        {
          // Empty strings stay as self-closing inlineStr — matches Apple Numbers' preferred shape
          // and avoids polluting the shared-strings table with empty entries.
          cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
        }
        else
        {
          stringValue = XmlStringHelper.Sanitize(stringValue);
          var index = GetOrAddSharedString(stringValue);
          cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
          cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
        }
      }
      else if (cellDfn.Value is TimeSpan timeSpanValue)
      {
        // Excel saves time in seconds divided by maximum seconds of a day
        var cellValue = timeSpanValue.TotalSeconds / 86400; // 86400 = 24 * 60 *60
        cell.CellValue = new CellValue(cellValue.ToString(CultureInfo.InvariantCulture));
      }
      else
      {
        throw new NotSupportedException($"Type {cellDfn.Value.GetType()} is not supported as a Cell value");
      }

      return cell;
    }

    private Row CreateHeaderRowForExcel(IEnumerable<CellDfn> columnHeadings, uint rowIndex)
    {
      var row = new Row { RowIndex = rowIndex };
      var columnIndex = 1;
      foreach (var cellDfn in columnHeadings)
      {
        _ = row.AppendChild(CreateCell(cellDfn, rowIndex, columnIndex));
        columnIndex++;
      }

      return row;
    }

    private uint CreateOrGetStylIndex(CellDfn cellDfn)
    {
      var styleHashCode = cellDfn.GetStyleHashCode();
      if (Table.TryGetValue(styleHashCode, out var stylIndex))
      {
        return stylIndex;
      }

      // https://stackoverflow.com/questions/11781210/c-sharp-open-xml-2-0-numberformatid-range
      var numberFormatId = cellDfn.CellDataType switch
      {
        CellDataType.Date => 14U, // d/m/yyyy
        CellDataType.String => 49U, // @
        CellDataType.Percentage => 10U,
        CellDataType.Time => 20U, // H:mm
        _ => 0U,
      };

      var index = (uint)_numberFormatIds.Count;
      _numberFormatIds.Add(numberFormatId);
      Table.Add(styleHashCode, index);

      return index;
    }

    private Row GenerateRowForChildPartDetail(RowDfn rowDfn, uint rowIndex)
    {
      var row = new Row { RowIndex = rowIndex };
      var columnIndex = 1;
      foreach (var cellDfn in rowDfn.Cells)
      {
        _ = row.AppendChild(CreateCell(cellDfn, rowIndex, columnIndex));
        columnIndex++;
      }

      return row;
    }

    private (SheetData SheetData, int MaxColumnCount, uint LastRowIndex) GenerateSheetDataForDetails(WorksheetDfn worksheet)
    {
      var sheetData1 = new SheetData();
      var currentRowIndex = 1U;
      var maxColumnCount = 0;

      if (worksheet.ColumnHeadings.Cells.Count > 0)
      {
        _ = sheetData1.AppendChild(CreateHeaderRowForExcel(worksheet.ColumnHeadings.Cells, currentRowIndex));
        maxColumnCount = worksheet.ColumnHeadings.Cells.Count;
        currentRowIndex++;
      }

      foreach (var row in worksheet.Rows)
      {
        var partsRows = GenerateRowForChildPartDetail(row, currentRowIndex);
        _ = sheetData1.AppendChild(partsRows);
        if (row.Cells.Count > maxColumnCount)
        {
          maxColumnCount = row.Cells.Count;
        }

        currentRowIndex++;
      }

      var lastRowIndex = currentRowIndex - 1U;
      return (sheetData1, maxColumnCount, lastRowIndex);
    }

    private T? GetAttributeFrom<T>(PropertyInfo propertyInfo)
      where T : Attribute
    {
      var key = $"{propertyInfo.Module.MetadataToken}_{propertyInfo.MetadataToken}_{typeof(T).Name}";
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
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        worksheet.ColumnHeadings.OrderCells();

        foreach (var rowDfn in worksheet.Rows)
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

      if (_workbookDfn.Worksheets.Count == 0)
      {
        throw new DefinitionException("WorkBook could not be null or empty.");
      }
    }

    private void WriteContentTypes(ZipArchive archive)
    {
      var entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement("Types", Ns.ContentTypes);

      writer.WriteStartElement("Default", Ns.ContentTypes);
      writer.WriteAttributeString("Extension", "xml");
      writer.WriteAttributeString("ContentType", "application/xml");
      writer.WriteEndElement();

      writer.WriteStartElement("Default", Ns.ContentTypes);
      writer.WriteAttributeString("Extension", "rels");
      writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
      writer.WriteEndElement();

      WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
      WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
      WriteOverride(writer, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
      WriteOverride(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");

      for (var i = 1; i <= _workbookDfn.Worksheets.Count; i++)
      {
        WriteOverride(writer, $"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
      }

      writer.WriteEndElement();
      writer.WriteEndDocument();
    }

    private int GetOrAddSharedString(string value)
    {
      _sharedStringsTotalCount++;
      if (_sharedStringsIndex.TryGetValue(value, out var index))
      {
        return index;
      }

      index = _sharedStrings.Count;
      _sharedStrings.Add(value);
      _sharedStringsIndex[value] = index;
      return index;
    }

    private void WriteSharedStrings(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement(string.Empty, "sst", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", _sharedStringsTotalCount.ToString(CultureInfo.InvariantCulture));
      writer.WriteAttributeString("uniqueCount", _sharedStrings.Count.ToString(CultureInfo.InvariantCulture));

      foreach (var value in _sharedStrings)
      {
        writer.WriteStartElement("si", Ns.SpreadsheetMl);
        writer.WriteStartElement("t", Ns.SpreadsheetMl);
        writer.WriteString(value);
        writer.WriteEndElement();
        writer.WriteEndElement();
      }

      writer.WriteEndElement();
      writer.WriteEndDocument();
    }

    private void WriteStyles(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/styles.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement(string.Empty, "styleSheet", Ns.SpreadsheetMl);

      // numFmts
      writer.WriteStartElement("numFmts", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "0");
      writer.WriteEndElement();

      // fonts
      writer.WriteStartElement("fonts", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "1");
      writer.WriteStartElement("font", Ns.SpreadsheetMl);
      writer.WriteStartElement("sz", Ns.SpreadsheetMl);
      writer.WriteAttributeString("val", "11");
      writer.WriteEndElement();
      writer.WriteStartElement("name", Ns.SpreadsheetMl);
      writer.WriteAttributeString("val", "Calibri");
      writer.WriteEndElement();
      writer.WriteStartElement("family", Ns.SpreadsheetMl);
      writer.WriteAttributeString("val", "2");
      writer.WriteEndElement();
      writer.WriteStartElement("scheme", Ns.SpreadsheetMl);
      writer.WriteAttributeString("val", "minor");
      writer.WriteEndElement();
      writer.WriteEndElement(); // font
      writer.WriteEndElement(); // fonts

      // fills
      writer.WriteStartElement("fills", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "1");
      writer.WriteStartElement("fill", Ns.SpreadsheetMl);
      writer.WriteStartElement("patternFill", Ns.SpreadsheetMl);
      writer.WriteAttributeString("patternType", "none");
      writer.WriteEndElement();
      writer.WriteEndElement();
      writer.WriteEndElement();

      // borders
      writer.WriteStartElement("borders", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "1");
      writer.WriteStartElement("border", Ns.SpreadsheetMl);
      foreach (var tag in new[] { "left", "right", "top", "bottom", "diagonal" })
      {
        writer.WriteStartElement(tag, Ns.SpreadsheetMl);
        writer.WriteEndElement();
      }

      writer.WriteEndElement(); // border
      writer.WriteEndElement(); // borders

      // cellStyleXfs
      writer.WriteStartElement("cellStyleXfs", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "1");
      writer.WriteStartElement("xf", Ns.SpreadsheetMl);
      writer.WriteAttributeString("numFmtId", "0");
      writer.WriteAttributeString("fontId", "0");
      writer.WriteAttributeString("fillId", "0");
      writer.WriteAttributeString("borderId", "0");
      writer.WriteEndElement();
      writer.WriteEndElement();

      // cellXfs — one xf per style produced by CreateOrGetStylIndex
      writer.WriteStartElement("cellXfs", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", _numberFormatIds.Count.ToString(CultureInfo.InvariantCulture));
      foreach (var numberFormatId in _numberFormatIds)
      {
        writer.WriteStartElement("xf", Ns.SpreadsheetMl);
        writer.WriteAttributeString("numFmtId", numberFormatId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("fontId", "0");
        writer.WriteAttributeString("fillId", "0");
        writer.WriteAttributeString("borderId", "0");
        writer.WriteAttributeString("xfId", "0");
        writer.WriteAttributeString("applyNumberFormat", "1");
        writer.WriteAttributeString("applyFont", "1");
        writer.WriteAttributeString("applyBorder", "1");
        writer.WriteEndElement();
      }

      writer.WriteEndElement(); // cellXfs

      // cellStyles
      writer.WriteStartElement("cellStyles", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "1");
      writer.WriteStartElement("cellStyle", Ns.SpreadsheetMl);
      writer.WriteAttributeString("name", "Normal");
      writer.WriteAttributeString("xfId", "0");
      writer.WriteAttributeString("builtinId", "0");
      writer.WriteEndElement();
      writer.WriteEndElement();

      // tableStyles
      writer.WriteStartElement("tableStyles", Ns.SpreadsheetMl);
      writer.WriteAttributeString("count", "0");
      writer.WriteAttributeString("defaultTableStyle", "TableStyleMedium9");
      writer.WriteAttributeString("defaultPivotStyle", "PivotStyleLight16");
      writer.WriteEndElement();

      writer.WriteEndElement(); // styleSheet
      writer.WriteEndDocument();
    }

    private void WriteWorkbook(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement(string.Empty, "workbook", Ns.SpreadsheetMl);
      writer.WriteAttributeString("xmlns", "r", null, Ns.Relationships);

      writer.WriteStartElement("bookViews", Ns.SpreadsheetMl);
      writer.WriteStartElement("workbookView", Ns.SpreadsheetMl);
      writer.WriteEndElement();
      writer.WriteEndElement();

      writer.WriteStartElement("sheets", Ns.SpreadsheetMl);
      var sheetId = 1U;
      var rId = 3;
      foreach (var ws in _workbookDfn.Worksheets)
      {
        writer.WriteStartElement("sheet", Ns.SpreadsheetMl);
        writer.WriteAttributeString("name", ws.Name);
        writer.WriteAttributeString("sheetId", sheetId.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("r", "id", Ns.Relationships, $"rId{rId}");
        writer.WriteEndElement();
        sheetId++;
        rId++;
      }

      writer.WriteEndElement(); // sheets

      writer.WriteEndElement(); // workbook
      writer.WriteEndDocument();
    }

    private void WriteWorkbookRelationships(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Optimal);
      using var stream = entry.Open();
      var settings = new XmlWriterSettings
      {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        CloseOutput = false,
      };
      using var writer = XmlWriter.Create(stream, settings);

      writer.WriteStartDocument(standalone: true);
      writer.WriteStartElement("Relationships", Ns.PackageRelationships);

      WriteRelationship(
        writer,
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "styles.xml");

      WriteRelationship(
        writer,
        "rId2",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "sharedStrings.xml");

      var rId = 3;
      for (var i = 1; i <= _workbookDfn.Worksheets.Count; i++)
      {
        WriteRelationship(
          writer,
          $"rId{rId}",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
          $"worksheets/sheet{i}.xml");
        rId++;
      }

      writer.WriteEndElement();
      writer.WriteEndDocument();
    }

    private void WriteWorksheets(ZipArchive archive)
    {
      var count = 1U;
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        var entry = archive.CreateEntry($"xl/worksheets/sheet{count}.xml", CompressionLevel.Optimal);
        using var stream = entry.Open();
        var settings = new XmlWriterSettings
        {
          Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
          CloseOutput = false,
        };
        using var writer = XmlWriter.Create(stream, settings);

        var (sheetData, maxColumnCount, lastRowIndex) = GenerateSheetDataForDetails(worksheet);

        var reference = lastRowIndex == 0U || maxColumnCount == 0
          ? "A1"
          : $"A1:{ColumnReferenceHelper.ToLetters(maxColumnCount)}{lastRowIndex}";

        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement(string.Empty, "worksheet", Ns.SpreadsheetMl);

        writer.WriteStartElement("dimension", Ns.SpreadsheetMl);
        writer.WriteAttributeString("ref", reference);
        writer.WriteEndElement();

        writer.WriteStartElement("sheetViews", Ns.SpreadsheetMl);
        writer.WriteStartElement("sheetView", Ns.SpreadsheetMl);
        writer.WriteAttributeString("tabSelected", count == 1U ? "1" : "0");
        writer.WriteAttributeString("workbookViewId", "0");
        writer.WriteStartElement("selection", Ns.SpreadsheetMl);
        writer.WriteAttributeString("activeCell", "A1");
        writer.WriteAttributeString("sqref", "A1");
        writer.WriteEndElement();
        writer.WriteEndElement(); // sheetView
        writer.WriteEndElement(); // sheetViews

        writer.WriteStartElement("sheetFormatPr", Ns.SpreadsheetMl);
        writer.WriteAttributeString("defaultRowHeight", "15");
        writer.WriteAttributeString("defaultColWidth", "15");
        writer.WriteEndElement();

        // sheetData
        writer.WriteStartElement("sheetData", Ns.SpreadsheetMl);
        foreach (var row in sheetData.Elements<Row>())
        {
          writer.WriteStartElement("row", Ns.SpreadsheetMl);
          if (row.RowIndex != null)
          {
            writer.WriteAttributeString("r", row.RowIndex.Value.ToString(CultureInfo.InvariantCulture));
          }

          foreach (var cell in row.Elements<Cell>())
          {
            writer.WriteStartElement("c", Ns.SpreadsheetMl);
            if (cell.CellReference != null)
            {
              writer.WriteAttributeString("r", cell.CellReference.Value);
            }

            if (cell.StyleIndex != null)
            {
              writer.WriteAttributeString("s", cell.StyleIndex.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (cell.DataType != null)
            {
              writer.WriteAttributeString("t", ToCellTypeAttribute(cell.DataType.Value));
            }

            if (cell.CellValue != null)
            {
              writer.WriteStartElement("v", Ns.SpreadsheetMl);
              writer.WriteString(cell.CellValue.Text);
              writer.WriteEndElement();
            }

            if (cell.InlineString != null)
            {
              writer.WriteStartElement("is", Ns.SpreadsheetMl);
              writer.WriteStartElement("t", Ns.SpreadsheetMl);
              writer.WriteString(cell.InlineString.Text?.Text ?? string.Empty);
              writer.WriteEndElement();
              writer.WriteEndElement();
            }

            writer.WriteEndElement(); // c
          }

          writer.WriteEndElement(); // row
        }

        writer.WriteEndElement(); // sheetData

        writer.WriteStartElement("pageMargins", Ns.SpreadsheetMl);
        writer.WriteAttributeString("left", "0.7");
        writer.WriteAttributeString("right", "0.7");
        writer.WriteAttributeString("top", "0.75");
        writer.WriteAttributeString("bottom", "0.75");
        writer.WriteAttributeString("header", "0.3");
        writer.WriteAttributeString("footer", "0.3");
        writer.WriteEndElement();

        writer.WriteEndElement(); // worksheet
        writer.WriteEndDocument();

        count++;
      }
    }

    private static class Ns
    {
      public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

      public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

      public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

      public const string SpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    }
  }
}
