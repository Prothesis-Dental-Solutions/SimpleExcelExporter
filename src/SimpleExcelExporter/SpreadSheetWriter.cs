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
  using System.Text.RegularExpressions;
  using System.Xml;
  using System.Xml.Linq;
  using DocumentFormat.OpenXml;
  using DocumentFormat.OpenXml.Packaging;
  using DocumentFormat.OpenXml.Spreadsheet;
  using SimpleExcelExporter.Annotations;
  using SimpleExcelExporter.Definitions;
  using SimpleExcelExporter.Resources;

  public partial class SpreadsheetWriter
  {
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string SpreadsheetMlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private readonly Dictionary<string, Attribute?> _cachedAttributes = [];

    private readonly Dictionary<string, (CellDfn, bool)> _headers = [];

    private readonly Dictionary<string, int> _multiColumnAttribute = [];

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

    private Dictionary<int, uint> Table { get; } = [];

    public void Write()
    {
      using var buffer = new MemoryStream();

      // Adding core file properties is mandatory to avoid a problem with Google Spreadsheet transforming from XLSX to XLSM.
      // cf. https://stackoverflow.com/questions/70319867/avoid-google-spreadsheet-to-convert-an-xlsx-file-created-by-open-xml-sdk-to-xlsm
      // cf. https://github.com/OfficeDev/Open-XML-SDK/issues/1093
      // cf. https://issuetracker.google.com/issues/210875597
      using (var document = SpreadsheetDocument.Create(buffer, SpreadsheetDocumentType.Workbook))
      {
        var coreFilePropPart = document.AddCoreFilePropertiesPart();
        using (var writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
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

        CreatePartsForExcel(document);
      }

      buffer.Position = 0;
      FixContentTypesXml(buffer);

      buffer.Position = 0;
      FixNamespacePrefixes(buffer);

      buffer.Position = 0;
      FixRelationshipTargets(buffer);

      buffer.Position = 0;
      FixWorkbookNamespaceDeclaration(buffer);

      buffer.Position = 0;
      buffer.CopyTo(_stream);
    }

    [GeneratedRegex("Target=\"/([^\"]+)\"")]
    private static partial Regex AbsoluteTargetRegex();

    [GeneratedRegex("Id=\"R[0-9a-fA-F]{16}\"")]
    private static partial Regex GuidIdRegex();

    [GeneratedRegex("Id=\"rId(\\d+)\"")]
    private static partial Regex ExistingRelIdRegex();

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

    private static void FixContentTypesXml(MemoryStream buffer)
    {
      using var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true);
      var entry = archive.GetEntry("[Content_Types].xml");
      if (entry == null)
      {
        return;
      }

      XDocument doc;
      using (var entryStream = entry.Open())
      {
        doc = XDocument.Load(entryStream);
      }

      var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types");
      var root = doc.Root!;

      var xmlDefault = root.Elements(ns + "Default")
        .FirstOrDefault(d => (string?)d.Attribute("Extension") == "xml");

      if (xmlDefault != null && (string?)xmlDefault.Attribute("ContentType") != "application/xml")
      {
        var displacedContentType = (string?)xmlDefault.Attribute("ContentType");
        var displacedPartName = displacedContentType switch
        {
          "application/vnd.openxmlformats-package.core-properties+xml" => "/docProps/core.xml",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => "/xl/workbook.xml",
          _ => null,
        };

        if (displacedPartName != null)
        {
          var alreadyHasOverride = root.Elements(ns + "Override")
            .Any(o => (string?)o.Attribute("PartName") == displacedPartName);

          if (!alreadyHasOverride)
          {
            root.Add(new XElement(
              ns + "Override",
              new XAttribute("PartName", displacedPartName),
              new XAttribute("ContentType", displacedContentType!)));
          }
        }

        xmlDefault.SetAttributeValue("ContentType", "application/xml");
      }

      entry.Delete();
      var newEntry = archive.CreateEntry("[Content_Types].xml");
      using var entryWriter = new StreamWriter(newEntry.Open(), new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
      doc.Save(entryWriter);
    }

    private static void FixNamespacePrefixes(MemoryStream buffer)
    {
      const string SpreadsheetMlNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      var prefixedDeclaration = $"xmlns:x=\"{SpreadsheetMlNamespace}\"";
      var defaultDeclaration = $"xmlns=\"{SpreadsheetMlNamespace}\"";

      using var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true);

      var entriesToFix = archive.Entries
        .Where(e => e.FullName.StartsWith("xl/", StringComparison.Ordinal)
          && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
          && !e.FullName.Contains("_rels/", StringComparison.Ordinal))
        .ToList();

      foreach (var entry in entriesToFix)
      {
        string content;
        using (var entryStream = entry.Open())
        using (var reader = new StreamReader(entryStream, System.Text.Encoding.UTF8))
        {
          content = reader.ReadToEnd();
        }

        if (!content.Contains(prefixedDeclaration, StringComparison.Ordinal))
        {
          continue;
        }

        content = content
          .Replace(prefixedDeclaration, defaultDeclaration, StringComparison.Ordinal)
          .Replace("<x:", "<", StringComparison.Ordinal)
          .Replace("</x:", "</", StringComparison.Ordinal);

        entry.Delete();
        var newEntry = archive.CreateEntry(entry.FullName);
        using var writerStream = newEntry.Open();
        using var writer = new StreamWriter(writerStream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
      }
    }

    private static void FixRelationshipTargets(MemoryStream buffer)
    {
      using var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true);

      var relsEntries = archive.Entries
        .Where(e => e.FullName.EndsWith(".rels", StringComparison.Ordinal))
        .ToList();

      foreach (var entry in relsEntries)
      {
        string content;
        using (var entryStream = entry.Open())
        using (var reader = new StreamReader(entryStream, System.Text.Encoding.UTF8))
        {
          content = reader.ReadToEnd();
        }

        // Compute base directory: for "_rels/.rels" base is ""; for "xl/_rels/workbook.xml.rels" base is "xl/"
        var relsIndex = entry.FullName.LastIndexOf("_rels/", StringComparison.Ordinal);
        var baseDir = relsIndex == 0 ? string.Empty : entry.FullName.Substring(0, relsIndex);

        // Transform absolute targets to relative
        // Strategy: replace Target="/x/y/z" with the part relative to baseDir
        content = AbsoluteTargetRegex().Replace(content, match =>
        {
          var absTarget = match.Groups[1].Value;
          if (baseDir.Length > 0 && absTarget.StartsWith(baseDir, StringComparison.Ordinal))
          {
            return $"Target=\"{absTarget.Substring(baseDir.Length)}\"";
          }

          return $"Target=\"{absTarget}\"";
        });

        // Normalize SDK-generated GUID-style IDs (Id="R<16 hex>") to sequential rIdN
        var maxExistingId = 0;
        foreach (Match existing in ExistingRelIdRegex().Matches(content))
        {
          var num = int.Parse(existing.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
          if (num > maxExistingId)
          {
            maxExistingId = num;
          }
        }

        var nextId = maxExistingId;
        content = GuidIdRegex().Replace(content, _ =>
        {
          nextId++;
          return $"Id=\"rId{nextId}\"";
        });

        entry.Delete();
        var newEntry = archive.CreateEntry(entry.FullName);
        using var writerStream = newEntry.Open();
        using var writer = new StreamWriter(writerStream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
      }
    }

    private static void FixWorkbookNamespaceDeclaration(MemoryStream buffer)
    {
      const string RelsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      var relsDeclaration = $" xmlns:r=\"{RelsNamespace}\"";

      using var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true);
      var entry = archive.GetEntry("xl/workbook.xml");
      if (entry == null)
      {
        return;
      }

      string content;
      using (var entryStream = entry.Open())
      using (var reader = new StreamReader(entryStream, System.Text.Encoding.UTF8))
      {
        content = reader.ReadToEnd();
      }

      // If xmlns:r is already on <workbook>, nothing to do
      var workbookTagEnd = content.IndexOf('>', content.IndexOf("<workbook", StringComparison.Ordinal));
      var workbookTag = content.Substring(0, workbookTagEnd + 1);
      if (workbookTag.Contains("xmlns:r=", StringComparison.Ordinal))
      {
        return;
      }

      // Remove xmlns:r declarations anywhere in the document
      content = content.Replace(relsDeclaration, string.Empty, StringComparison.Ordinal);

      // Add xmlns:r to the <workbook> element
      content = content.Replace(
        $"<workbook xmlns=\"{MainNamespace}\"",
        $"<workbook xmlns:r=\"{RelsNamespace}\" xmlns=\"{MainNamespace}\"",
        StringComparison.Ordinal);

      entry.Delete();
      var newEntry = archive.CreateEntry(entry.FullName);
      using var writerStream = newEntry.Open();
      using var writer = new StreamWriter(writerStream, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
      writer.Write(content);
    }

    private static void GenerateWorksheetPartContent(
      WorksheetPart worksheetPart,
      SheetData sheetData,
      bool tabSelectedFlag,
      int maxColumnCount,
      uint lastRowIndex)
    {
      var worksheet = new Worksheet();
      var reference = lastRowIndex == 0U || maxColumnCount == 0
        ? "A1"
        : $"A1:{ColumnReferenceHelper.ToLetters(maxColumnCount)}{lastRowIndex}";
      var sheetDimension = new SheetDimension { Reference = reference };

      var sheetViews = new SheetViews();

      var sheetView = new SheetView { TabSelected = tabSelectedFlag, WorkbookViewId = 0U };
      var selection = new Selection { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" } };

      _ = sheetView.AppendChild(selection);

      _ = sheetViews.AppendChild(sheetView);
      var sheetFormatProperties = new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 15D };

      var pageMargins = new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
      _ = worksheet.AppendChild(sheetDimension);
      _ = worksheet.AppendChild(sheetViews);
      _ = worksheet.AppendChild(sheetFormatProperties);
      _ = worksheet.AppendChild(sheetData);
      _ = worksheet.AppendChild(pageMargins);
      worksheetPart.Worksheet = worksheet;
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

    private static void WriteOverride(XmlWriter writer, string partName, string contentType)
    {
      writer.WriteStartElement("Override", ContentTypesNamespace);
      writer.WriteAttributeString("PartName", partName);
      writer.WriteAttributeString("ContentType", contentType);
      writer.WriteEndElement();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
      writer.WriteStartElement("Relationship", PackageRelationshipsNamespace);
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
        (var currentPlayer, var currentPlayerTypePropertyInfos, var currentIteration, var currentParentIndex) = objectQueue.Dequeue();

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
              if (childPlayerType != null && childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
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
        (var currentPlayer, var currentPlayerType, var currentPlayerTypePropertyInfos, var currentIteration, var currentParentIndex) = objectQueue.Dequeue();

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
              if (childPlayerType != null && childPlayerTypePropertyInfos != null && numberOfEmptyCellToAdd > 0)
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
              (var headerCellDfn, var textCorrectlySetFlag) = value;
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

    private void BuildStylesheetSkeleton()
    {
      // Number formats (empty — using built-in formats only)
      var numberingFormats = new NumberingFormats { Count = 0U };

      var fonts = new Fonts { Count = 1U };

      // Font 1
      var font = new Font
      {
        FontSize = new FontSize { Val = 11D },
        FontName = new FontName { Val = "Calibri" },
        FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
        FontScheme = new FontScheme { Val = FontSchemeValues.Minor },
      };

      _ = fonts.AppendChild(font);

      // Default Fill
      var fills = new Fills { Count = 1U };
      var fill = new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } };
      _ = fills.AppendChild(fill);

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
      _ = borders.AppendChild(border);

      // CellStyleFormats
      var cellStyleFormats = new CellStyleFormats { Count = 1U };
      var cellFormat = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
      _ = cellStyleFormats.AppendChild(cellFormat);

      // CellFormats — empty, populated by CreateOrGetStylIndex during worksheet processing
      var cellFormats = new CellFormats { Count = 0U };

      // CellStyles — Excel requires the "Normal" built-in style
      var cellStyles = new CellStyles { Count = 1U };
      _ = cellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });

      // TableStyles — empty but required by strict parsers
      var tableStyles = new TableStyles
      {
        Count = 0U,
        DefaultTableStyle = "TableStyleMedium9",
        DefaultPivotStyle = "PivotStyleLight16",
      };

      _ = _stylesheet.AppendChild(numberingFormats);
      _ = _stylesheet.AppendChild(fonts);
      _ = _stylesheet.AppendChild(fills);
      _ = _stylesheet.AppendChild(borders);
      _ = _stylesheet.AppendChild(cellStyleFormats);
      _ = _stylesheet.AppendChild(cellFormats);
      _ = _stylesheet.AppendChild(cellStyles);
      _ = _stylesheet.AppendChild(tableStyles);
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
        cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
        if (!string.IsNullOrEmpty(stringValue))
        {
          stringValue = XmlStringHelper.Sanitize(stringValue);
          cell.InlineString = new InlineString { Text = new Text(stringValue) };
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
      _ = _stylesheet.CellFormats.AppendChild(cellFormat);
      Table.Add(styleHashCode, index);

      return index;
    }

    private void CreatePartsForExcel(SpreadsheetDocument document)
    {
      var workbookPart = document.AddWorkbookPart();
      var workbook = new Workbook();
      workbook.Append(new BookViews(new WorkbookView()));
      workbookPart.Workbook = workbook;
      var sheets = new Sheets();
      _ = workbook.AppendChild(sheets);

      // Styles first, then worksheets — rId2 reserved for styles, rId3+ for sheets
      var workbookStylesPart1 = workbookPart.AddNewPart<WorkbookStylesPart>("rId2");
      GenerateWorkbookStylesPartContent(workbookStylesPart1);

      // Thank you https://stackoverflow.com/questions/9120544/openxml-multiple-sheets
      var count = 1U;
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>($"rId{count + 2}");  // rId3, rId4, ...
        var sheet = new Sheet { Name = worksheet.Name, SheetId = count, Id = workbookPart.GetIdOfPart(worksheetPart) };
        _ = sheets.AppendChild(sheet);
        var (sheetData, maxColumnCount, lastRowIndex) = GenerateSheetDataForDetails(worksheet);
        GenerateWorksheetPartContent(worksheetPart, sheetData, count == 1U, maxColumnCount, lastRowIndex);
        count++;
      }
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

    private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
    {
      // Number formats (empty — using built-in formats only)
      var numberingFormats = new NumberingFormats { Count = 0U };

      var fonts = new Fonts { Count = 1U };

      // Font 1
      var font = new Font
      {
        FontSize = new FontSize { Val = 11D },
        FontName = new FontName { Val = "Calibri" },
        FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
        FontScheme = new FontScheme { Val = FontSchemeValues.Minor },
      };

      _ = fonts.AppendChild(font);

      // Default Fill
      var fills = new Fills { Count = 1U };
      var fill = new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } };
      _ = fills.AppendChild(fill);

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
      _ = borders.AppendChild(border);

      // CellStyleFormats
      var cellStyleFormats = new CellStyleFormats { Count = 1U };
      var cellFormat = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
      _ = cellStyleFormats.AppendChild(cellFormat);

      // CellFormats
      var cellFormats = new CellFormats { Count = 0U };

      // CellStyles — Excel requires the "Normal" built-in style
      var cellStyles = new CellStyles { Count = 1U };
      _ = cellStyles.AppendChild(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });

      // TableStyles — empty but required by strict parsers
      var tableStyles = new TableStyles
      {
        Count = 0U,
        DefaultTableStyle = "TableStyleMedium9",
        DefaultPivotStyle = "PivotStyleLight16",
      };

      _ = _stylesheet.AppendChild(numberingFormats);
      _ = _stylesheet.AppendChild(fonts);
      _ = _stylesheet.AppendChild(fills);
      _ = _stylesheet.AppendChild(borders);
      _ = _stylesheet.AppendChild(cellStyleFormats);
      _ = _stylesheet.AppendChild(cellFormats);
      _ = _stylesheet.AppendChild(cellStyles);
      _ = _stylesheet.AppendChild(tableStyles);

      workbookStylesPart.Stylesheet = _stylesheet;
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
      writer.WriteStartElement("Types", ContentTypesNamespace);

      writer.WriteStartElement("Default", ContentTypesNamespace);
      writer.WriteAttributeString("Extension", "xml");
      writer.WriteAttributeString("ContentType", "application/xml");
      writer.WriteEndElement();

      writer.WriteStartElement("Default", ContentTypesNamespace);
      writer.WriteAttributeString("Extension", "rels");
      writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
      writer.WriteEndElement();

      WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
      WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
      WriteOverride(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");

      for (var i = 1; i <= _workbookDfn.Worksheets.Count; i++)
      {
        WriteOverride(writer, $"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
      }

      writer.WriteEndElement();
      writer.WriteEndDocument();
    }

    private void WriteCoreProperties(ZipArchive archive)
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

    private void WritePackageRels(ZipArchive archive)
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
      writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);

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

    private void WriteStyles(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/styles.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      using var writer = OpenXmlWriter.Create(stream, Encoding.UTF8);

      writer.WriteStartDocument(standalone: true);
      writer.WriteElement(_stylesheet);
    }

    private void WriteWorkbook(ZipArchive archive)
    {
      var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
      using var stream = entry.Open();
      using var writer = OpenXmlWriter.Create(stream, Encoding.UTF8);

      var workbook = new Workbook();
      workbook.Append(new BookViews(new WorkbookView()));

      var sheets = new Sheets();
      _ = workbook.AppendChild(sheets);

      var sheetId = 1U;
      var rId = 2;
      foreach (var worksheet in _workbookDfn.Worksheets)
      {
        _ = sheets.AppendChild(new Sheet
        {
          Name = worksheet.Name,
          SheetId = sheetId,
          Id = $"rId{rId}",
        });
        sheetId++;
        rId++;
      }

      writer.WriteStartDocument(standalone: true);

      var namespaceDeclarations = new List<KeyValuePair<string, string>>
      {
        new("r", RelationshipsNamespace),
      };

      writer.WriteStartElement(workbook, Array.Empty<OpenXmlAttribute>(), namespaceDeclarations);
      foreach (var child in workbook.ChildElements)
      {
        writer.WriteElement(child);
      }

      writer.WriteEndElement();
    }

    private void WriteWorkbookRels(ZipArchive archive)
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
      writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);

      WriteRelationship(
        writer,
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "styles.xml");

      var rId = 2;
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
        using var writer = OpenXmlWriter.Create(stream, Encoding.UTF8);

        var (sheetData, maxColumnCount, lastRowIndex) = GenerateSheetDataForDetails(worksheet);

        var reference = lastRowIndex == 0U || maxColumnCount == 0
          ? "A1"
          : $"A1:{ColumnReferenceHelper.ToLetters(maxColumnCount)}{lastRowIndex}";
        var sheetDimension = new SheetDimension { Reference = reference };

        var sheetViews = new SheetViews();
        var sheetView = new SheetView { TabSelected = count == 1U, WorkbookViewId = 0U };
        var selection = new Selection { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" } };
        _ = sheetView.AppendChild(selection);
        _ = sheetViews.AppendChild(sheetView);

        var sheetFormatProperties = new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 15D };
        var pageMargins = new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

        var worksheetElement = new Worksheet();
        _ = worksheetElement.AppendChild(sheetDimension);
        _ = worksheetElement.AppendChild(sheetViews);
        _ = worksheetElement.AppendChild(sheetFormatProperties);
        _ = worksheetElement.AppendChild(sheetData);
        _ = worksheetElement.AppendChild(pageMargins);

        writer.WriteStartDocument(standalone: true);
        writer.WriteElement(worksheetElement);

        count++;
      }
    }
  }
}
