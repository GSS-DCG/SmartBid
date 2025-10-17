using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Field = Microsoft.Office.Interop.Word.Field;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace SmartBid
{
  public class MirrorXML
  {
    private string _fileName;
    private Dictionary<string, string[]> _varList;
    // Add the following field and property to the MirrorXML class to fix CS0246

    private Dictionary<string, string> _basicData;

    public Dictionary<string, string[]> VarList
    {
      get => _varList;
      set => _varList = value;
    }

    public string FileName
    {
      get => _fileName;
      set => _fileName = value;
    }

    public Dictionary<string, string> BasicData
    {
      get => _basicData;
      set => _basicData = value;
    }
    public MirrorXML(ToolData tool)
    {
      _varList = [];
      _basicData = [];


      //tool = Regex.Replace(tool, "_Call\\d+", "");

      string? toolResource = tool.Resource;
      string toolsPath = Path.Combine(H.GetSProperty("ToolsPath"));
      string templatesPath = Path.Combine(H.GetSProperty("TemplatesPath"));
      string? directoryPath;

      if ((toolResource != null) && (toolResource == "TOOL"))
      {
        directoryPath = toolsPath;
      }
      else if (toolResource == "TEMPLATE")
      {
        directoryPath = templatesPath;
      }
      else
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌  - LoadData", $"Tool Resource not found for file: {tool.Code}:{toolResource}");
        return;
      }


      FileName = tool.FileName;
      string toolPath = Path.Combine(directoryPath, FileName);
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(toolPath);
      string xmlFile = Path.Combine(directoryPath, fileNameWithoutExtension + ".xml");
      DateTime fileModified = new(2000,01,01,00,00,00); //Set default to a old date so for .exe will always use the XML
      DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;

      //if tool.filetype is docx, xlsx or xlsm, check if the file exists
      if (tool.FileType == "docx" || tool.FileType == "xlsx" || tool.FileType == "xlsm")
      {
        if (File.Exists(toolPath))
        {
          fileModified = File.GetLastWriteTime(toolPath);
        }
        else
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "LoadData", $"{toolPath} does not exist.");
          return;
        }
      }

      if (fileModified > xmlModified)
      {
        VariablesMap varMap = VariablesMap.Instance;

        if (tool.FileType.Equals("docx", StringComparison.CurrentCultureIgnoreCase))
        {
          BasicData = ExtractGSSDataFromDocx(toolPath);
          VarList = ExtractVariablesFromDocx(toolPath);
        }

        else if (tool.FileType.ToLower().StartsWith("xls"))
        {
          BasicData = ExtractGSSDataFromXlsx(toolPath);
          VarList = ExtractVariablesFromXlsx(toolPath);
        }

        foreach (string variable in VarList.Keys.ToList())
        {
          if (!variable.Contains("\\s")) //si la llamada no es la primera
          {
            string id = System.Text.RegularExpressions.Regex.Replace(variable, @"^Call[0-9]_", "");

            if (varMap.GetNewVariableData(id) != null)
            {
              VarList[variable] = [
                varMap.GetNewVariableData(id)!.Source, 
                VarList[variable][1], 
                VarList[variable][2], 
                VarList[variable][3]
                ]; // Fill up source data before saving to XML
            }
          }
        }

        CreateXMLMirror(directoryPath);
      }

      if (File.Exists(xmlFile))
      {
        XElement root = XElement.Load(xmlFile);

        //reading basicData from GSS_DATA attributes
        root.Attributes().ToList().ForEach(attr => _basicData[attr.Name.LocalName] = attr.Value);

        //reading special attribute FileName for .exe Tools
        //if fileName attribute exists set FileName value 
        if (root.Attribute("fileName") != null)
        {
          FileName = root.Attribute("fileName")!.Value;
        }

        foreach (var variable in root.Elements("variable"))
        {
          VarList[variable.Value] =
          [ variable?.Attribute("source")?.Value ?? "",
            variable?.Attribute("inOut")?.Value ?? "",
            variable?.Attribute("call")?.Value ?? "",
            variable?.Attribute("type")?.Value ?? ""
          ];
        }
      }
    }

    static Dictionary<string, string> ExtractGSSDataFromDocx(string docxPath)
    {
      Dictionary<string, string> bookmarks = new();

      try
      {
        using WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false);
        var body = doc.MainDocumentPart.Document.Body;
        var bookmarksList = body.Descendants<BookmarkStart>();

        foreach (var bookmark in bookmarksList)
        {
          string bookmarkName = bookmark.Name!;
          string bookmarkValue = "";

          // Buscar el siguiente nodo que sea un `Run`
          var currentElement = bookmark.NextSibling();
          while (currentElement != null)
          {
            var run = currentElement as Run;
            if (run != null)
            {
              var textElements = run.Descendants<Text>();
              foreach (var textElement in textElements)
              {
                bookmarkValue += textElement.Text;
              }
            }
            currentElement = currentElement.NextSibling();
          }

          bookmarks[bookmarkName] = bookmarkValue.Trim();
        }
      }
      catch (FileNotFoundException)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ExtractGSSDataFromDocx", $"File not found: {docxPath}");
      }
      catch (Exception ex)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ExtractGSSDataFromDocx", $"An error occurred: {ex.Message}");
      }

      bookmarks = bookmarks.Where(kvp => kvp.Key.StartsWith("GSS_"))
          .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

      return bookmarks;
    }

    private Dictionary<string, string[]> ExtractVariablesFromDocx(string fileName)
    {
      List<string> varList = [];
      List<string> bookmarkList = [];
      string varPrefix = H.GetSProperty("VarPrefix");

      Application wordApp = new();
      Microsoft.Office.Interop.Word.Document? doc = null;

      try
      {
        wordApp.Visible = false;
        doc = wordApp.Documents.Open(fileName, ReadOnly: true, Visible: false);

        // --- Extract cross-reference fields ---
        foreach (Field field in doc.Fields)
        {
          if (field.Type == WdFieldType.wdFieldRef || field.Type == WdFieldType.wdFieldEmpty)
          {
            string fieldCode = field.Code.Text.Trim();

            if (!string.IsNullOrWhiteSpace(fieldCode) &&
                fieldCode.StartsWith(varPrefix, StringComparison.OrdinalIgnoreCase))
            {
              string cleaned = fieldCode
                  .Replace(varPrefix, "")
                  .Replace("\\* MERGEFORMAT", "", StringComparison.OrdinalIgnoreCase)
                  .Replace("REF ", "", StringComparison.OrdinalIgnoreCase)
                  .Trim();

              if (!string.IsNullOrWhiteSpace(cleaned))
                {
                  cleaned = cleaned.Split('\\')[0];
                  varList.Add(cleaned); 
                }
            }
          }

        }
        varList = varList.Distinct().ToList();
        varList = varList.Select(item => item.Contains("/") ? item[(item.IndexOf('/') + 1)..] : item).ToList();

        // --- Extract bookmarks ---
        foreach (Bookmark bookmark in doc.Bookmarks)
        {
          if (!string.IsNullOrWhiteSpace(bookmark.Name) &&
              bookmark.Name.StartsWith(varPrefix, StringComparison.OrdinalIgnoreCase))
          {
            string cleaned = bookmark.Name.Replace(varPrefix, "").Trim();
            if (!string.IsNullOrWhiteSpace(cleaned))
              bookmarkList.Add(cleaned);
          }
        }

        // --- Build dictionaries ---
        var crossRefDict = varList
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                key => key,
                key => new string[] { "", "in", "1", "cross-reference" },
                StringComparer.OrdinalIgnoreCase);

        var bookmarkDict = bookmarkList
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                key => key,
                key => new string[] { "", "in", "1", "bookmark" },
                StringComparer.OrdinalIgnoreCase);

        var mergedDict = crossRefDict
            .Concat(bookmarkDict)
            .ToDictionary(
                pair => pair.Key,
                pair => pair.Value,
                StringComparer.OrdinalIgnoreCase);

        VariablesMap varMap = VariablesMap.Instance;
        List<string> nonDeclaredVars = [.. mergedDict.Keys.Where(v => !varMap.IsVariableExists(v))];

        if (nonDeclaredVars.Count > 0)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "❌❌ Error ❌❌  - ExtractVariablesFromDocx", "Declaration Error");
          throw new InvalidOperationException(
              $"{nonDeclaredVars.Count} variables found in {Path.GetFileName(fileName)} are not declared in VariableMap:\n\n{string.Join("\n", nonDeclaredVars)}\n");
        }

        return mergedDict;
      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "❌❌ Error ❌❌  - ExtractVariablesFromDocx", $"❌ Error reading DOCX: {ex.Message}");
        return [];
      }
      finally
      {
        if (doc != null)
        {
          doc.Close(SaveChanges: false);
          Marshal.ReleaseComObject(doc);
        }

        wordApp.Quit();
        Marshal.ReleaseComObject(wordApp);
      }
    }

    private static Dictionary<string, string> ExtractGSSDataFromXlsx(string fileName)
    {
      var gssData = new Dictionary<string, string>();
      using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
      {
        WorkbookPart workbookPart = document.WorkbookPart!;
        var listData = GetCellValuesFromRange(workbookPart, "GSS_DATA");

        foreach (var row in listData)
        {
          string key = row[0];
          string value = row[1];
          gssData[key] = value;
        }
      }
      return gssData;
    }

    private static Dictionary<string, string[]> ExtractVariablesFromXlsx(string fileName)
    {
      Dictionary<string, string[]> varList = [];
      List<string> varNames;
      VariablesMap varMap = VariablesMap.Instance;

      using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
      {
        WorkbookPart workbookPart = document.WorkbookPart!;
        varNames = GetAllRangeNames(workbookPart);
      }

      string inPrefix1 = H.GetSProperty("IN_VarPrefix").ToLower();
      string inPrefix2 = H.GetSProperty("VarPrefix").ToLower();
      string outPrefix = H.GetSProperty("OUT_VarPrefix").ToLower();

      // Filtrar la lista
      varNames = [.. varNames
          .Where(name => name.StartsWith(inPrefix1, StringComparison.OrdinalIgnoreCase) ||
                         name.StartsWith(inPrefix2, StringComparison.OrdinalIgnoreCase) ||
                         name.StartsWith(outPrefix, StringComparison.OrdinalIgnoreCase))];

      varNames.Remove("GSS_DATA"); // Remove GSS_DATA from the list

      // generating a list with variables missing in varMap
      List<string> nonDeclaredVars = [];

      foreach (string item in varNames)
      {
        // eliminamos la marca de lista /xx
        string varName = item.Split('\\')[0];

        string[] value = ["", "", "1", ""];

        if (varName.ToLower().StartsWith(inPrefix1))
        {
          value[1] = "in";
          varName = varName[inPrefix1.Length..];
        }
        else if (varName.ToLower().StartsWith(inPrefix2))
        {
          value[1] = "in";
          varName = varName[inPrefix2.Length..];
        }
        else if (varName.ToLower().StartsWith(outPrefix))
        {
          value[1] = "out";
          varName = varName[outPrefix.Length..];
        }

        Match match = Regex.Match(varName.ToLower(), @"^call(\d)_");
        if (match.Success)
        {
          value[2] = match.Groups[1].Value;
          varName = varName[match.Length..];
        }

        if (varMap.IsVariableExists(varName))
        {
          value[3] = varMap.GetVariableData(varName).Type;
        }
        else
        {
          nonDeclaredVars.Add(varName);
        }

        if (!varList.ContainsKey(varName))
        {
          varList.Add(new string(varName), value);
        }

      }

      if (nonDeclaredVars.Count > 0)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌  - ExtractVariablesFromXlsx", $"Declaration Error");
        throw new InvalidOperationException(
            $"{nonDeclaredVars.Count} variables found in {Path.GetFileName(fileName)} are not declared in VariableMap:\n\n{string.Join("\n", nonDeclaredVars)}\n"
        );
      }

      return varList;
    }

    private static List<string> GetAllRangeNames(WorkbookPart workbookPart)
    {
      List<string> rangeNames = [];
      try
      {
        var definedNames = workbookPart.Workbook.DefinedNames;
        if (definedNames != null)
        {
          foreach (var definedName in definedNames.Elements<DefinedName>())
          {
            rangeNames.Add(definedName.Name!);
          }
        }
      }
      catch (Exception ex)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetAllRangeNames", $"❌Error❌ reading range names: {ex.Message}");
      }
      return rangeNames;
    }

    private static List<List<string>> GetCellValuesFromRange(WorkbookPart workbookPart, string rangeName)
    {
      List<List<string>> cellValues = [];
      try
      {
        var definedNames = workbookPart.Workbook.DefinedNames;
        if (definedNames != null)
        {
          var gssInputRange = definedNames.Elements<DefinedName>().FirstOrDefault(dn => dn.Name == rangeName);
          if (gssInputRange != null)
          {
            string[] range = gssInputRange.Text.Split('!')[1].Split(':');
            string sheetName = gssInputRange.Text.Split('!')[0].Trim('\'');
            Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName)!;
            if (sheet != null)
            {
              WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
              string startCellReference = range[0].Replace("$", ""); // Remove dollar signs
              string endCellReference = (range.Length == 1) ? startCellReference : range[1].Replace("$", ""); // Remove dollar signs

              var cells = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                  .Where(c => IsCellInRange(c.CellReference!, startCellReference, endCellReference));

              int numberOfColumns = GetColumnRowIndices(endCellReference).column - GetColumnRowIndices(startCellReference).column + 1;
              int numberOfRows = cells.Count() / numberOfColumns;
              int i = 0;
              List<string> row = [];
              foreach (var cell in cells)
              {
                i++;
                if (i <= numberOfColumns)
                {
                  row.Add(GetCellValue(workbookPart, cell));
                }
                else
                {
                  cellValues.Add(row); // Add the row to the main list
                  i = 1;
                  row = [GetCellValue(workbookPart, cell)];
                }
              }
              cellValues.Add(row); // Add the row to the main list
            }
            else
            {
              H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetCellValuesFromRange", $"Sheet {sheetName} not found.");
            }
          }
          else
          {
            H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetCellValuesFromRange", $"Range {rangeName} not found.");
          }
        }
        else
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetCellValuesFromRange", "No defined names found.");
        }
      }
      catch (Exception ex)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetCellValuesFromRange", $"❌Error❌ reading {rangeName} range: {ex.Message}");
      }
      return cellValues;
    }

    private static bool IsCellInRange(string cellReference, string startCellReference, string endCellReference)
    {
      // Convert cell references to row and column indices
      (int startColumn, int startRow) = GetColumnRowIndices(startCellReference);
      (int endColumn, int endRow) = GetColumnRowIndices(endCellReference);
      (int cellColumn, int cellRow) = GetColumnRowIndices(cellReference);

      // Check if the cell is within the specified range
      return cellColumn >= startColumn && cellColumn <= endColumn && cellRow >= startRow && cellRow <= endRow;
    }

    private static (int column, int row) GetColumnRowIndices(string cellReference)
    {
      // Extract column letters and row numbers from the cell reference
      string columnLetters = new ([.. cellReference.Where(char.IsLetter)]);
      string rowNumbers = new ([.. cellReference.Where(char.IsDigit)]);

      // Convert column letters to column index (A=1, B=2, ..., Z=26, AA=27, etc.)
      int columnIndex = 0;
      foreach (char letter in columnLetters)
      {
        columnIndex = (columnIndex * 26) + letter - 'A' + 1;
      }

      // Convert row numbers to row index
      int rowIndex = int.Parse(rowNumbers);

      return (columnIndex, rowIndex);
    }

    private static string GetCellValue(WorkbookPart workbookPart, DocumentFormat.OpenXml.Spreadsheet.Cell cell)
    {
      string value = cell.InnerText;
      if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
      {
        var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        if (stringTable != null)
        {
          value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        }
      }
      return value;
    }

    private void CreateXMLMirror(string directory)
    {
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
      string outputPath = Path.Combine(directory, fileNameWithoutExtension + ".xml");

      ToXMLDocument().Save(outputPath);
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CreateXMLMirror", $"XML file created at: {outputPath}");

      ToXMLDocument().Load(outputPath);
    }

    public XmlDocument ToXMLDocument()
    {
      XElement docElement = new("doc");
      docElement.Add(new XAttribute("fileName", FileName));

      foreach (string key in BasicData.Keys)
        docElement.Add(new XAttribute(key, BasicData.TryGetValue(key, out var keyValue) ? keyValue : ""));

      docElement.Add(new XAttribute("date", DateTime.Now.ToString("MM/dd/yy_HH:mm")));

      foreach (string varName in VarList.Keys.ToList())
      {
        XElement varElement = new("variable", varName);
        varElement.Add(new XAttribute("source", VarList[varName][0]));
        varElement.Add(new XAttribute("inOut", VarList[varName][1]));
        varElement.Add(new XAttribute("call", VarList[varName][2]));
        varElement.Add(new XAttribute("type", VarList[varName][3])
        {

        });

        docElement.Add(varElement);
      }

      // Create an XDocument with declaration
      XDocument xDoc = new (new ("1.0", "utf-8", "yes"), docElement);

      // Convert XDocument to XmlDocument
      XmlDocument xmlDoc = new();
      using (var reader = xDoc.CreateReader())
      {
        xmlDoc.Load(reader);
      }

      return xmlDoc;
    }

    public int GetVarCallLevel(string varName)
    {
      if (VarList.ContainsKey(varName)) return int.Parse(_varList[varName][2]);
      return 0;
    }
  }
}
