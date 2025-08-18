using System.Data;
using System.Xml;
using ExcelDataReader;
using File = System.IO.File;

namespace SmartBid
{
  public record ToolData(string Resource, string ID, int Call, string Name, string FileType, string Version, string Description, string FileName)
  {
    public static Dictionary<string, string> OpcionesHerramientas = new Dictionary<string, string>();

    // Constructor to initialize all properties
    public ToolData(string resource, string id, int call, string name, string version, string filetype, string description) : this(resource, id, call, name, filetype, version, description, $"{name}_{version}.{filetype}")
    {
    }

    // Methods
    public XmlDocument ToXMLDocument()
    {
      XmlDocument doc = new XmlDocument();
      _ = doc.AppendChild(ToXML(doc));
      return doc;
    }

    public XmlElement ToXML(XmlDocument mainDoc)
    {
      XmlElement toolElement = mainDoc.CreateElement("tool");
      toolElement.SetAttribute("resource", Resource);
      toolElement.SetAttribute("code", ID);
      toolElement.SetAttribute("call", Call.ToString());
      toolElement.SetAttribute("name", Name);
      toolElement.SetAttribute("fileType", FileType);
      toolElement.SetAttribute("version", Version);
      toolElement.SetAttribute("description", Description);
      toolElement.SetAttribute("fileName", FileName);

      return toolElement;
    }

  }

  public class ToolsMap
  {
    // Private static instance
    private static ToolsMap? _instance;

    // Lock object for thread safety
    private static readonly object _lock = new object();

    // Class variables
    public List<ToolData> Tools { get; private set; } = new List<ToolData>();
    private VariablesMap _variablesMap = VariablesMap.Instance; // Instance of VariablesMap

    // Public static property to get the single instance
    public static ToolsMap Instance
    {
      get
      {
        lock (_lock)
        {
          if (_instance == null)
          {
            _instance = new ToolsMap();
          }
          return _instance;
        }
      }
    }

    // Private constructor to prevent instantiation from outside
    private ToolsMap()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
      string vmFile = Path.GetFullPath(H.GetSProperty("VarMap"));
      if (!File.Exists(vmFile))
      {
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ToolsMap", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
        _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", vmFile);
      }
      string directoryPath = Path.GetDirectoryName(vmFile);
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
      string xmlFile = Path.Combine(directoryPath, "ToolsMap" + ".xml");
      DateTime fileModified = File.GetLastWriteTime(vmFile);
      DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;
      if (fileModified > xmlModified)
      {
        LoadFromXLS(vmFile);
        SaveToXml(xmlFile);
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ToolsMap", $"XML file created at: {xmlFile}");
      }
      else
      {
        LoadFromXml(xmlFile);
      }
    }

    // Methods
    private void LoadFromXml(string xmlPath)
    {
      XmlDocument doc = new XmlDocument();
      doc.Load(xmlPath);
      foreach (XmlNode node in doc.SelectNodes("//tool"))
      {
        ToolData data = new ToolData(
            node.Attributes["resource"]?.InnerText ?? string.Empty,
            node.Attributes["code"]?.InnerText ?? string.Empty,
            int.TryParse(node.Attributes["call"]?.InnerText, out int callValue) ? callValue : 1, // Call
            node.Attributes["name"]?.InnerText ?? string.Empty,
            node.Attributes["version"]?.InnerText ?? string.Empty,
            node.Attributes["fileType"]?.InnerText ?? string.Empty,
            node.Attributes["description"]?.InnerText ?? string.Empty
        );

        Tools.Add(data);
      }
    }

    public void SaveToXml(string xmlPath, List<string>? varList = null)
    {
      XmlDocument doc = ToXml(varList);
      doc.Save(xmlPath);
    }

    public XmlDocument ToXml(List<string>? varList = null)
    {
      XmlDocument doc = new XmlDocument();
      XmlElement root = doc.CreateElement("root");
      _ = doc.AppendChild(root);
      foreach (var tool in Tools)
      {
        _ = root.AppendChild(tool.ToXML(doc));
      }
      return doc;
    }

    private void LoadFromXLS(string vmFile)
    {
      DataSet dataSet;
      using (var stream = File.Open(vmFile, FileMode.Open, FileAccess.Read))
      {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
          using var _ = dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
          {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration()
            {
              UseHeaderRow = true
            }
          });
        }
      }

      System.Data.DataTable dataTable = dataSet.Tables["ToolMap"];

      // Iterate over the rows, stopping when column A is empty
      for (int i = 1; i < dataTable.Rows.Count; i++)
      {
        DataRow row = dataTable.Rows[i];

        if (row.IsNull(0))
          break;

        ToolData data = new ToolData(
            row[0]?.ToString() ?? string.Empty,  // resource 
            row[1]?.ToString() ?? string.Empty,  // code     
            int.TryParse(row[2]?.ToString(), out int callValue) ? callValue : 1, //Call
            row[3]?.ToString() ?? string.Empty,  // name     
            int.TryParse(row[4]?.ToString(), out int value) ? value.ToString("D3") : "000",  // version  
            row[5]?.ToString() ?? string.Empty,  // filetype 
            row[6]?.ToString() ?? string.Empty  // description
        );

        Tools.Add(data);
      }
    }

    public ToolData getToolDataByCode(string code)
    {
      return Tools.FirstOrDefault(tool => tool.ID == code);
    }

    public XmlDocument CalculateExcel(string toolID, DataMaster dm)
    {
      // 1. Retrieve the tool info
      ToolData tool = ToolsMap.Instance.getToolDataByCode(toolID);
      if (tool == null)
        throw new ArgumentException($"ToolID '{toolID}' not found.");

      // 2. Check if the file type is Excel or otherwise
      if (!(tool.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) || tool.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase)))
        throw new InvalidOperationException("The file is not an Excel type (.xlsx or .xlsm)");

      // 3. Retrieve theMirrorXML instance of the tool
      MirrorXML mirror = new MirrorXML(toolID);
      var variableMap = mirror.VarList;

      // 4. Find the original file
      string filePath = Path.Combine(
          tool.Resource == "TOOL" ? H.GetSProperty("ToolsPath") : H.GetSProperty("TemplatesPath"),
          tool.FileName
      );

      // 5. Create a copy of the file to work with
      string newFilePath = Path.Combine(
          H.GetSProperty("processPath"),
          dm.GetInnerText(@"dm/utils/utilsData/opportunityFolder"),
          "TOOLS",
          tool.FileName
      );

      // Make sure the directory exists and copying the file
      _ = Directory.CreateDirectory(Path.GetDirectoryName(newFilePath)!);
      File.Copy(filePath, newFilePath, true);

      // 6. Abrir el archivo en Excel Interop
      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"Calculating tool {toolID}");

      SB_Excel? workbook = null;

      try
      {
        workbook = new SB_Excel(newFilePath);

        // Escribir valores en las celdas
        foreach (var entry in variableMap)
        {
          if (mirror.GetVarCallLevel(entry.Key) == tool.Call)
          {
            string variableID = entry.Key;
            string direction = entry.Value[1];

            if (direction == "in")
            {
              string rangeName = $"{H.GetSProperty("IN_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";
              string type = _variablesMap.GetVariableData(variableID).Type;

              if (type == "table")
              {
                XmlNode tableData = dm.GetValueXmlNode(variableID);
                if (tableData.SelectSingleNode("t") != null && tableData.SelectSingleNode("t").HasChildNodes)
                {
                  workbook.WriteTable(rangeName, tableData);
                }
                else
                {
                  H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"No table data found for variable '{variableID}'.");
                }
              }
              else
              {
                _ = workbook.FillUpValue(rangeName, dm.GetValueString(variableID));
              }
            }
          }
        }


        workbook.Calculate();

        // Read output variables
        XmlDocument results = new XmlDocument();
        XmlElement root = results.CreateElement("root");
        _ = results.AppendChild(root);
        XmlElement varNode = results.CreateElement("variables");
        _ = root.AppendChild(varNode);

        // Obtener la fecha y hora actual en formato dd-HH:mm
        string timestamp = DateTime.Now.ToString("dd-HH:mm");

        foreach (var entry in variableMap)
        {
          string variableID = entry.Key;
          string direction = entry.Value[1];

          // find all out variables from the actual level
          if (mirror.GetVarCallLevel(variableID) == tool.Call && direction == "out")
          {
            XmlElement varElement = results.CreateElement(variableID);
            string rangeName = $"{H.GetSProperty("OUT_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";

            try
            {
              string type = _variablesMap.GetVariableData(variableID).Type;
              bool setValue = false;

              XmlElement value = results.CreateElement("value");
              value.SetAttribute("type", type);
              _ = varElement.AppendChild(value);

              XmlElement note = results.CreateElement("note");
              _ = varElement.AppendChild(note);

              _ = varElement.AppendChild(H.CreateElement(results, "origin", $"{toolID}+{timestamp}"));

              if (type != "table")
              {
                value.InnerText = workbook.GetSValue(rangeName);
                note.InnerText = "Calculated Value";
                setValue = true;
              }
              else
              {
                XmlNode tableXmlValue = results.ImportNode(workbook.GetTValue(rangeName), true);

                if (tableXmlValue != null && tableXmlValue.HasChildNodes)
                {
                  _ = value.AppendChild(tableXmlValue);
                  note.InnerText = "Calculated Value";
                  setValue = true;
                }
                else
                {
                  //por el momento no tenemos valores por defecto para tabla
                  value.InnerText = "No data found in table.";
                }
              }
              if (!setValue)
              {
                value.InnerText = _variablesMap.GetVariableData(variableID).Default ?? string.Empty;
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{newFilePath}'. Default value used");
              }


            }
            catch (Exception ex)
            {
              H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"❌Error❌ reading range '{rangeName}': {ex.Message}");
            }

            _ = varNode.AppendChild(varElement);
          }
        }
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"Values returned from {toolID} calculation");
        H.PrintXML(2, results);
        workbook.Close(); // No need to save changes, 
        return results;
      }
      finally
      {
        workbook.Close(); // Close the workbook
        workbook.ReleaseComObjectSafely();

        // 🧹 Clean up unmanaged resources
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }

    }

    public void GenerateOuputWord(string templateID, DataMaster dm)
    {
      // 1. Retrieve the tool data by ID
      ToolData tool = ToolsMap.Instance.getToolDataByCode(templateID);
      if (tool == null)
        throw new ArgumentException($"ToolID '{templateID}' not found.");

      // 2. Check if the file type is Word
      if (!tool.FileType.Equals("docx", StringComparison.OrdinalIgnoreCase))
        throw new InvalidOperationException("The file is not a Word file (.docx)");

      // 3. Create a MirrorXML instance
      MirrorXML mirror = new MirrorXML(templateID);

      // 4. Get the variables list from the mirror
      var variableMap = mirror.VarList;

      // 5. Build the full path to the file
      string templatePath = Path.Combine(tool.Resource == "TEMPLATE" ? H.GetSProperty("TemplatesPath") : H.GetSProperty("ToolsPath"), tool.FileName);

      // 6. Crear copia del archivo para trabajar
      string filePath = Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "OUTPUT", tool.FileName);

      // 7. Ensure the output directory exists and copy the template file if it doesn't exist
      if (!File.Exists(filePath))
      {
        if (!Directory.Exists(Path.GetDirectoryName(filePath)))
          _ = Directory.CreateDirectory(Path.GetDirectoryName(filePath));
      }

      try
      {
        File.Copy(templatePath, filePath, true);
      }
      catch (Exception)
      {
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "** Error - GenerateOutputWord", $"Cannot copy template {filePath}");
        throw;
      }

      // Confeccionamos la lista de marcas a reemplazar
      Dictionary<string, VariableData> varList = new Dictionary<string, VariableData>();

      List<string> varNotFound = new List<string>();

      mirror.VarList.Keys.ToList().ForEach(var =>
      {
        try { varList[var] = dm.GetVariableData(var); }
        catch (KeyNotFoundException) { varNotFound.Add(var); }
      });

      // If there are variables not found in the DataMaster stop process and print out missing variables
      if (varNotFound.Count > 0) throw new Exception("keysNotFoundInDataMaster: " + string.Join(", ", varNotFound));


      SB_Word? doc = null;

      // 8. Open the Word document using SB_Word class
      try
      {
        H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOutputWord", $"Generating output - Open file: {templateID}");
        doc = new SB_Word(filePath);

        List<string> removeBkm = mirror.VarList
            .Where(kvp => kvp.Value.Length > 3 && kvp.Value[3] == "bookmark")
            .Select(kvp => kvp.Key)
            .ToList();

        removeBkm.RemoveAll(key => varList.ContainsKey(key) && bool.TryParse(varList[key].Value.Trim(), out var result) && result || varList[key].Value.Trim() == "1");



        doc.DeleteBookmarks(removeBkm);

        doc.ReplaceFieldMarks(varList);

        doc.Save();

        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOuputWord", "Reemplazo realizado con éxito.");

        if (H.GetBProperty("generatePDF"))
        {
          _ = doc.SaveAsPDF();
        }
      }
      catch (Exception ex)
      {
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOuputWord", $"❌❌ Error ❌❌ : " + ex.Message);
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "", "       " + ex.StackTrace);
      }
      finally
      {
        // 🔒 Ensure Word document and app close cleanly
        doc.Close(); // Close the document
        doc.ReleaseComObjectSafely();

        // 🧹 Clean up unmanaged resources
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOutputWord", $"Generated output: {templateID} finished\n\n");
    }

  }// End of class ToolsMap
}// End of namespace SmartBid
