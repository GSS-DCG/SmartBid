using System.Data;
using System.Xml;
using ExcelDataReader;
using File = System.IO.File;

namespace SmartBid
{
  public record ToolData(string Resource, string Code, int Call, string Name, string FileType, string Version, string Lenguage, string Description, string FileName)
  {
    // Constructor to initialize all properties
    public ToolData(string resource, string code, int call, string name, string version, string language, string filetype, string description) : this(resource, code, call, name, filetype, version, language, description, $"{name}.{filetype}")
    {
    }

    // Methods
    public XmlDocument ToXMLDocument()
    {
      XmlDocument doc = new();
      _ = doc.AppendChild(ToXML(doc));
      return doc;
    }

    public XmlElement ToXML(XmlDocument mainDoc)
    {
      XmlElement toolElement = mainDoc.CreateElement("tool");
      toolElement.SetAttribute("resource", Resource);
      toolElement.SetAttribute("code", Code);
      toolElement.SetAttribute("call", Call.ToString());
      toolElement.SetAttribute("name", Name);
      toolElement.SetAttribute("fileType", FileType);
      toolElement.SetAttribute("version", Version);
      toolElement.SetAttribute("language", Lenguage);
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
    private static readonly object _lock = new();

    // Class variables
    public List<ToolData> Tools { get; private set; } = new();
    public List<string> DeliveryDocsPack { get; private set; } = new();

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
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ToolsMap", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
        _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", vmFile);
      }
      string directoryPath = Path.GetDirectoryName(vmFile)!;
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
      string xmlFile = Path.Combine(directoryPath, "ToolsMap" + ".xml");
      DateTime fileModified = File.GetLastWriteTime(vmFile);
      DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;
      if (fileModified > xmlModified)
      {
        LoadFromXLS(vmFile);
        SaveToXml(xmlFile);
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ToolsMap", $"XML file created at: {xmlFile}");
      }
      else
      {
        LoadFromXml(xmlFile);
      }
    }

    // Methods
    private void LoadFromXml(string xmlPath)
    {
      XmlDocument doc = new();
      doc.Load(xmlPath);
      foreach (XmlNode node in doc.SelectNodes("//tools/tool")!)
      {
        ToolData data = new(
            node.Attributes["resource"]!.InnerText ?? string.Empty,
            node.Attributes["code"]!.InnerText ?? string.Empty,
            int.TryParse(node.Attributes["call"]!.InnerText, out int callValue) ? callValue : 1, // Call
            node.Attributes["name"]!.InnerText ?? string.Empty,
            node.Attributes["version"]!.InnerText ?? string.Empty,
            node.Attributes["language"]!.InnerText ?? string.Empty,
            node.Attributes["fileType"]!.InnerText ?? string.Empty,
            node.Attributes["description"]!.InnerText ?? string.Empty
        );

        Tools.Add(data);
      }

      foreach (XmlNode node in doc.SelectNodes("//deliveryDocsPack/deliveryDocs")!)
      {
        DeliveryDocsPack.Add(node.InnerText);
      }


    }

    public void SaveToXml(string xmlPath, List<string>? varList = null)
    {
      XmlDocument doc = ToXml(varList);
      doc.Save(xmlPath);
    }

    public XmlDocument ToXml(List<string>? varList = null)
    {
      XmlDocument doc = new();
      XmlElement root = doc.CreateElement("root");
      doc.AppendChild(root);


      XmlElement tools= doc.CreateElement("tools");
      root.AppendChild(tools);
      foreach (var tool in Tools)
      {
        tools.AppendChild(tool.ToXML(doc));
      }

      XmlElement deliveryDocs = doc.CreateElement("deliveryDocsPack");
      root.AppendChild(deliveryDocs);
      foreach (string template in DeliveryDocsPack)
      {
        deliveryDocs.AppendChild(H.CreateElement(doc, "deliveryDocs", template));
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

      System.Data.DataTable dataTable = dataSet.Tables["ToolMap"]!;

      // Iterate over the rows, stopping when column A is empty

      for (int i = 0; i < dataTable.Rows.Count; i++)
      {
        DataRow row = dataTable.Rows[i];

        if (row.IsNull("RESOURCE"))
          break;

        ToolData data = new(
            row["RESOURCE"]?.ToString() ?? string.Empty,
            row["CODE"]?.ToString() ?? string.Empty,
            int.TryParse(row["CALL"]?.ToString(), out int callValue) ? callValue : 1,
            row["name"]?.ToString() ?? string.Empty,
            int.TryParse(row["VERSION"]?.ToString(), out int value) ? value.ToString("D3") : "000",
            row["LANGUAGE"]?.ToString() ?? string.Empty,
            row["FILE TYPE"]?.ToString() ?? string.Empty,
            row["DESCRIPTION"]?.ToString() ?? string.Empty
        );

        Tools.Add(data);
      }

      //Delivery Docs Pack
      dataTable = dataSet.Tables["DeliveryDocsPack"]!;

      for (int i = 0; i < dataTable.Rows.Count; i++)
      {
        DataRow row = dataTable.Rows[i];

        if (row.IsNull("Full Pack List"))
          break;

        DeliveryDocsPack.Add(row["Full Pack List"]?.ToString()!);
      }
    }

    public ToolData getToolDataByCode(string code, string language = "", int version = -1)
    {
      // returns the only ToolData that matches the code, language and version, first search all the tools that match code and language and out of all of them selects the one with the version, or with the highest version in case version is == -1.

      if (language == "")
      {
        language = H.GetSProperty("defaultLanguage");
      }
      var filteredTools = Tools.Where(tool => tool.Code.Equals(code, StringComparison.OrdinalIgnoreCase) && tool.Lenguage.Equals(language, StringComparison.OrdinalIgnoreCase)).ToList();

      if (filteredTools.Count == 0) //try default language
      {
        filteredTools = Tools.Where(tool => tool.Code.Equals(code, StringComparison.OrdinalIgnoreCase) && tool.Lenguage.Equals(H.GetSProperty("defaultLanguage"), StringComparison.OrdinalIgnoreCase)).ToList();

        H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, $"❌❌Error❌❌ -- No tool found with code '{code}' and language '{language}, used defaultLanguage ({H.GetSProperty("defaultLanguage")})instead'.");
      }

      if (filteredTools.Count == 0) //try default language
      {
        throw new KeyNotFoundException($"No tool found with code '{code}', language '{language}'.");
      }

      if (version == -1)
      {
        return filteredTools.OrderByDescending(tool => int.TryParse(tool.Version, out int ver) ? ver : 0).First();
      }
      else
      {
        var tool = filteredTools.FirstOrDefault(tool => int.TryParse(tool.Version, out int ver) && ver == version);
        if (tool == null)
        {
          throw new KeyNotFoundException($"No tool found with code '{code}', language '{language}', and version '{version}'.");
        }
        return tool;
      }
    }
    
    public List<ToolData> getToolsByResource(string resource)
    {
      return Tools.Where(tool => tool.Resource.Equals(resource, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    public XmlDocument CalculateExcel(ToolData tool, DataMaster dm)
    {
      // 2. Check if the file type is Excel or otherwise
      if (!(tool.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) || tool.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase)))
        throw new InvalidOperationException("The file is not an Excel type (.xlsx or .xlsm)");

      // 3. Retrieve theMirrorXML instance of the template
      MirrorXML mirror = new(tool);
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
      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"Calculating tool {tool.Code}");

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
                if (tableData.Name == "t" && tableData.HasChildNodes)
                {
                  workbook.WriteTable(rangeName, tableData);
                }
                else
                {
                  H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"No table data found for variable '{variableID}'.");
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
        XmlDocument results = new();
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

              _ = varElement.AppendChild(H.CreateElement(results, "origin", $"{tool.Code}+{timestamp}"));

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
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{newFilePath}'. Default value used");
              }


            }
            catch (Exception ex)
            {
              H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"❌Error❌ reading range '{rangeName}': {ex.Message}");
            }

            _ = varNode.AppendChild(varElement);
          }
        }
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"Values returned from {tool.Code} calculation");
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

    public void GenerateOuput(ToolData template, DataMaster dm)
    {

      // 3. Retrieves the MirrorXML instance
      MirrorXML mirror = new(template);

      // 4. Get the variables list from the mirror
      var variableMap = mirror.VarList;

      // 5. Build the full path to the file
      string templatePath = Path.Combine(template.Resource == "TEMPLATE" ? H.GetSProperty("TemplatesPath") : H.GetSProperty("ToolsPath"), template.FileName);

      // 6. Crear copia del archivo para trabajar
      string filePath = Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "OUTPUT", template.FileName);

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
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "** Error - GenerateOutputWord", $"Cannot copy template {filePath}");
        throw;
      }

      // Confeccionamos la lista de marcas a reemplazar
      Dictionary<string, VariableData> varList = [];

      List<string> varNotFound = [];

      mirror.VarList.Keys.ToList().ForEach(var =>
      {
        try { varList[var] = dm.GetVariableData(var); }
        catch (KeyNotFoundException) { varNotFound.Add(var); }
      });

      // If there are variables not found in the DataMaster stop process and print out missing variables
      if (varNotFound.Count > 0) throw new Exception("keysNotFoundInDataMaster: " + string.Join(", ", varNotFound));


      if (template.FileType.Equals("docx", StringComparison.OrdinalIgnoreCase))
      {
        SB_Word? doc = null;

        // 8. Open the document using SB_Word class
        try
        {
          H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOutputWord", $"Generating output - Open file: {template.Code}");
          doc = new SB_Word(filePath);

          List<string> removeBkm = mirror.VarList
              .Where(kvp => kvp.Value.Length > 3 && kvp.Value[3] == "bookmark")
              .Select(kvp => kvp.Key)
              .ToList();

          _ = removeBkm.RemoveAll(key => (varList.ContainsKey(key) && bool.TryParse(varList[key].Value.Trim(), out var result) && result) || varList[key].Value.Trim() == "1");

          doc.DeleteBookmarks(removeBkm);

          doc.ReplaceFieldMarks(varList);

          doc.Save();

          H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOuputWord", "Reemplazo realizado con éxito.");

          if (H.GetBProperty("generatePDF"))
          {
            _ = doc.SaveAsPDF();
          }
        }
        catch (Exception ex)
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOuputWord", $"❌❌ Error ❌❌ : " + ex.Message);
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "", "       " + ex.StackTrace);
        }
        finally
        {
          // 🔒 Ensure Word document and app close cleanly
          if (doc != null)
          {
            doc.Close(); // Close the document
            doc.ReleaseComObjectSafely(); 
          }

          // 🧹 Clean up unmanaged resources
          GC.Collect();
          GC.WaitForPendingFinalizers();
        }
        H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOutputWord", $"Generated output: {template.Code} finished\n\n");

      }
      else if (template.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) ||
                 template.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase))
      {

        SB_Excel? doc = null;

        // 8. Open the Word document using SB_Word class
        try
        {
          H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOutputWord", $"Generating output - Open file: {template.Code}");
          doc = new SB_Excel(filePath);

          foreach (VariableData entry in varList.Values)
          {
            string variableID = entry.ID;
            string rangeName = $"{H.GetSProperty("VarPrefix")}{variableID}";
            string type = _variablesMap.GetVariableData(variableID).Type;

            if (type == "table")
            {
              XmlNode tableData = dm.GetValueXmlNode(variableID);
              if (tableData.Name == "t" && tableData.HasChildNodes)
              {
                doc.WriteTable(rangeName, tableData);
              }
              else
              {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CalculateExcel", $"No table data found for variable '{variableID}'.");
              }
            }
            else
            {
              _ = doc.FillUpValue(rangeName, dm.GetValueString(variableID));
            }
          }

          doc.Save();

          H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOuputWord", "Reemplazo realizado con éxito.");

          //if (H.GetBProperty("generatePDF"))
          //{
          //  _ = doc.SaveAsPDF();
          //}
        }
        catch (Exception ex)
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOuputWord", $"❌❌ Error ❌❌ : " + ex.Message);
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "", "       " + ex.StackTrace);
        }
        finally
        {
          // 🔒 Ensure Word document and app close cleanly
          if (doc != null)
          {
            doc.Close(); // Close the document
            doc.ReleaseComObjectSafely();
          }

          // 🧹 Clean up unmanaged resources
          GC.Collect();
          GC.WaitForPendingFinalizers();
        }
        H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "GenerateOutputWord", $"Generated output: {template.Code} finished\n\n");

      }
      else
      {
        throw new InvalidOperationException($"Unsupported file type: {template.FileType}");

      }
    }


  }// End of class ToolsMap
}// End of namespace SmartBid
