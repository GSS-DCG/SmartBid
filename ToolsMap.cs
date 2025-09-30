using System.Data;
using System.Diagnostics;
using System.Text;
using System.Xml;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Mysqlx.Prepare;
using OfficeOpenXml.Utils;
using Exception = System.Exception;
using File = System.IO.File;

namespace SmartBid
{
  public class ToolData
  {
    public string Resource { get; }
    public string Code { get; }
    public int Call { get; }
    public string Name { get; }
    public string FileType { get; }
    public string Interpreter { get; }
    public string Version { get; }
    public string Lenguage { get; }
    public string Description { get; }
    public string FileName { get; }

    public ToolData(string resource, string code, int call, string name, string filetype, string interpreter,string version, string language, string description)
    {
      Resource = resource;
      Code = code;
      Call = call;
      Name = name;
      FileType = filetype.ToLowerInvariant(); // aquí se normaliza
      Interpreter = interpreter;
      Version = version;
      Lenguage = language;
      Description = description;
      FileName = $"{name}.{FileType}"; // usa el valor ya normalizado
    }

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
      toolElement.SetAttribute("interpreter", Interpreter);
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
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ToolsMap", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
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
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ToolsMap", $"XML file created at: {xmlFile}");
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
            node.Attributes["fileType"]!.InnerText ?? string.Empty,
            node.Attributes["interpreter"]!.InnerText ?? string.Empty,
            node.Attributes["version"]!.InnerText ?? string.Empty,
            node.Attributes["language"]!.InnerText ?? string.Empty,
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
            row["FILE TYPE"]?.ToString() ?? string.Empty,
            row["INTERPRETER"]?.ToString() ?? string.Empty,
            int.TryParse(row["VERSION"]?.ToString(), out int value) ? value.ToString("D3") : "000",
            row["LANGUAGE"]?.ToString() ?? string.Empty,
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

        H.PrintLog(4, 
          TC.ID.Value!.Time(), 
          TC.ID.Value!.User, 
          "getToolDataByCode",
					$"❌❌Error❌❌ -- No tool found with code '{code}' and language '{language}, used defaultLanguage ({H.GetSProperty("defaultLanguage")})instead'.");
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

    public XmlDocument Calculate(ToolData tool, DataMaster dm)
    {
      if (tool.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) ||
          tool.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase))
      {
        return CalculateExcel(tool, dm);
      }
      else if (tool.FileType.Equals("exe", StringComparison.OrdinalIgnoreCase))
      {
        return CalculateExe(tool, dm);
      }
      else
      {
        throw new InvalidOperationException($"The file: {tool.FileName}.{tool.FileType} is not a supported type (.xlsx, .xlsm, or .exe)");
      }
    }

    private XmlDocument CalculateExcel(ToolData tool, DataMaster dm)
    {
      MirrorXML mirror = new(tool);
      var variableMap = mirror.VarList;

      string originalToolPath = Path.Combine(
          tool.Resource == "TOOL" ? H.GetSProperty("ToolsPath") : H.GetSProperty("TemplatesPath"),
          tool.FileName
      );

      string toolPath = Path.Combine(
          H.GetSProperty("processPath"),
          dm.GetInnerText($@"dm/utils/utilsData/opportunityFolder"),
          dm.SBidRevision,
          "TOOLS",
          tool.FileName
      );

      _ = Directory.CreateDirectory(Path.GetDirectoryName(toolPath)!);
      File.Copy(originalToolPath, toolPath, true);

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"  Calling tool {tool.Code}");

      SB_Excel? workbook = null;
      XmlDocument results = new(); // Declarado fuera del try para asegurar retorno

      try
      {
        workbook = new SB_Excel(toolPath);

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
                  H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"No table data found for variable '{variableID}'.");
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

        XmlElement answer = results.CreateElement("answer");
        _ = results.AppendChild(answer);
        XmlElement varNode = results.CreateElement("variables");
        _ = answer.AppendChild(varNode);

        string timestamp = DateTime.Now.ToString("dd-HH:mm");

        foreach (var entry in variableMap)
        {
          string variableID = entry.Key;
          string direction = entry.Value[1];

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
                  value.InnerText = "No data found in table.";
                }
              }

              if (!setValue)
              {
                value.InnerText = _variablesMap.GetVariableData(variableID).Default ?? string.Empty;
                H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{toolPath}'. Default value used");
              }
            }
            catch (Exception ex)
            {
              H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"❌Error❌ reading range '{rangeName}': {ex.Message}");
            }

            _ = varNode.AppendChild(varElement);
          }
        }

        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"Values returned from {tool.Code} calculation", results);
      }
      finally
      {
        workbook?.Close();
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }

      // Adding informtion for Utils
      // Check whether <utils> exists in results and append the <utils> node if it doesn't
      XmlElement utilsNode = (XmlElement)results.SelectSingleNode("//root/utils");
      if (utilsNode == null)
      {
        utilsNode = results.CreateElement("utils");
        results.DocumentElement.AppendChild(utilsNode);
      }

      // Check whether <tools> exists in <utils>  and append the <tools> node if it doesn't
      XmlElement toolsNode = (XmlElement)utilsNode.SelectSingleNode("tools");
      if (toolsNode == null)
      {
        toolsNode = results.CreateElement("tools");
        utilsNode.AppendChild(toolsNode);
      }

      XmlElement toolNode = tool.ToXML(results);
      string newNodeName = toolNode.GetAttribute("code");

      XmlElement newToolNode = results.CreateElement(newNodeName);
      foreach (XmlAttribute attr in toolNode.Attributes)
      {
        if (attr.Name != "code")
        {
          if (attr.Name != "fileName")
            newToolNode.SetAttribute(attr.Name, attr.Value);
          else
            newToolNode.SetAttribute(attr.Name, toolPath);
        }
      }
      toolsNode.AppendChild(newToolNode);




      return results;
    }

    private XmlDocument CalculateExe(ToolData tool, DataMaster dm)
    {
      //Executes type .exe tools, first creates and xml to send as input call and then reads the output xml
      //from the stdout of the process. the exact location of the file is read from the mirrorXML of the tool. 
      //originalToolPath is coming as an argument from the answer of the tool

      MirrorXML mirror = new(tool);
      var variableList = mirror.VarList;

      string? filePath = mirror.FileName;
      string? arguments = "";


      if (File.Exists(filePath)) {
        if (Path.GetExtension(filePath).Equals(".py", StringComparison.OrdinalIgnoreCase))
        {
          arguments = $"\"{filePath}\" 00"; // any argument should be sent to the call to indicate that the input is coming from stdin

          if (tool.Interpreter != null && tool.Interpreter != "")
            filePath = tool.Interpreter;
          else
            filePath = @"python.exe";
        }
      }
      else
      {
        throw new InvalidOperationException($"The tool {tool.Code} does not have a valid file path in the mirror.\n" +
          $"check that Tool mirror xml file has the argument 'fileName' pointing to the correct executable file of the tool" +
          $"file {filePath} Not Found");
      }

      //Generate the CallXML to send as input to the tool
      XmlDocument callXml = new();
      XmlDeclaration xmlDeclaration = callXml.CreateXmlDeclaration("1.0", "UTF-8", null);
      _ = callXml.AppendChild(xmlDeclaration);

      XmlElement callNode = callXml.CreateElement("call");
      callXml.AppendChild(callNode);

      XmlElement variablesIn = callXml.CreateElement("variables");
      callNode.AppendChild(variablesIn);
      XmlElement variablesOut= callXml.CreateElement("out");
      callNode.AppendChild(variablesOut);

      foreach (var entry in variableList)
      {
        string variableID = entry.Key;
        string direction = entry.Value[1];

        if (direction == "in" && mirror.GetVarCallLevel(variableID) == tool.Call)
        {
          XmlElement varElement = callXml.CreateElement(variableID);
          varElement.SetAttribute("unit", dm.GetValueUnit(variableID));
          varElement.InnerText = dm.GetValueString(variableID);
          variablesIn.AppendChild(varElement);
        }
        else if (direction == "out" && mirror.GetVarCallLevel(variableID) == tool.Call)
        {
          XmlElement varElement = _variablesMap.GetVariableData(variableID).ToXML(callXml);
          variablesOut.AppendChild(varElement);
        }
      }

      string xmlVarList = callXml.OuterXml;

      H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"   Calling Tool: {Path.GetFileName(filePath)} {arguments}");
      H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"   Calling Tool: {filePath} {arguments}");
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", "   call message:", callXml);


      ProcessStartInfo psi = new()
      {
        FileName = filePath,
        Arguments = arguments,
        RedirectStandardInput = true,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true,
        StandardInputEncoding = Encoding.UTF8
      };

      string output;
      string error;

      using (Process process = new() { StartInfo = psi })
      {
        _ = process.Start();

        using (StreamWriter writer = process.StandardInput)
        {
          writer.Write(xmlVarList);
          writer.Flush();
          writer.Close();
        }

        output = process.StandardOutput.ReadToEnd();
        error = process.StandardError.ReadToEnd();
        process.WaitForExit();
      }

      XmlDocument results = new();
      results.LoadXml(output);

      H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", $"Return from tool", results);

      if (!string.IsNullOrWhiteSpace(error))
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", $"❌Error❌:\n{error}");

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", "-----------------------------------");

      // Adding information for Utils
      // Check whether <utils> exists in results and append the <utils> node if it doesn't
      XmlElement utilsNode = (XmlElement)results.SelectSingleNode("//answer/utils");
      if (utilsNode == null)
      {
        utilsNode = results.CreateElement("utils");
        results.DocumentElement.AppendChild(utilsNode);
      }

      // Check whether <tools> exists in <utils>  and append the <tools> node if it doesn't
      XmlElement toolsNode = (XmlElement)utilsNode.SelectSingleNode("tools");
      if (toolsNode == null)
      {
        toolsNode = results.CreateElement("tools");
        utilsNode.AppendChild(toolsNode);
      }

      XmlElement toolNode = tool.ToXML(results);
      string newNodeName = toolNode.GetAttribute("code");

      XmlElement newToolNode = results.CreateElement(newNodeName);
      foreach (XmlAttribute attr in toolNode.Attributes)
      {
        if (attr.Name != "code")
        {
          if (attr.Name == "fileName")
            newToolNode.SetAttribute(attr.Name, filePath);
          else
            newToolNode.SetAttribute(attr.Name, attr.Value);
        }
      }
      toolsNode.AppendChild(newToolNode);


      // add default value to every expected 'out' variable that is not coming in the results by reading all expected out variables
      // from variableList direction = "out" and checking if they are in the results xml
      // add all missing out variables using the default value from VariablesMap
      // first get the variables node from results
      XmlElement answerNode = (XmlElement)results.SelectSingleNode("//answer");
      XmlElement resultsVariablesNode = (XmlElement)answerNode.SelectSingleNode("variables");
      List<string> resultVarIDs = new();
      foreach (XmlElement varNode in resultsVariablesNode.ChildNodes)
      {
        resultVarIDs.Add(varNode.Name);
      }
      foreach (var entry in variableList)
      {
        string variableID = entry.Key;
        string direction = entry.Value[1];
        if (direction == "out" && mirror.GetVarCallLevel(variableID) == tool.Call && !resultVarIDs.Contains(variableID))
        {
          XmlElement varElement = CreateXmlVariable(results, variableID, _variablesMap.GetVariableData(variableID).Type, _variablesMap.GetVariableData(variableID).Default ?? string.Empty, $"{tool.Code}+{DateTime.Now:dd-HH:mm}");
          resultsVariablesNode.AppendChild(varElement);
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"❌ Missing output variable '{variableID}' added with default value.");
        }
      }



      return results;
    }

    private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, string value, string origin)
    {
      XmlElement xmlVar = outputXML.CreateElement(ID);
      outputXML.SelectSingleNode("//answer/variables")?.AppendChild(xmlVar);

      xmlVar.AppendChild(CreateElementWithText(outputXML, "value", value, ("type", type)));
      xmlVar.AppendChild(CreateElementWithText(outputXML, "origin", $"{origin}-{DateTime.Now:yyMMdd-HHmm}"));
      xmlVar.AppendChild(CreateElementWithText(outputXML, "note", "Calculated Value"));

      return xmlVar;
    }
    private static XmlElement CreateElementWithText(XmlDocument doc, string elementName, string textContent, params (string, string)[] attributes)
    {
      XmlElement element = doc.CreateElement(elementName);
      element.InnerText = textContent;
      foreach (var (attrName, attrValue) in attributes)
      {
        element.SetAttribute(attrName, attrValue);
      }
      return element;
    }



    public void GenerateOuput(ToolData template, DataMaster dm)
    {

      // 3. Retrieves the MirrorXML instance
      MirrorXML mirror = new(template);

      // 4. Get the variables list from the mirror
      var variableMap = mirror.VarList;

      // 5. Build the full path to the file
      string templatePath = Path.Combine(
        template.Resource == "TEMPLATE" ? H.GetSProperty("TemplatesPath") : H.GetSProperty("ToolsPath"), 
        template.FileName
        );

      // 6. Crear copia del archivo para trabajar
      string filePath = Path.Combine(
        H.GetSProperty("processPath"), 
        dm.GetValueString("opportunityFolder"),
        dm.SBidRevision,
        "OUTPUT", 
        template.FileName
        );

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
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "** Error - GenerateOutputWord", $"Cannot copy template {filePath}");
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
          H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOutputWord", $"Generating output - Open file: {template.Code}");
          doc = new SB_Word(filePath);

          List<string> removeBkm = mirror.VarList
              .Where(kvp => kvp.Value.Length > 3 && kvp.Value[3] == "bookmark")
              .Select(kvp => kvp.Key)
              .ToList();

          _ = removeBkm.RemoveAll(key => (varList.ContainsKey(key) && bool.TryParse(varList[key].Value.Trim(), out var result) && result) || varList[key].Value.Trim() == "1");

          doc.DeleteBookmarks(removeBkm);

          doc.ReplaceFieldMarks(varList);

          doc.Save();

          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOuputWord", "Reemplazo realizado con éxito.");

          if (H.GetBProperty("generatePDF"))
          {
            _ = doc.SaveAsPDF();
          }
        }
        catch (Exception ex)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOuputWord", $"❌❌ Error ❌❌ : " + ex.Message);
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "", "       " + ex.StackTrace);
        }
        finally
        {
          // 🔒 Ensure Word document and app close cleanly
          if (doc != null)
          {
            doc.Close(); // Close the document
          }

          // 🧹 Clean up unmanaged resources
          GC.Collect();
          GC.WaitForPendingFinalizers();
        }
        H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOutputWord", $"Generated output: {template.Code} finished\n\n");

      }
      else if (template.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) ||
                 template.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase))
      {

        SB_Excel? doc = null;

        // 8. Open the Word document using SB_Word class
        try
        {
          H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOutputWord", $"Generating output - Open file: {template.Code}");
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
                H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"No table data found for variable '{variableID}'.");
              }
            }
            else
            {
              _ = doc.FillUpValue(rangeName, dm.GetValueString(variableID));
            }
          }

          doc.Save();

          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOuputWord", "Reemplazo realizado con éxito.");

          //if (H.GetBProperty("generatePDF"))
          //{
          //  _ = doc.SaveAsPDF();
          //}
        }
        catch (Exception ex)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOuputWord", $"❌❌ Error ❌❌ : " + ex.Message);
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "", "       " + ex.StackTrace);
        }
        finally
        {
          // 🔒 Ensure Word document and app close cleanly
          if (doc != null)
          {
            doc.Close(); // Close the document
          }

          // 🧹 Clean up unmanaged resources
          GC.Collect();
          GC.WaitForPendingFinalizers();
        }
        H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOutputWord", $"Generated output: {template.Code} finished\n\n");

      }
      else
      {
        throw new InvalidOperationException($"Unsupported file type: {template.FileType}");

      }
    }


  }// End of class ToolsMap
}// End of namespace SmartBid
