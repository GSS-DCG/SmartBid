using System.Data;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using ExcelDataReader;
using Exception = System.Exception;
using File = System.IO.File;

namespace SmartBid
{
  public class ToolData
  {

    private string _resource;
    public string Resource
    {
      get => _resource;
      set => _resource = value.ToUpper();
    }
    public string Code { get; }
    public int Call { get; }
    public string Name { get; }
    public string FileType { get; }
    public bool IsThreadSafe { get; } = true;
    public int TimeoutMinutes { get; } = 0;
    public bool IsIterative { get; } = false;
    public string Interpreter { get; }
    public string Version { get; }
    public string Language { get; }
    public string Description { get; }
    public string FileName { get; }
    public string InPrefix { get; }
    public string OutPrefix { get; }


    public ToolData(string resource, string code, int call, string name, string filetype, bool isThreadSafe, int timeoutMinutes, bool isIterative, string interpreter, string version, string language, string description)
    {
      Resource = resource;
      Code = code;
      Call = call;
      Name = name;
      FileType = filetype.ToLowerInvariant(); // aquí se normaliza
      IsThreadSafe = isThreadSafe;
      TimeoutMinutes = timeoutMinutes;
      IsIterative = isIterative;
      Interpreter = interpreter;
      Version = version;
      Language = language;
      Description = description;
      FileName = $"{name}.{FileType}"; // usa el valor ya normalizado

      string prefix = string.Empty;
      if (Resource == "TOOL")
        prefix = H.GetSProperty("IN_VarPrefix") + $"call{Call}_";
      else if (Resource == "TEMPLATE")
        prefix = H.GetSProperty("VarPrefix");
      else
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ToolData", $"Resource not defined for tool {Name}");
      }

      InPrefix = prefix;

      prefix = string.Empty;
      if (Resource == "TOOL")
        prefix = H.GetSProperty("OUT_VarPrefix") + $"call{Call}_";
      else if (Resource == "TEMPLATE")
        prefix = "";
      OutPrefix = prefix;

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
      toolElement.SetAttribute("isThreadSafe", IsThreadSafe.ToString());
      toolElement.SetAttribute("isIterative", IsIterative.ToString());
      toolElement.SetAttribute("timeoutMinutes", TimeoutMinutes.ToString());
      toolElement.SetAttribute("version", Version);
      toolElement.SetAttribute("language", Language);
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

    private readonly object _trafficLightLock = new object();
    private Dictionary<string, List<int>> _trafficLight = new Dictionary<string, List<int>>();

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

    public bool CheckForGreenLight(string toolID, int callID, out int order)
    {
      bool green = true;
      order = 0;
      List<int> callList;

      lock (_trafficLightLock)
      {
        if (!_trafficLight.ContainsKey(toolID))
        {
          _trafficLight[toolID] = new List<int> { callID };
          callList = _trafficLight[toolID];
          green = true;
        }
        else
        {
          callList = _trafficLight[toolID];

          if (callList.Count == 0)
          {
            callList.Add(callID);
            green = true;
          }
          else
          {
            if (callList[0] == callID)
            {
              green = true;
            }
            else
            {
              if (!callList.Contains(callID))
              {
                callList.Add(callID);
              }
              order = callList.IndexOf(callID) + 1;
              green = false;
            }
          }
        }
        H.PrintLog(1, TC.ID.Value?.Time() ?? "00:00.000", TC.ID.Value?.User ?? "SYSTEM", "CheckForGreenLight", $"  semaforo: {green} for {toolID}: order: {order}");
        H.PrintLog(1, TC.ID.Value?.Time() ?? "00:00.000", TC.ID.Value?.User ?? "SYSTEM", "CheckForGreenLight", $"  callList = {string.Join(",", callList)}");
      }
      return green;
    }
    public void ReleaseProcess(string toolID, int callID)
    {
      lock (_trafficLightLock)
      {
        if (_trafficLight.ContainsKey(toolID))
        {
          List<int> callList = _trafficLight[toolID];
          if (callList.Contains(callID))
          {
            _ = callList.Remove(callID);
            H.PrintLog(2, TC.ID.Value?.Time() ?? "00:00.000", TC.ID.Value?.User ?? "SYSTEM", "ReleaseProcess", $"  Proceso {callID} liberado para la herramienta {toolID}. Lista restante: {string.Join(",", callList)}");
          }
        }
      }
    }
    private void LoadFromXml(string xmlPath)
    {
      XmlDocument doc = new();
      doc.Load(xmlPath);
      foreach (XmlNode node in doc.SelectNodes("//tools/tool")!)
      {
        ToolData data = new(
            node.Attributes["resource"]!.InnerText ?? string.Empty,
            node.Attributes["code"]!.InnerText ?? string.Empty,
            int.TryParse(node.Attributes["call"]!.InnerText, out int callValue) ? callValue : 1,
            node.Attributes["name"]!.InnerText ?? string.Empty,
            node.Attributes["fileType"]!.InnerText ?? string.Empty,
            bool.TryParse(node.Attributes["isThreadSafe"]?.InnerText, out var val) ? val : true,
            int.TryParse(node.Attributes["timeoutMinutes"]?.InnerText, out int timeoutValue) ? timeoutValue : 0, // Timeout Minutes
            bool.TryParse(node.Attributes["isIterative"]?.InnerText, out var iterVal) ? iterVal : false,
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
      _ = doc.AppendChild(root);


      XmlElement tools = doc.CreateElement("tools");
      _ = root.AppendChild(tools);
      foreach (var tool in Tools)
      {
        _ = tools.AppendChild(tool.ToXML(doc));
      }

      XmlElement deliveryDocs = doc.CreateElement("deliveryDocsPack");
      _ = root.AppendChild(deliveryDocs);
      foreach (string template in DeliveryDocsPack)
      {
        _ = deliveryDocs.AppendChild(H.CreateElement(doc, "deliveryDocs", template));
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

        try
        {
          ToolData data = new(
          row["RESOURCE"]?.ToString() ?? string.Empty,
          row["CODE"]?.ToString() ?? string.Empty,
          int.TryParse(row["CALL"]?.ToString(), out int call) ? call : 1,
          row["NAME"]?.ToString() ?? string.Empty,
          row["FILE TYPE"]?.ToString() ?? string.Empty,
          bool.TryParse(row["THREAD_SAFE"]?.ToString(), out bool isThreadSafe) ? isThreadSafe : true,
          int.TryParse(row["TIMEOUT"]?.ToString(), out int timeoutValue) ? timeoutValue : 0,
          bool.TryParse(row["ITERATIVE"]?.ToString(), out bool isIterative) ? isIterative : false,
          row["INTERPRETER"]?.ToString() ?? string.Empty,
          int.TryParse(row["VERSION"]?.ToString(), out int value) ? value.ToString("D3") : "000",
          row["LANGUAGE"]?.ToString() ?? string.Empty,
          row["DESCRIPTION"]?.ToString() ?? string.Empty
      );

          Tools.Add(data);

        }
        catch (Exception)
        {

          throw new Exception("Error de lectura de los parámetros de las herramientas en ToolMap excel");
        }

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
      var filteredTools = Tools.Where(tool => tool.Code.Equals(code, StringComparison.OrdinalIgnoreCase) && tool.Language.Equals(language, StringComparison.OrdinalIgnoreCase)).ToList();

      if (filteredTools.Count == 0) //try default language
      {
        filteredTools = Tools.Where(tool => tool.Code.Equals(code, StringComparison.OrdinalIgnoreCase) && tool.Language.Equals(H.GetSProperty("defaultLanguage"), StringComparison.OrdinalIgnoreCase)).ToList();

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
      //reading metadata
      MirrorXML mirror = new(tool);
      var variableMap = mirror.VarList;

      string originalToolPath = Path.Combine(
          tool.Resource == "TOOL" ? H.GetSProperty("ToolsPath") : H.GetSProperty("TemplatesPath"),
          tool.FileName
      );

      string toolPath = "";
      List<string> indexer = tool.IsIterative ? indexer = dm.GetValueList($"{tool.Code}.indexer", false) : new() { "" };
      int i = -1;
      string timestamp;
      List<string> varTracking = new(); //list to keep track of the variables already created in the results xml (to avoid duplicates in iterative tools)

      XmlDocument results = new(); // Declarado fuera del try para asegurar retorno


      XmlElement answer = results.CreateElement("answer");
      _ = results.AppendChild(answer);
      XmlElement varNode = results.CreateElement("variables");
      _ = answer.AppendChild(varNode);


      while (++i < indexer.Count)
      {
        {
          toolPath = Path.Combine(
            H.GetSProperty("processPath"),
            dm.GetInnerText($@"dm/utils/utilsData/opportunityFolder"),
            dm.SBidRevision,
            "TOOLS",
            Path.GetFileNameWithoutExtension(tool.FileName) + (tool.IsIterative ? $"_{i + 1}-{indexer[i]}" : "") + Path.GetExtension(tool.FileName));

          //create folder (if does not exists) and copy the tool to process folder
          _ = Directory.CreateDirectory(Path.GetDirectoryName(toolPath)!);

          //copy tool template only for first call, check for the existance of the file toolPath
          if (!File.Exists(toolPath))
            File.Copy(originalToolPath, toolPath, true);

          H.PrintLog(2,
            TC.ID.Value!.Time(),
            TC.ID.Value!.User,
            "CalculateExcel",
            $"  Calling tool {tool.Code} {(tool.IsIterative ? $" <iteration: {i + 1}-{indexer[i]}  " : "")}");

          SB_Excel? workbook = null;

          try
          {
            workbook = new SB_Excel(toolPath);

            // Escribir valores en las celdas
            foreach (var entry in variableMap)
            {
                string variableID = entry.Key;
                string direction = entry.Value[1];

                if (direction == "in")
                {
                  string rangeName = tool.InPrefix + variableID;
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
                  else if (type == "list<num>" || type == "list<str>")
                  {
                    //lists in iterative tools are used to provide values for each iteration. Funcionality of lists are modified

                    if (!tool.IsIterative)//Is Not Iterative
                    {
                      //in non iterative tools all the values of the list are written to the workbook
                      List<string> listData = dm.GetValueList(variableID, type == "list<num>");

                      // call fillupValue for each item in the listData
                      foreach (var (item, index) in listData.Select((value, j) => (value, j)))
                      {
                        _ = workbook.FillUpValue($"{rangeName}\\{index}", item);
                      }

                    }
                    else //Is Iterative
                    {
                      //in iterative tools only one value of the list is used for each iteration
                      string? cellValue = dm.GetValueList(variableID, type == "list<num>").ElementAtOrDefault(i);
                      if (cellValue != null)
                      {
                        try
                        {
                          _ = workbook.FillUpValue(rangeName, cellValue);
                        }
                        catch (Exception ex)
                        {
                          if (rangeName.Contains(tool.Code + ".indexer"))
                          {
                            //for the indexer variable of iterative tools do nothing, this var is only to tracking the iterations
                            continue;
                          }
                          else
                          {
                            throw new Exception($"Error filling value for variable '{variableID}' at iteration {i}.", ex);
                          }
                        }
                      }
                      else
                      {
                        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"No list data found for variable '{variableID}' at iteration {i}.");
                      }
                    }
                  }
                  else
                  {
                    _ = workbook.FillUpValue(rangeName, dm.GetValueString(variableID));
                  }
                }
             
            }

            workbook.Calculate();

            // Leer valores de las celdas de salida
            foreach (var entry in variableMap)
            {
              string variableID = entry.Key;
              string direction = entry.Value[1];
              string type = _variablesMap.GetVariableData(variableID).Type;


              if (direction == "out")
              {
                string rangeName = tool.OutPrefix + variableID;
                XmlElement? varElement = null;
                XmlElement value;
                XmlElement note;
                try
                {
                  //for the call1 variables remove the added Call1_ prefix (added for normalization)
                  if (Regex.IsMatch(rangeName, @"(?i)Call1_"))
                    rangeName = Regex.Replace(rangeName, @"(?i)Call1_", "");



                  //Keep tracking of the variables already created for avoid duplicates in iterative tools
                  if (!varTracking.Contains(variableID))
                  {
                    varElement = results.CreateElement(variableID);
                    value = results.CreateElement("value");
                    value.SetAttribute("type", type);
                    _ = varElement.AppendChild(value);
                    note = results.CreateElement("note");
                    note.InnerText = "Calculated Value";
                    _ = varElement.AppendChild(note);
                    _ = varElement.AppendChild(H.CreateElement(results, "origin", $"{tool.Code}+{DateTime.Now.ToString("dd-HH:mm")}"));

                    varTracking.Add(variableID);
                  }
                  else
                  {
                    varElement = (XmlElement)results.SelectSingleNode($"//answer/variables/{variableID}")!;
                    value = (XmlElement)varElement.SelectSingleNode("value")!;
                    note = (XmlElement)varElement.SelectSingleNode("note")!;
                  }


                  if (type == "table")
                  {
                    XmlNode tableXmlValue = results.ImportNode(workbook.GetTValue(rangeName), true);

                    if (tableXmlValue != null && tableXmlValue.HasChildNodes)
                    {
                      _ = value.AppendChild(tableXmlValue);
                      note.InnerText = "Calculated Value";
                    }
                    else
                    {
                      value.InnerText = "No data found in table.";
                    }
                  }
                  else if (type == "list<num>" || type == "list<str>")
                  {
                    if (!tool.IsIterative)
                    {
                      //we cannot read all list values at once, we're going to store the values in a  <List<string> by searching for values in cellRanges named rangeName\0, rangeName\1, rangeName\2... until we find no more values. Then with the list complete we will create the <l><li>item1</li><li>item2</li>...</l> structure to be added to the <value>

                      bool isNumber = type == "list<num>";
                      List<string> listValues = new List<string>();
                      int index = 0;
                      while (true)
                      {
                        // check the existance of the range name before calling to get the value

                        List<string> namedRanges = workbook.ListNamedRanges();
                        string namedRangeToFind = $"{rangeName}\\{index}";
                        string? cellValue = null;

                        if (namedRanges is not null)
                        {
                          var match = namedRanges
                              .FirstOrDefault(n => string.Equals(n, namedRangeToFind, StringComparison.OrdinalIgnoreCase));

                          if (match is not null)
                          {
                            // Use the matched original-casing name
                            cellValue = workbook.GetSValue(match, isNumber);
                          }
                        }
                        if (cellValue != null && cellValue != string.Empty)
                        {
                          listValues.Add(cellValue);
                          index++;
                        }
                        else
                        {
                          break;
                        }
                      }

                      if (listValues.Count > 0)
                      {

                        value.SetAttribute("type", type);
                        XmlElement listElement = results.CreateElement("l");
                        foreach (string item in listValues)
                        {
                          _ = listElement.AppendChild(H.CreateElement(results, "li", item));
                        }
                        _ = value.AppendChild(listElement);
                        note.InnerText = "Calculated Value";

                      }
                    }
                    else //IsIterative
                    {
                      //Si es iterativo tratamos cada valor de la lista como un valor único
                      if (rangeName.Contains(tool.Code + ".indexer"))
                      {
                        //for the indexer variable of iterative tools do nothing, this var is only to tracking the iterations
                        continue;
                      }

                      string cellValue = workbook.GetSValue(rangeName, type == "list<num>");

                      if (cellValue != null && cellValue != string.Empty)
                      {
                        //for iterative tools values type list should be stored as xml lists like <l><li>value</li></l> according to the index (i) of the iteration.
                        // first get the actual value as xmlNode, if it's empty create a new list <l>
                        // create a new element to add <li/>
                        // add the value to this element
                        // insert the li element to the l element

                        XmlElement listElement;
                        if (value.HasChildNodes)
                        {
                          listElement = (XmlElement)value.SelectSingleNode("l")!;
                        }
                        else
                        {
                          listElement = results.CreateElement("l");
                          _ = value.AppendChild(listElement);
                        }
                        XmlElement listItem = results.CreateElement("li");
                        listItem.InnerText = cellValue;
                        _ = listElement.AppendChild(listItem);

                      }
                      else
                      {
                        value.InnerText = _variablesMap.GetVariableData(variableID).Default ?? string.Empty;
                        note.InnerText = "Default Value";

                        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{toolPath}'. Default value used");
                      }

                    }
                  }
                  else
                  {
                    string cellValue = workbook.GetSValue(rangeName, type == "num");

                    if (cellValue != null && cellValue != string.Empty)
                    {
                      value.InnerText = cellValue;
                      note.InnerText = "Calculated Value";
                    }
                    else
                    {
                      value.InnerText = _variablesMap.GetVariableData(variableID).Default ?? string.Empty;
                      note.InnerText = "Default Value";

                      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{toolPath}'. Default value used");
                    }
                  }
                  answer.SetAttribute("result", "OK");
                }
                catch (Exception ex)
                {
                  answer.SetAttribute("result", "NO OK");
                  H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"❌Error❌ reading range '{rangeName}': {ex.Message}");
                }

                _ = varNode.AppendChild(varElement);
              }
            }


            H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExcel", $"Values returned from {tool.Code} calculation", results);
          }
          catch { throw; }
          finally
          {
            workbook?.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
          }
        }

        // Adding information for Utils
        // Check whether <utils> exists in results and append the <utils> node if it doesn't
        XmlElement utilsNode = (XmlElement)results.SelectSingleNode("//answer/utils");
        if (utilsNode == null)
        {
          utilsNode = results.CreateElement("utils");
          _ = results.DocumentElement.AppendChild(utilsNode);
        }

        // Check whether <tools> exists in <utils>  and append the <tools> node if it doesn't
        XmlElement toolsNode = (XmlElement)utilsNode.SelectSingleNode("tools");
        if (toolsNode == null)
        {
          toolsNode = results.CreateElement("tools");
          _ = utilsNode.AppendChild(toolsNode);
        }

        XmlElement toolNode = (XmlElement)toolsNode.SelectSingleNode(tool.Code);
        if (toolNode == null)
        {
          XmlElement _toolInfo = tool.ToXML(results);

          toolNode = results.CreateElement(tool.Code);
          foreach (XmlAttribute attr in _toolInfo.Attributes)
          {
            if (attr.Name != "code")
            {
              if (attr.Name != "fileName")
                toolNode.SetAttribute(attr.Name, attr.Value);
              else
                toolNode.SetAttribute(attr.Name, toolPath);
            }
          }
          _ = toolsNode.AppendChild(toolNode);
        }

        if (tool.IsIterative)
        {
          _ = toolNode.AppendChild(H.CreateElement(results, "note", $"Iterative call: {dm.GetValueList($"{tool.Code}.indexer", false)[i]}"));
        }
      }
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
      string processType = Path.GetExtension(filePath).Equals(".py", StringComparison.OrdinalIgnoreCase) ? "python" : "exe";


      if (File.Exists(filePath))
      {
        if (processType == "python")
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
      _ = callXml.AppendChild(callNode);

      XmlElement variables = callXml.CreateElement("variables");
      _ = callNode.AppendChild(variables);
      XmlElement variablesIn = callXml.CreateElement("in");
      _ = variables.AppendChild(variablesIn);
      XmlElement variablesOut = callXml.CreateElement("out");
      _ = variables.AppendChild(variablesOut);

      foreach (var entry in variableList)
      {
        string variableID = entry.Key;
        string direction = entry.Value[1];

        if (direction == "in" && mirror.GetVarCallLevel(variableID) == tool.Call)
        {
          XmlElement varElement = callXml.CreateElement(variableID);
          varElement.SetAttribute("unit", dm.GetValueUnit(variableID));
          varElement.InnerText = dm.GetValueString(variableID);
          _ = variablesIn.AppendChild(varElement);
        }
        else if (direction == "out" && mirror.GetVarCallLevel(variableID) == tool.Call)
        {
          XmlElement varElement = _variablesMap.GetVariableData(variableID).ToXML(callXml);
          _ = variablesOut.AppendChild(varElement);
        }
      }

      string xmlVarList = callXml.OuterXml;

      H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"   Calling Tool: {Path.GetFileName(filePath)} {arguments}");
      H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"   Calling Tool: {filePath} {arguments}");
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", "1   call message:", callXml);


      // SmartSize does not accept Encoding.UTF8 in the ProcessStartInfo StandardInputEncoding property
      // it works using the default value.
      ProcessStartInfo psi;
      if (processType == "python")
      {
        psi = new()
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
      }
      else
      {
        psi = new()
        {
          FileName = filePath,
          Arguments = arguments,
          RedirectStandardInput = true,
          RedirectStandardOutput = true,
          RedirectStandardError = true,
          UseShellExecute = false,
          CreateNoWindow = true,
        };
      }


      string output;
      string error;

      //// Use the current call's CancellationToken (from TC.ThreadInfo)
      //var ct = TC.ID.Value?.Token ?? CancellationToken.None;

      //Stopwatch sw = Stopwatch.StartNew();

      //using (Process process = new() { StartInfo = psi })
      //{
      //  //take the timestamp when the process starts to print out the total time spent in the process


      //  _ = process.Start();

      //  // Register this external process to be auto-killed on cancel/timeout
      //  TC.ID.Value?.RegisterProcess(process);

      //  // If the call is cancelled, kill the process (and its tree)
      //  using var cancelReg = ct.Register(() =>
      //  {
      //    try
      //    {
      //      if (!process.HasExited)
      //      {
      //        try { process.Kill(true); } catch { process.Kill(); }
      //      }
      //    }
      //    catch { /* ignore */ }
      //  });

      //  using var _scoped = TC.ID.Value!.ArmScopedTimeoutMinutes(
      //      tool.TimeoutMinutes,
      //      reason: $"Timeout Tool {tool.Code}",
      //      killChildren: true
      //  );

      //  // Send input and read outputs (same as you already do)
      //  using (StreamWriter writer = process.StandardInput)
      //  {
      //    writer.Write(xmlVarList);
      //    writer.Flush();
      //    writer.Close();
      //  }

      //  // NOTE: These ReadToEnd() calls will now unblock if the process is killed by cancel/timeout
      //  output = process.StandardOutput.ReadToEnd();
      //  error = process.StandardError.ReadToEnd();

      //  // Wait for exit (fast returns if already killed)
      //  process.WaitForExit();
      //}

      //sw.Stop();

      //ReleaseProcess(tool.Code, (int)TC.ID.Value.CallId!);


      //// If cancelled, surface a meaningful error
      ///


      //if (ct.IsCancellationRequested)
      //{     
      //  //show a canellation exception telling the total time spent until cancellation and the time since the tool calling (whenever the watchdog started)
      //  throw new OperationCanceledException($"Tool {tool.Code} was timeout after {sw.Elapsed.Minutes} min.");
      //}

      /////////////////////////////////////////////////////////////////////////////////////////////
      ///
      ///
      ///
      ///
      ///
      ///
      ///

      // Use the current call's CancellationToken (from TC.ThreadInfo)
      var ct = TC.ID.Value?.Token ?? CancellationToken.None;

      Stopwatch sw = Stopwatch.StartNew();

        var manualCts = new CancellationTokenSource();
        var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct, manualCts.Token);
        var effectiveCt = linkedCts.Token;

        if(tool.TimeoutMinutes != 0 ) Task.Delay(TimeSpan.FromMinutes(tool.TimeoutMinutes - 2))
            .ContinueWith(_ =>
            {
                if (!ct.IsCancellationRequested) // solo si aún no se disparó el timeout
                {
                    H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", "Cancelando ejecucion manualmente. 2 minutos para el timeout", callXml);
                    manualCts.Cancel();
                }
            });


      using (Process process = new() { StartInfo = psi })
      {
        //take the timestamp when the process starts to print out the total time spent in the process

        _ = process.Start();

        // Register this external process to be auto-killed on cancel/timeout
        TC.ID.Value?.RegisterProcess(process);

        // If the call is cancelled, kill the process (and its tree)
        using var cancelReg = effectiveCt.Register(() =>
        {
          try
          {
            if (!process.HasExited)
            {
              try { process.Kill(true); } catch { process.Kill(); }
            }
          }
          catch { /* ignore */ }
        });

        using var _scoped = TC.ID.Value!.ArmScopedTimeoutMinutes(
            tool.TimeoutMinutes,
            reason: $"Timeout Tool {tool.Code}",
            killChildren: true
        );

        // Send input and read outputs (same as you already do)
        using (StreamWriter writer = process.StandardInput)
        {
          writer.Write(xmlVarList);
          writer.Flush();
          writer.Close();
        }

        // NOTE: These ReadToEnd() calls will now unblock if the process is killed by cancel/timeout
        output = process.StandardOutput.ReadToEnd();
        error = process.StandardError.ReadToEnd();

        // Wait for exit (fast returns if already killed)
        process.WaitForExit();
      }

      sw.Stop();

      ReleaseProcess(tool.Code, (int)TC.ID.Value.CallId!);


        // If cancelled, surface a meaningful error
        if (ct.IsCancellationRequested && !manualCts.IsCancellationRequested)
        {
            // Timeout real
            throw new OperationCanceledException(
                $"Tool {tool.Code} timed out after {sw.Elapsed.Minutes} min.");
        }
        else if (manualCts.IsCancellationRequested)
        {
                H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", "Intentnado continuar con la ejecucion.", callXml);
            }

        ///
        ///
        ///
        ///
        ///
        ///
        ///
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

      // Validate output to avoid XML parse crash after a forced kill
      if (string.IsNullOrWhiteSpace(output))
      {
        var msg = $"Tool {tool.Code} produced no output.";
        if (!string.IsNullOrWhiteSpace(error)) msg += $" Stderr: {error}";
        throw new Exception(msg);
      }

      XmlDocument results = new();

      try
      {
        results.LoadXml(output);
      }
      catch (Exception ex)
      {
        // Include some stderr to help debug malformed XML cases
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe",
                   $"❌ Invalid XML from tool '{tool.Code}'. Error: {ex.Message}\nSTDERR:\n{error}");
        throw;
      }

      H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", $"Return from tool {tool.Code} after {sw.Elapsed.Minutes} minutes.", results);

      if (!string.IsNullOrWhiteSpace(error))
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", $"❌Error❌:\n{error}");

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculate", "-----------------------------------");

      // Adding information for Utils
      // Check whether <utils> exists in results and append the <utils> node if it doesn't
      XmlElement utilsNode = (XmlElement)results.SelectSingleNode("//answer/utils");
      if (utilsNode == null)
      {
        utilsNode = results.CreateElement("utils");
        _ = results.DocumentElement.AppendChild(utilsNode);
      }

      // Check whether <tools> exists in <utils>  and append the <tools> node if it doesn't
      XmlElement toolsNode = (XmlElement)utilsNode.SelectSingleNode("tools");
      if (toolsNode == null)
      {
        toolsNode = results.CreateElement("tools");
        _ = utilsNode.AppendChild(toolsNode);
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
      _ = toolsNode.AppendChild(newToolNode);


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
          XmlElement varElement = CreateXmlVariable(results,
                                                  variableID,
                                                  _variablesMap.GetVariableData(variableID).Type,
                                                  _variablesMap.GetVariableData(variableID).Default ?? string.Empty,
                                                  $"{tool.Code}+{DateTime.Now:dd-HH:mm}",
                                                  "Calculated Value");

          _ = resultsVariablesNode.AppendChild(varElement);
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CalculateExe", $"❌ Missing output variable '{variableID}' added with default value.");
        }
      }
      return results;
    }

    private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, string value, string origin, string? note = null)
    {
      XmlElement xmlVar = outputXML.CreateElement(ID);
      _ = (outputXML.SelectSingleNode("//answer/variables")?.AppendChild(xmlVar));

      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "value", value, ("type", type)));
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "origin", $"{origin}-{DateTime.Now:yyMMdd-HHmm}"));
      if (note != null)
        _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "note", note));

      return xmlVar;
    }

    private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, XmlElement value, string origin, string? note = null)
    {
      XmlElement xmlVar = outputXML.CreateElement(ID);
      _ = (outputXML.SelectSingleNode("//answer/variables")?.AppendChild(xmlVar));

      XmlElement valueElement = (XmlElement)xmlVar.AppendChild(value!)!;
      valueElement.SetAttribute("type", type);
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "origin", $"{origin}-{DateTime.Now:yyMMdd-HHmm}"));
      if (note != null)
        _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "note", note));

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

          doc.ReplaceFieldMarks(varList, dm);

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
            else if (type == "list<num>" || type == "list<str>")
            {
              //if list<num> isNumber set to true, if list<str> isNumber set to false
              bool isNumber = type == "list<num>";

              List<string> listData = dm.GetValueList(variableID, isNumber);

              // call fillupValue for each item in the listData
              foreach (var (item, index) in listData.Select((value, i) => (value, i)))
              {
                _ = doc.FillUpValue($"{rangeName}\\{index}", item);
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
