using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SmartBid
{
  internal class Calculator
  {
    List<string> _targets;
    List<string> _calcRoute;
    ToolsMap tm;
    DataMaster dm;

    public Calculator(DataMaster dataMaster, XmlDocument call)
    {
      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "Calculator", $"REQUEST:");

      this.dm = dataMaster;

      _targets = GetDeliveryDocs(call);
    }
    public void RunCalculations()
    {
      tm = ToolsMap.Instance;

      //Find the list of variable to get from Preparation Tool
      string xmlPrepVarList = GetRouteMap(_targets).OuterXml;

      //Buscamos los datos necesarios con la PreparationTool y los guardamos en el DataMaster
      dm.UpdateData(CallPrepTool(xmlPrepVarList));
      dm.CheckMandatoryValues(); // thows and exception if not all Mandatory values are present in the DataMaster
      dm.SaveDataMaster(); //Save the DataMaster after preparation

      //Generate files structure and move input files
      //Call each tool in the list of _targets and update the DataMaster with the results

      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "RunCalculations", $"rute: {string.Join(" > ", _calcRoute)}");


      //CALCULATE
      foreach (string target in _calcRoute)
      {
        if (tm.Tools.Exists(tool => tool.ID == target))
        {
          ToolData toolData = tm.Tools.First(tool => tool.ID == target);
          if (toolData.Resource == "TOOL")
          {
            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "RunCalculations", $"Calling Tool: {toolData.ID} - {toolData.Description}");

            //Call calculation
            dm.UpdateData(tm.CalculateExcel(target, dm)); //Update the DataMaster with the results from the tool
            dm.SaveDataMaster(); //Save the DataMaster after each calculation
          }
        }
        else
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, $"❌❌ Error ❌❌  - RunCalculations", $"Tool {target} not found in ToolsMap.");
        }
      }

      //GENERATE DOCUMENTS
      foreach (string target in _targets)
      {
        if (tm.Tools.Exists(tool => tool.ID == target))
        {
          ToolData toolData = tm.Tools.First(tool => tool.ID == target);
          if (toolData.Resource == "TEMPLATE")
          {
            H.PrintLog(3, ThreadContext.CurrentThreadInfo.Value.User, "RunCalculations", $"Populating Template: {toolData.ID} - {toolData.Description}");

            //Call generate output document
            tm.GenerateOuputWord(target, dm);

            // GENERATE INFO ABOUT TEMPLATES GENERATED (PENDIENTE)
            // dm.UpdateData(NEW_INFO); //Update the DataMaster with the information about generated documents
          }
        }
        else
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, $"❌❌ Error ❌❌  - RunCalculations", $"Template {target} not found in ToolsMap.");
        }
      }

      _ = DBtools.InsertNewProjectWithBid(dm);

    }
    private XmlDocument CallPrepTool(string xmlVarList)
    {
      XmlDocument prepCall = new XmlDocument();
      prepCall.LoadXml(xmlVarList); // Load the XML string into the XmlDocument 
      string prepToolPath = Path.GetFullPath(H.GetSProperty("PreparationTool"));

      H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", $"\n");
      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", $"- CALLING PREPARATION: {prepToolPath} ------------------");
      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", "- ARGUMENTO PASADO A PREPTOOL:");
      H.PrintXML(2, prepCall); // Print the XML for debugging
      H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", $"\n\n\n\n");
      H.SaveXML(1, prepCall, Path.Combine(H.GetSProperty("processPath"),dm.GetValueString("opportunityFolder"),  "prepCall.xml")); // Save the XML to a file for debugging



      ProcessStartInfo psi = new ProcessStartInfo
      {
        FileName = prepToolPath,  // Path to the executable
        RedirectStandardInput = true,  // Send input through StandardInput
        RedirectStandardOutput = true, // Capture output
        RedirectStandardError = true,  // Capture errors
        UseShellExecute = false,
        CreateNoWindow = true,
        StandardInputEncoding = Encoding.UTF8  // Ensure proper encoding
      };
      string output;
      string error;

      using (Process process = new Process { StartInfo = psi })
      {
        _ = process.Start();

        using (StreamWriter writer = process.StandardInput)
        {
          writer.Write(xmlVarList);  // Send XML via Standard Input
          writer.Flush();  // Ensure all data is sent
          writer.Close();  // Signal EOF
        }
        output = process.StandardOutput.ReadToEnd();
        error = process.StandardError.ReadToEnd();
        process.WaitForExit();
      }

      XmlDocument xmlPrepAnswer = new XmlDocument();

      xmlPrepAnswer.LoadXml(output); // Load the XML content

      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", $"Return from Preparation");
      H.PrintXML(2, xmlPrepAnswer);
      H.SaveXML(1, xmlPrepAnswer, Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "PrepAnswer.xml")); // Save the XML to a file for debugging

      if (error != "") H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", $"❌Error❌:\n{error}");
      H.PrintLog(0, ThreadContext.CurrentThreadInfo.Value.User, "CallPrepTool", "-----------------------------------");

      return xmlPrepAnswer;
    }
    public XmlDocument GetRouteMap(List<string> targets)
    {
      List<string> sourcesSearched = new List<string>();
      sourcesSearched.AddRange(new[] { "INIT", "AUTO", "PREP" });//Adding souces that does not need to be searched

      List<VariableData> prepVarList = new List<VariableData>();
      List<List<string>> calcTools = new List<List<string>>(); //List to keep track of the calculation tools used in the recursion
      XmlDocument prepCallXML;

      prepVarList.AddRange(Get_PREP_Variables(targets, sourcesSearched, 0, calcTools));

      // Build up the list of calls to tools in order
      _calcRoute = new List<string>();
      HashSet<string> uniqueElements = new HashSet<string>(_calcRoute);
      for (int i = calcTools.Count - 1; i >= 0; i--) // Iterate backwards through calcTools
      {
        foreach (var element in calcTools[i])
        {
          if (uniqueElements.Add(element)) // Add only if it's not already present
          {
            _calcRoute.Add(element);
          }
        }
      }

      targets = _calcRoute; // Update _targets with the ordered list of calculation tools

      _ = prepVarList.RemoveAll(item => !(item.Source == "PREP")); //Remove all non-preparation source variables

      List<string> prepVarIDList = prepVarList.Select(variable => variable.ID).ToList(); //Getting the list of IDs for the preparation variables  

      prepCallXML = VariablesMap.Instance.ToXml(prepVarIDList);


      XmlNode inputDocsNode = H.CreateElement(prepCallXML, "inputDocs", ""); //Adding the call element to the XML   
      _ = prepCallXML.DocumentElement.AppendChild(inputDocsNode); //Adding the call element to the XML


      string inputFilesTimeStamp = dm.DM.SelectSingleNode(@$"dm/utils/rev_{dm.BidRevision:D2}/inputDocs")?.Attributes?["timeStamp"]?.Value ?? DateTime.Now.ToString("yyMMdd");


      foreach (XmlElement doc in dm.DM.SelectNodes(@$"dm/utils/rev_{dm.BidRevision.ToString("D2")}/inputDocs/doc"))
      {
        string fileType = doc.GetAttribute("type");
        string fileName = doc.InnerText;
        string filePath = Path.Combine(Path.Combine
          (H.GetSProperty("oppsPath"),
          dm.GetInnerText(@"dm/utils/utilsData/opportunityFolder")),
          "1.DOC",
          inputFilesTimeStamp,
          fileType,
          fileName);

        XmlElement docElement = H.CreateElement(prepCallXML, "doc", filePath);
        docElement.SetAttribute("type", fileType);
        _ = inputDocsNode.AppendChild(docElement); //Adding the input documents to the XML
      }

      // Guardamos el XML por si fuese necesario
      prepCallXML.Save(Path.Combine(Path.GetDirectoryName(H.GetSProperty("ToolsPath")), "preparationCall.xml"));
      return prepCallXML; //Returning all variables to be read at Preparation in XML format
    }
    private static List<VariableData> Get_PREP_Variables(List<string> targets, List<string> sourcesExcluded, int deep, List<List<string>> calcTools)
    {

      deep++; //Registering the depth of the recursion

      if (deep > 10)
      {
        throw new Exception("Entrada en recursividad");
      }


      calcTools.Add(new List<string>(targets)); //Adding the current _targets to the list of calculation tools

      if (targets.Count == 0)
        return new List<VariableData>(); 

      VariablesMap varMap = VariablesMap.Instance;
      List<string> newSources = new List<string>();
      List<VariableData> variableList = new List<VariableData>();

      foreach (string target in targets) //Iterating through all _targets
      {
        MirrorXML mirror = new MirrorXML(target); //Processing all _targets of this deep
        int call = 1; // Valor predeterminado si no se encuentra un número

        // Verificar si el string termina en "_Callx" donde x es un numeral del 0 al 9
        Match match = Regex.Match(target, @"_Call(\d)$");
        if (match.Success)
        {
          call = int.Parse(match.Groups[1].Value);
          //                        varName = varName.Substring(0, varName.LastIndexOf("_Call")); // Modificar v para eliminar "_Cx"
        }

        foreach (var item in mirror.VarList.Keys)

        {
          string varName = item;

          if (mirror.VarList[varName][1] == "in" && call == Convert.ToInt16(mirror.VarList[varName][2]))
          {
            VariableData variableData = varMap.GetNewVariableData(varName); // Retrieve the variable data
            variableData.InOut = "in"; // Set the InOut property
            variableData.Call = Convert.ToInt16(mirror.VarList[varName][2]);// Set the Call property
            variableData.Deep = deep; // Set the depth property
            variableList.Add(variableData); // Add it to the list
          }
        }
      }

      variableList = variableList.Distinct().ToList(); //removing duplicates

      newSources.AddRange(variableList.Select(item => item.Source));

      newSources = newSources.Distinct().ToList(); //removing duplicates
      newSources = newSources.Except(sourcesExcluded).ToList(); //removing especial sources that do not need to be searched

      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "GetRouteMap", $"newSources deep ({deep}): {string.Join(", ", newSources)}");
      List<string> varlist = variableList.Select(variable => variable.ID).ToList(); //Getting the list of IDs for the variables found
      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "GetRouteMap", $"varlist: {string.Join(", ", varlist)}");


      variableList.AddRange(Get_PREP_Variables(newSources, sourcesExcluded, deep, calcTools)); //Recursion to process the next target

      variableList = variableList.Distinct().ToList();

      return variableList;
    }
    public static List<string> GetDeliveryDocs(XmlDocument xmlDoc)
    {
      XmlNode deliveryDocsNode = xmlDoc.SelectSingleNode("/request/requestInfo/deliveryDocs");
      List<string> deliveryDocs = new List<string>();

      if (deliveryDocsNode != null) foreach (XmlNode docNode in deliveryDocsNode.SelectNodes("doc")) deliveryDocs.Add(docNode.InnerText);

      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "GetDeliveryDocs", $"-----------EXECUTING THE FOLLOWING DOCUMENTS:-----------\n{string.Join("\n", deliveryDocs)}\n\n");

      return deliveryDocs;
    }
    public static string ReadFileContent(string filePath)
    {
      if (File.Exists(filePath))
      {
        return File.ReadAllText(filePath);
      }
      else
      {
        throw new FileNotFoundException($"El archivo '{filePath}' no existe.");
      }
    }

  }

}
