using System.Collections.Immutable;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Microsoft.Office.Interop.Word;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;

namespace SmartBid
{
  internal class Calculator
  {
    readonly List<ToolData> _targets;
    List<ToolData> _calcRoute = [];
    ToolsMap? tm;
    DataMaster dm;

    public Calculator(DataMaster dataMaster, List<ToolData> targets)
    {
      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "Calculator", $"REQUEST:");

      this.dm = dataMaster;

      _targets = targets;
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
      //Call each toolD in the list of _targets and update the DataMaster with the results

      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"***  CALCULATE  ***");
      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"Calculate rute: {string.Join(" >> ", _calcRoute.Select(tool => tool.Code))}");

      //CALCULATE
      foreach (ToolData tool in _calcRoute)
      {
        if (tm.Tools.Exists(tool => tool.Code == tool.Code))
        {
          if (tool.Resource == "TOOL")
          {
            H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"Calling Tool: {tool.Code}");
            H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"Calling Tool: -- {tool.Description}");


            //Call calculation
            XmlDocument CalcResults = tm.Calculate(tool, dm);

            //Update the DataMaster with the results from the toolD
            dm.UpdateData(CalcResults);

            //Save the DataMaster after each calculation
            dm.SaveDataMaster(); //Save the DataMaster after each calculation
          }
        }
        else
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, $"❌❌ Error ❌❌  - RunCalculations", $"Tool {tool} not found in ToolsMap.");
        }
      }
      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"- Calculate Done -");
      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"***   GENERATE DOCUMENTS   ***");
      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"Generation rute: {string.Join(" >> ", _targets.Select(tool => tool.Code))}");

      //GENERATE DOCUMENTS
      foreach (ToolData target in _targets)
      {
        if (target.Resource == "TEMPLATE")
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "RunCalculations", $"Populating Template: {target.Code} - {target.Description}");

          tm.GenerateOuput(target, dm);

          // GENERATE INFO ABOUT TEMPLATES GENERATED (PENDIENTE)
          // dm.UpdateData(NEW_INFO); //Update the DataMaster with the information about generated documents
        }
      }

      _ = DBtools.InsertNewProjectWithBid(dm);

    }
    private XmlDocument CallPrepTool(string xmlVarList)
    {
      XmlDocument prepCall = new();
      prepCall.LoadXml(xmlVarList); // Load the XML string into the XmlDocument 

      H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", $"\n");
      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", $"***   PREPARATION   ***");
      H.PrintLog(0, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", "- Argumento pasado a PREP:", prepCall);
      H.PrintLog(3, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", $"\n\n");
      H.SaveXML(3, prepCall, Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "prepCall.xml")); // Save the XML to a file for debugging


      XmlDocument xmlPrepAnswer = PREP.Run(prepCall);


      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", $"Return from Preparation", xmlPrepAnswer);
      H.SaveXML(1, xmlPrepAnswer, Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "PrepAnswer.xml")); // Save the XML to a file for debugging

      H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "CallPrepTool", "- Preparation Done -");

      return xmlPrepAnswer;
    }
    public XmlDocument GetRouteMap(List<ToolData> targets)
    {
      List<string> sourcesSearched =
      [
        .. new[] { "INIT", "AUTO", "PREP", "UTILS" },//Adding souces that does not need to be searched
      ];

      List<VariableData> prepVarList = [];
      List<List<ToolData>> calcTools = []; //List to keep track of the calculation tools used in the recursion
      XmlDocument prepCallXML = new();
      XmlDeclaration xmlDeclaration = prepCallXML.CreateXmlDeclaration("1.0", "UTF-8", null);
      prepCallXML.AppendChild(xmlDeclaration);
      XmlElement root = prepCallXML.CreateElement("call");
      prepCallXML.AppendChild(root);

      prepVarList.AddRange(Get_PREP_Variables(targets, sourcesSearched, 0, calcTools));

      // Build up the list of calls to tools in order
      _calcRoute = [];
      HashSet<ToolData> uniqueElements = new(_calcRoute);
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

      XmlElement variablesXml = VariablesMap.Instance.ToXml(prepCallXML, prepVarIDList);
      _ = root.AppendChild(variablesXml); //Adding the variables element to the XML


      XmlNode? dmInputDocs = dm.DM.SelectSingleNode(@$"dm/utils/rev_{dm.BidRevision.ToString("D2")}/inputDocs");
      if (dmInputDocs != null)
        prepCallXML.DocumentElement.AppendChild(prepCallXML.ImportNode(dmInputDocs, true));

      H.PrintLog(5,ThreadContext.CurrentThreadInfo.Value!.User, "GetRouteMap", $"- Calculated Route Map: \n {string.Join(" >> ", _calcRoute.Select(tool => tool.Code))}");

      // Guardamos el XML por si fuese necesario
      prepCallXML.Save(Path.Combine(Path.GetDirectoryName(H.GetSProperty("ToolsPath")), "preparationCall.xml"));
      return prepCallXML; //Returning all variables to be read at Preparation in XML format
    }
    private static List<VariableData> Get_PREP_Variables(List<ToolData> targets, List<string> sourcesExcluded, int deep, List<List<ToolData>> calcTools)
    {
      deep++; // Registering the depth of the recursion

      if (deep > 10)
        throw new Exception("Recursion limit exceeded");

      calcTools.Add(new List<ToolData>(targets));

      if (targets.Count == 0)
        return [];

      VariablesMap varMap = VariablesMap.Instance;
      ToolsMap tm = ToolsMap.Instance;

      List<ToolData> newSources = [];
      List<VariableData> variableList = [];

      foreach (ToolData target in targets)
      {
        MirrorXML mirror = new(target);
        int call = 1;

        Match match = Regex.Match(target.Code, @"_Call(\d)$");
        if (match.Success)
        {
          call = int.Parse(match.Groups[1].Value);
        }

        foreach (var item in mirror.VarList.Keys)
        {
          string varName = item;

          if (mirror.VarList[varName][1] == "in" && call == Convert.ToInt16(mirror.VarList[varName][2]))
          {
            VariableData variableData = varMap.GetNewVariableData(varName)!;
            variableData.InOut = "in";
            variableData.Call = Convert.ToInt16(mirror.VarList[varName][2]);
            variableData.Deep = deep;
            variableList.Add(variableData);

            // 🔄 Convert Source to ToolData when source is not in excluded list
            if (!sourcesExcluded.Contains(variableData.Source))
            {
              ToolData sourceTool = tm.getToolDataByCode(variableData.Source);
              newSources.Add(sourceTool);
            }

          }
        }
      }

      variableList = variableList.Distinct().ToList();

      // 🔍 Remove duplicates and excluded sources
      newSources = newSources
          .DistinctBy(t => t.Code) // Requires System.Linq
          .Where(t => !sourcesExcluded.Contains(t.Code))
          .ToList();

      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "GetRouteMap", $"newSources deep ({deep}): {string.Join(", ", newSources.Select(t => t.Code))}");

      List<string> varlist = variableList.Select(v => v.ID).ToList();
      H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "GetRouteMap", $"varlist: {string.Join(", ", varlist)}");

      variableList.AddRange(Get_PREP_Variables(newSources, sourcesExcluded, deep, calcTools));
      variableList = variableList.Distinct().ToList();

      return variableList;
    }
    public static List<ToolData> GetDeliveryDocs(XmlDocument xmlDoc, string language ="")
    {
      //If a list of deliveryDocs comes from Call , we use it. If not, we use all templates from ToolsMap
      ToolsMap tm = ToolsMap.Instance;
      string origen="";

      XmlNode deliveryDocsNode = xmlDoc.SelectSingleNode("/request/requestInfo/deliveryDocs")!;
      List<ToolData> deliveryDocs = [];

      if (deliveryDocsNode != null) foreach (XmlNode docNode in deliveryDocsNode.SelectNodes("doc")) deliveryDocs.Add( tm.getToolDataByCode(docNode.InnerText));

      if (deliveryDocs.Count > 0)
      {
        origen = "Call.xml";
      }
      else {
       // In case there are no specific delivery documents from call, we use all active templates from ToolsMap

        foreach (string doc in ToolsMap.Instance.DeliveryDocsPack) deliveryDocs.Add(tm.getToolDataByCode(doc, language));

        origen = "Default List";


      }

      var fileNames = (deliveryDocs ?? new List<ToolData>())
          .Where(d => d is not null && !string.IsNullOrEmpty(d.FileName))
          .Select(d => d.FileName);

      H.PrintLog(
          2,
          ThreadContext.CurrentThreadInfo.Value!.User,
          "GetDeliveryDocs",
          $"\n -- EXECUTING THE FOLLOWING DOCUMENTS (filenames): --\n Origen:({origen})" + Environment.NewLine +
          string.Join(Environment.NewLine, fileNames) +
          Environment.NewLine + Environment.NewLine
      );
      return deliveryDocs!;
    }
    public static string ReadFileContent(string filePath)
    {
      if (File.Exists(filePath))
      {
        return File.ReadAllText(filePath);
      }
      else
      {
        throw new FileNotFoundException($"The file: '{filePath}' does not exist.");
      }
    }
  }

}
