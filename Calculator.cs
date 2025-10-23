﻿using System.Collections.Immutable;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Microsoft.Office.Interop.Word;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace SmartBid
{
  internal class Calculator
  {
    private List<ToolData> targets;
    private List<ToolData> calcRoute = [];
    public ToolsMap? tm;
    public DataMaster dm;

    public Calculator(DataMaster dataMaster, List<ToolData> targets)
    {
      dm = dataMaster; // Asigna la instancia de DataMaster pasada
      string opportunityFolder = dm.GetValueString("opportunityFolder");
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "Calculator", $"REQUEST: ");

      this.targets = targets;
    }
    public void RunCalculations()
    {
      tm = ToolsMap.Instance;

      //Find the list of variable to get from Preparation Tool
      string xmlPrepVarList = GetRouteMap(targets).OuterXml;

      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"***  CALCULATE  *** ");
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Calculate rute: {string.Join(" >> ", calcRoute.Select(tool => tool.Code))} ");


      //Buscamos los datos necesarios con la PreparationTool y los guardamos en el DataMaster
      DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, "PREP", "Running");
      dm.UpdateData(CallPrepTool(xmlPrepVarList));
      dm.CheckMandatoryValues(); // thows and exception if not all Mandatory values are present in the DataMaster
      dm.SaveDataMaster(); //Save the DataMaster after preparation
      DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, "PREP", "DONE");

      //Generate files structure and move input files
      //Call each toolD in the list of targets and update the DataMaster with the results

      //CALCULATE
      foreach (ToolData tool in calcRoute)
      {
        // por el momento nos saltamos PREP porque la ejecutamos manualmente antes de entrar en calculations
        if (tool.Code == "PREP")
          break;

        if (tm.Tools.Exists(tool => tool.Code == tool.Code))
        {

          if (tool.Resource == "TOOL")
          {
            H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Calling Tool: {tool.Code} - threadSafe: {tool.IsThreadSafe}");

            DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, tool.Code, "Running");

            int callID = (int)TC.ID.Value.CallId!;
            //Call calculation

            if (!tool.IsThreadSafe)
            {

              Stopwatch sw = Stopwatch.StartNew();
              int order, newOrder = 0;
              double timeout = 60*24;

              while (!tm.CheckForGreenLight(tool.Code, callID, out order))
              {
                if ((sw.Elapsed.TotalMinutes + order * 60) < timeout) timeout = sw.Elapsed.TotalMinutes + order * 60; //Adjusting timeout to the position in the queue
                if (sw.Elapsed.TotalMinutes > timeout)
                {
                  throw new TimeoutException($"Timeout para {tool.Code} after {sw.Elapsed.TotalMinutes:F1} minutes");
                }
                //printing only order change with log 5, every 30 seconds with log 2
                if (order == newOrder)
                  H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Esperando turno: {tool.Code} por {sw.Elapsed.TotalMinutes:F2} minutos. (Puesto en cola: {order} Timeout set to: {timeout:F0} minutes)");
                else
                {
                  newOrder = order;
                  H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Esperando turno: {tool.Code} por {sw.Elapsed.TotalMinutes:F2} minutos. (Puesto en cola: {order} Timeout set to: {timeout:F0} minutes)");
                }

                Thread.Sleep(30000);
              }

              H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Turno: {tool.Code} Liberado    ....calculando");
            }

            XmlDocument CalcResults = tm.Calculate(tool, dm);

            //Once the tool is finished, we release the lock if it is not threadSafe
            if (!tool.IsThreadSafe)
              tm.ReleaseProcess(tool.Code, callID);

            //Update the DataMaster with the results from the toolD
            dm.UpdateData(CalcResults);

            //Save the DataMaster after each calculation
            dm.SaveDataMaster(); //Save the DataMaster after each calculation

            // Update the DB with the result from the tool calculation
            try
            {
              // Get the <answer> node and read its "result" attribute
              var answerNode = CalcResults.DocumentElement; // expects <answer ...> as the root
              var resultRaw = answerNode?.GetAttribute("result") ?? string.Empty;

              string status = answerNode?.GetAttribute("result")?.Equals("OK", StringComparison.OrdinalIgnoreCase) == true
                  ? "DONE"
                  : "FAIL";

              // Report to DB
              DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, tool.Code, status);
            }
            catch (Exception ex)
            {
              H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User,
                  "RunCalculations",
                  "❌❌ Error ❌❌  reading result from XML or updating DB: " + ex.Message);
            }
          }
        }
        else
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌  - RunCalculations", $"Tool {tool} not found in ToolsMap. ");
        }
      }
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"- Calculate Done - ");
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"***   GENERATE DOCUMENTS   *** ");
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Generation rute: {string.Join(" >> ", targets.Select(tool => tool.Code))} ");

      //GENERATE DOCUMENTS
      foreach (ToolData target in targets)
      {
        if (target.Resource == "TEMPLATE")
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Populating Template: {target.Code} - {target.Description} ");

          DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, target.Code, "Running");

          tm.GenerateOuput(target, dm);
          //
          if (true) //Here we could check if the generation was OK or not
            DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, target.Code, "DONE");


          // GENERATE INFO ABOUT TEMPLATES GENERATED (PENDIENTE)
          // dm.UpdateData(NEW_INFO); //Update the DataMaster with the information about generated documents
        }
      }

      _ = DBtools.InsertNewProjectWithBid(dm);

    }
    private XmlDocument CallPrepTool(string xmlVarListString)
    {
      XmlDocument prepCall = new();
      prepCall.LoadXml(xmlVarListString); // Load the XML string into the XmlDocument

      H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", $"\n");
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", $"***   PREPARATION   *** ");
      H.PrintLog(0, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", "- Argumento pasado a PREP:", prepCall);
      H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", $"\n\n");
      H.SaveXML(3, prepCall, Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "prepCall.xml")); // Save the XML to a file for debugging


      PREP prepInstance = new PREP(); // Crear una NUEVA instancia de PREP para este hilo
      XmlDocument xmlPrepAnswer = prepInstance.Run(prepCall); // Llamar al método de instancia

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", $"Return from Preparation ", xmlPrepAnswer);
      H.SaveXML(1, xmlPrepAnswer, Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "PrepAnswer.xml")); // Save the XML to a file for debugging

      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "CallPrepTool", "- Preparation Done - ");

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
      calcRoute = [];
      HashSet<ToolData> uniqueElements = new(calcRoute);
      for (int i = calcTools.Count - 1; i >= 0; i--) // Iterate backwards through calcTools
      {
        foreach (var element in calcTools[i])
        {
          if (uniqueElements.Add(element)) // Add only if it's not already present
          {
            calcRoute.Add(element);
          }
        }
      }

      targets = calcRoute; // Update targets with the ordered list of calculation tools

      _ = prepVarList.RemoveAll(item => !(item.Source == "PREP")); //Remove all non-preparation source variables

      List<string> prepVarIDList = prepVarList.Select(variable => variable.ID).ToList(); //Getting the list of IDs for the preparation variables

      XmlElement variablesXml = VariablesMap.Instance.ToXml(prepCallXML, prepVarIDList);
      _ = root.AppendChild(variablesXml); //Adding the variables element to the XML


      XmlNode? dmInputDocs = dm.DM.SelectSingleNode(@$"dm/utils/rev_{dm.BidRevision.ToString("D2")}/inputDocs");
      if (dmInputDocs != null)
        prepCallXML.DocumentElement.AppendChild(prepCallXML.ImportNode(dmInputDocs, true));

      List<string> route = calcRoute.Select(tool => tool.Code).ToList();
      route.Insert(0, "PREP");


      DBtools.CreateRouteProgress(TC.ID.Value!.CallId!.Value, route);

      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetRouteMap", $"- Calculated Route Map: \n {string.Join(" >> ", route)} ");

      // Guardamos el XML por si fuese necesario

      prepCallXML.Save(Path.Combine(Path.GetDirectoryName(H.GetSProperty("ToolsPath"))!, "preparationCall.xml"));
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

        foreach (var item in mirror.VarList.Keys)
        {
          string varName = item;

          if (mirror.VarList[varName][1] == "in" && call == Convert.ToInt16(mirror.VarList[varName][2]))
          {
            VariableData variableData = varMap.GetNewVariableData(varName)!;
            variableData.InOut = "in";
            //variableData.Call = Convert.ToInt16(mirror.VarList[varName][2]);
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

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetRouteMap", $"newSources deep ({deep}): {string.Join(", ", newSources.Select(t => t.Code))}");

      List<string> varlist = variableList.Select(v => v.ID).ToList();
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GetRouteMap", $"varlist: {string.Join(", ", varlist)}");

      variableList.AddRange(Get_PREP_Variables(newSources, sourcesExcluded, deep, calcTools));
      variableList = variableList.Distinct().ToList();

      return variableList;
    }
    public static List<ToolData> GetDeliveryDocs(XmlDocument xmlDoc, string language = "")
    {
      //If a list of deliveryDocs comes from Call , we use it. If not, we use all templates from ToolsMap
      ToolsMap tm = ToolsMap.Instance;
      string origen = "";

      XmlNode deliveryDocsNode = xmlDoc.SelectSingleNode("/request/requestInfo/deliveryDocs")!;
      List<ToolData> deliveryDocs = [];

      if (deliveryDocsNode != null) foreach (XmlNode docNode in deliveryDocsNode.SelectNodes("doc")) deliveryDocs.Add(tm.getToolDataByCode(docNode.InnerText));

      if (deliveryDocs.Count > 0)
      {
        origen = "Call.xml";
      }
      else
      {
        // In case there are no specific delivery documents from call, we use all active templates from ToolsMap

        foreach (string doc in ToolsMap.Instance.DeliveryDocsPack) deliveryDocs.Add(tm.getToolDataByCode(doc, language));

        origen = "Default List";


      }

      var fileNames = (deliveryDocs ?? new List<ToolData>())
          .Where(d => d is not null && !string.IsNullOrEmpty(d.FileName))
          .Select(d => d.FileName);

      H.PrintLog(
          2,
          TC.ID.Value!.Time(), TC.ID.Value!.User,
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