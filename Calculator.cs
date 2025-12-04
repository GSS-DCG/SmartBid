using System.Diagnostics;
using System.Xml;

namespace SmartBid
{
  internal class Calculator
  {
    private List<ToolData> targets;
    private List<ToolData?> calcRoute;
    private List<string> statusList;

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

      List<string> route = (calcRoute ?? Enumerable.Empty<ToolData?>())
            .Select(tool => tool?.Code ?? "")
            .ToList();
      route[0] = "PREP";

      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"***  CALCULATE  *** ");
      H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Calculate rute: {string.Join(" >> ", route.Select(tool => tool))} ");



      int i = -1;
      bool printEndCalc = true;
      string status = "Ready";

      //create a new list status to store the status of the tools while there are run. it should be a list of the same size of route with "Ready" as all values

      ToolData tool;

      // remove all files from TOOLs and TEMPLATEs previous to calculation
      string outputFolder = Path.Combine(H.GetSProperty("processPath"),
                                              dm.GetInnerText($@"dm/utils/utilsData/opportunityFolder"),
                                              dm.SBidRevision,
                                              "TOOLS");
      DirectoryInfo di = new DirectoryInfo(Path.Combine(outputFolder, "TOOLS"));
      if (Directory.Exists(Path.Combine(outputFolder, "TOOLS")))
      {
        foreach (FileInfo file in di.GetFiles())
        {
          file.Delete();
        }
      }
      di = new DirectoryInfo(Path.Combine(outputFolder, "TEMPLATES"));
      if (Directory.Exists(Path.Combine(outputFolder, "TEMPLATES")))
      {

        foreach (FileInfo file in di.GetFiles())
        {
          file.Delete();
        }
      }

      //CALCULATE
      while (++i < calcRoute.Count)
      {

        statusList[i] = "RUNNING"; //Updating the status list
        DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, i, "RUNNING");

        if (i == 0) // PREP
        {
          //Buscamos los datos necesarios con la PreparationTool y los guardamos en el DataMaster

          XmlDocument prepAnswer = CallPrepTool(xmlPrepVarList);

          dm.UpdateData(prepAnswer);
          dm.CheckMandatoryValues(); // thows and exception if not all Mandatory values are present in the DataMaster

          statusList[0] = "DONE"; //Updating the status list for PREP as done.
          DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, 0, "DONE");

          continue;
        }

        tool = calcRoute[i];

        if (tm.Tools.Exists(tool => tool.Code == tool.Code))
        {

          if (tool.Resource == "TOOL")
          {
            H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Calling Tool: {tool.Code} - threadSafe: {tool.IsThreadSafe}");

            int callID = (int)TC.ID.Value.CallId!;
            //Call calculation

            if (!tool.IsThreadSafe)
            {
              DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, i, "WAITING");

              Stopwatch sw = Stopwatch.StartNew();
              int order, newOrder = 0;
              double timeout = 60 * 24;

              while (!tm.CheckForGreenLight(tool.Code, callID, out order))
              {
                if ((sw.Elapsed.TotalMinutes + (order * 60)) < timeout) timeout = sw.Elapsed.TotalMinutes + (order * 60); //Adjusting timeout to the position in the queue
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

            //Check if there is update in the route:
            //--find if there is a variable call __execute__ in CalcResults
            //--read the values from the variable and insert one by one in the route after the current tool
            //--remove the variable from CalcResults
            //--update the DB route progress with the new tools added
            XmlNode? executeNode = CalcResults.SelectSingleNode("/answer/variables/__execute__/value");
            if (executeNode != null)
            {
              string executeValue = executeNode.InnerText;
              List<string> newToolsCodes = executeValue.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                  .Select(code => code.Trim())
                  .ToList();
              if (newToolsCodes.Count > 0)
              {
                H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"New tools to execute after {tool.Code} found: {string.Join(", ", newToolsCodes)}");
                List<ToolData> newTools = new();
                List<string> newStatus = new();
                foreach (string code in newToolsCodes)
                {
                  ToolData newTool = tm.getToolDataByCode(code);
                  newTools.Add(newTool);
                  newStatus.Add("Ready");
                }
                //Insert the new tools in the route after the current tool
                calcRoute.InsertRange(i + 1, newTools);
                statusList.InsertRange(i + 1, newStatus);
                //Update the DB route progress with the new tools added

                route = (calcRoute ?? Enumerable.Empty<ToolData?>())
                .Select(tool => tool?.Code ?? "")
                .ToList();
                route[0] = "PREP";

                DBtools.CreateRouteProgress(TC.ID.Value!.CallId!.Value, route, statusList);
                H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Updated calculation route: {string.Join(" >> ", route)}");
              }
              //Remove the __execute__ variable from CalcResults

              var node = CalcResults.SelectSingleNode("/answer/variables/__execute__");
              if (node != null && node.ParentNode != null)
              {
                _ = node.ParentNode.RemoveChild(node);
              }
            }

            //Once the tool is finished, we release the lock if it is not threadSafe
            if (!tool.IsThreadSafe)
              tm.ReleaseProcess(tool.Code, callID);

            //Updating the deliveryDocs information to the DataMaster and removing the auxiliar variable form CalcResults

            // Find if there is a variable called __deliveryDocs__ in CalcResults
            XmlNodeList? deliveryDocsNodes = CalcResults.SelectNodes("/answer/variables/__deliveryDocs__/value/l/li");


            if (deliveryDocsNodes.Count > 0 )
            {
              // Iterate through each <li> element
              foreach (XmlNode liNode in deliveryDocsNodes)
              {
                string filePath = liNode.InnerText.Trim();



                if (!string.IsNullOrEmpty(filePath))
                {
                  // Insert each filePath into DataMaster using updateDeliveryDoc(filePath, "")
                  dm.updateDeliveryDoc(filePath, tool.Code);
                }
              }
              var node = CalcResults.SelectSingleNode("/answer/variables/__deliveryDocs__");
              if (node != null && node.ParentNode != null)
              {
                _ = node.ParentNode.RemoveChild(node);
              }
            }

            //Update the DataMaster with the results from the toolD
            dm.UpdateData(CalcResults);

            // Update the DB with the result from the tool calculation
            try
            {
              // Get the <answer> node and read its "result" attribute
              var answerNode = CalcResults.DocumentElement; // expects <answer ...> as the root
              var resultRaw = answerNode?.GetAttribute("result") ?? string.Empty;

              status = answerNode?.GetAttribute("result")?.Equals("OK", StringComparison.OrdinalIgnoreCase) == true
                  ? "DONE"
                  : "FAIL";
            }
            catch (Exception ex)
            {
              H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User,
                  "RunCalculations",
                  "❌❌ Error ❌❌  reading result from XML or updating DB: " + ex.Message);
            }
          }

          else if (tool.Resource == "TEMPLATE")
          {       //GENERATE DOCUMENTS

            if (printEndCalc)
            {
              printEndCalc = false;
              H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"- Calculate Done - ");
              H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"***   GENERATE DOCUMENTS   *** ");
            }

            H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "RunCalculations", $"Populating Template: {tool.Code} - {tool.Description} ");

            tm.GenerateOuput(tool, dm);
            //
            if (true)
            {  //Here we could check if the generation was OK or not
              statusList[i] = "DONE"; //Updating the status list

              dm.updateDeliveryDoc(tool.FileName, "GenerateDocs", tool.Code, tool.Version);

              DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, i, "DONE");
              // GENERATE INFO ABOUT TEMPLATES GENERATED (PENDIENTE)
              // dm.UpdateData(NEW_INFO); //Update the DataMaster with the information about generated documents
            }
          }
          else
          {
            H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌  - RunCalculations", $"Tool {tool} not found in ToolsMap. ");
          }

          status = "DONE";
        }
        // Report to DB
        statusList[i] = status; //Updating the status list
        DBtools.UpdateRouteProgress(TC.ID.Value!.CallId!.Value, i, status);
      }


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
      _ = prepCallXML.AppendChild(xmlDeclaration);
      XmlElement root = prepCallXML.CreateElement("call");
      _ = prepCallXML.AppendChild(root);

      prepVarList.AddRange(Get_PREP_Variables(targets, sourcesSearched, 0, calcTools));

      // Build up the list of calls to tools in order
      calcRoute = [];
      HashSet<ToolData> uniqueElements = new(calcRoute);

      calcRoute.Add(null); //Adding a empty element to represent PREP

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

      _ = prepVarList.RemoveAll(item => !(item.Source == "PREP")); //Remove all non-preparation source variables

      List<string> prepVarIDList = prepVarList.Select(variable => variable.ID).ToList(); //Getting the list of IDs for the preparation variables

      XmlElement variablesXml = VariablesMap.Instance.ToXml(prepCallXML, prepVarIDList);
      _ = root.AppendChild(variablesXml); //Adding the variables element to the XML


      XmlNode? dmInputDocs = dm.DM.SelectSingleNode(@$"dm/utils/rev_{dm.BidRevision.ToString("D2")}/inputDocs");
      if (dmInputDocs != null)
        _ = prepCallXML.DocumentElement.AppendChild(prepCallXML.ImportNode(dmInputDocs, true));


      List<string> route =
          (calcRoute ?? Enumerable.Empty<ToolData?>())
          .Select(tool => tool?.Code ?? string.Empty)
          .ToList();

      route[0] = "PREP";


      statusList = route.Select(_ => "Ready").ToList();

      DBtools.CreateRouteProgress(TC.ID.Value!.CallId!.Value, route, statusList);

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