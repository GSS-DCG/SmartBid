using System.Collections.ObjectModel;
using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.EMMA;
using Microsoft.Office.Interop.Word;
using Path = System.IO.Path;

namespace SmartBid
{
  public class DataMaster
  {
    private readonly XmlDocument _dm;
    private readonly VariablesMap _vm;
    private readonly Dictionary<string, VariableData> _data;
    private readonly XmlNode _projectDataNode;
    private readonly XmlNode _utilsNode;
    private readonly XmlNode _dataNode;

    public string FileName { get; set; }
    public string User { get; set; } = ThreadContext.CurrentThreadInfo.Value!.User;
    public Dictionary<string, VariableData> Data { get { return _data; } }
    public XmlDocument DM { get { return _dm; } }
    public int BidRevision { get; set; }
    public string SBidRevision => BidRevision.ToString("D2");



    // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
    public DataMaster(XmlDocument xmlRequest, List<ToolData> targets)
    {
      _vm = VariablesMap.Instance;
      BidRevision = 1;
      _dm = new();
      _data = [];

      H.PrintLog(5, User, "DataMaster constructor", "Create DataMaster using call:", xmlRequest);


      // Check if opportunityFolder exists, otherwise throw an exception
      if (GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder") == null)
      {
        H.PrintLog(5, User, "CargaXML", "⚠️ Nodo 'opportunityFolder' no encontrado.");
        throw new InvalidOperationException("El XML está incompleto: falta '//requestInfo/opportunityFolder'.");
      }

      string opportunityFolder = GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder").InnerText;
      FileName = Path.Combine(H.GetSProperty("processPath"), opportunityFolder, $"{opportunityFolder[..7]}_DataMaster.xml");

      // register actual revision number in _data (no need to store it in DM)
      StoreValue("revision", new VariableData("revision", "current Revision", "utils", "utils", true, true, "code", "", "", "", "", 0, [], "rev_01"));

      if (((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")!).GetAttribute("type") == "create")
      {
        XmlDeclaration xmlDeclaration = _dm.CreateXmlDeclaration("1.0", "utf-8", null);
        _ = _dm.AppendChild(xmlDeclaration);

        XmlElement root = _dm.CreateElement("dm");
        _ = _dm.AppendChild(root);

        //Creating DataMaster structure
        XmlElement init = _dm.CreateElement("projectData");
        _projectDataNode = root.AppendChild(init)!;

        XmlElement utils = _dm.CreateElement("utils");
        _utilsNode = root.AppendChild(utils)!;

        XmlElement data = _dm.CreateElement("data");
        _dataNode = root.AppendChild(data)!;


        List<VariableData> varList = VariablesMap.Instance.GetVarListBySource("INIT");

        XmlDocument configDataXML = new();
        XmlElement configDataRoot = configDataXML.CreateElement("root");
        _ = configDataXML.AppendChild(configDataRoot);
        XmlElement variables = configDataXML.CreateElement("variables");
        _ = configDataRoot.AppendChild(variables);

        //Saving INIT variables in DataMaster
        foreach (VariableData variable in varList)
        {
          // Load PROJECTDATA Data
          if (variable.Area == "projectData")
          {
            XmlNode element = GetImportedElement(xmlRequest, @$"//projectData/{variable.ID}");
            _ = _projectDataNode.AppendChild(element);

            StoreValue(variable.ID, element.InnerText);//Extraer el valor para almacenar en _data.
          }

          // Load CONFIG Init Data
          // (creating an xml with config data variables and send it to UpdateData for inserting in DM)
          if (variable.Area == "config")
          {
            //tomamos el nombre de la variable
            //                        _ = _dataNode.AppendChild(GetImportedElement(xmlRequest, @$"//config/{variable.Code}"));

            // Variable 
            XmlElement newVar = configDataXML.CreateElement(variable.ID);
            _ = variables.AppendChild(newVar);

            // Value node

            XmlElement importedElement = GetImportedElement(xmlRequest, @$"//config/{variable.ID}");
            XmlAttribute unitAttribute = importedElement?.GetAttributeNode("unit")!;


            XmlNode value = newVar.AppendChild(CreateElement(configDataXML, "value", importedElement?.InnerText ?? ""));
            if (unitAttribute != null)
              ((XmlElement)value!).SetAttribute("unit", unitAttribute.Value);
            _ = newVar.AppendChild(CreateElement(configDataXML, "origin", "INIT from callXML"));
            _ = newVar.AppendChild(CreateElement(configDataXML, "note", "Variable leida de Hermes"));
          }
        }
        UpdateData(configDataXML);

        // Load Utils Data
        XmlNode utilsData = _utilsNode!.AppendChild(DM.CreateElement("utilsData"))!;

        // Add opportunityFolder to dataMaster and _data dictionary
        _ = utilsData.AppendChild(CreateElement("dataMasterFileName", FileName));
        _ = utilsData.AppendChild(GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder"));
        StoreValue("opportunityFolder", GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder").InnerText);
        StoreValue("createdBy", GetImportedElement(xmlRequest, "//requestInfo/createdBy").InnerText);
        StoreValue("requestTimestamp", GetImportedElement(xmlRequest, "//requestInfo/requestTimestamp").InnerText);

        XmlElement revision = CreateRevisionElement(1, xmlRequest, targets, opportunityFolder);

        _utilsNode.AppendChild(revision);
      }
      else if(((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")!).GetAttribute("Type") == "modify")
      {
      // Pending implementation for new revision
      }
      else
      {
        H.PrintLog(5, User, "CargaXML", "⚠️ Atributo 'Type' en '//request/requestInfo' non valid or not found.");
        throw new InvalidOperationException("The XML is incomplete: missing or invalid 'Type' attribute in '//request/requestInfo'.");
      }

    }

    private XmlElement CreateRevisionElement(int revisionNo, XmlDocument xmlRequest, List<ToolData> targets, string opportunityFolder)
    {
      //Add first revision element
      _ = _utilsNode.AppendChild(_dm.CreateComment("First Revision"));
      XmlElement revision = _dm.CreateElement($"rev_{revisionNo!:D2}");

      _ = revision.AppendChild(CreateElement("dateTime", DateTime.Now.ToString("yyMMdd_HHmm")));
      _ = revision.AppendChild(CreateElement("outputFolder", Path.Combine(  H.GetSProperty("processPath"),
                                                                            GetValueString("opportunityFolder"),
                                                                            $"rev_{SBidRevision}",
                                                                            "OUTPUT")));

      //Adding Request Info from Call
      XmlElement newNode;
      newNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo")!;
      _ = newNode != null ? revision.AppendChild(_dm.ImportNode(newNode, true)) : null;

      // Adding Processed InputDocs
      // checks that all files exits and stores the checksum of the fileName for comparison
      revision.AppendChild(ProcessInputFiles((XmlElement)xmlRequest.SelectSingleNode("//requestInfo/inputDocs")!));

      //Adding deliveryDocs from either Call or Default Delivery Docs
      newNode = H.CreateElement(_dm, "deliveryDocs", "");
      //add each one of the fileName in targets as a new element called "doc" to newNode
      foreach (ToolData target in targets)
      {
        XmlElement newChild = CreateElement("doc", target.FileName);
        newChild.SetAttribute("code", target.Code);
        newChild.SetAttribute("version", target.Version);
        newNode.AppendChild(newChild);
      }
      _ = newNode != null ? revision.AppendChild(_dm.ImportNode(newNode, true)) : null;



      //Adding node to store tools used in this revision
      newNode = H.CreateElement(_dm, "tools", "");
      newNode.SetAttribute("processedFolder", Path.Combine(H.GetSProperty("processPath"), $@"{opportunityFolder}\TOOLS\rev_{SBidRevision}"));
      _ = newNode != null ? revision.AppendChild(_dm.ImportNode(newNode, true)) : null;
      return revision;
    }

    // Constructor público con nombre de archivo ==> para cargar un DataMaster existente
    public DataMaster(string dmFileName)
    {
      _vm = VariablesMap.Instance;
      // Implementación pendiente
    }

    private XmlElement ProcessInputFiles(XmlElement inputDocs)
    {
      XmlElement newInputDocs = H.CreateElement(_dm, "inputDocs", "");
      foreach (XmlElement doc in inputDocs)
      {
        string fileType = doc.GetAttribute("type");
        string filePath = doc.InnerText;
        string fileName = Path.GetFileName(filePath);

        if (!File.Exists(filePath))
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, $"❌❌ Error ❌❌ - ProcessFile", $"⚠️ File: '{filePath}' is not found.");
          continue; // Saltar este documento y seguir con los demás
        }

        string hash = H.GetFileMD5(filePath); // Calculate MD5 hash for the fileName
        string lastModified = File.GetLastWriteTime(filePath).ToString("yyyy-MM-dd HH:mm:ss");

        newInputDocs.AppendChild(H.CreateElement(_dm, fileType, filePath, new Dictionary<string, string>
        {
          { "hash", hash },
          { "lastModified", lastModified }
        }));

        DBtools.InsertFileHash(filePath, fileType, hash, lastModified); // Store the fileName hash in the database 

        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"Archivo '{filePath}' registered");
      }

      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"All input files have been registered'.");

      return newInputDocs;
    }


    public void UpdateData(XmlDocument newData)
    {
      //storing the tool node from /root/utils in newData to the node dm/utils/rev_01/tool in _dm
      //where the number of the revision can be found at  _dm.BidRevision, the name of the revision node 
      //should be created as rev_XX where XX is the number with two digits


      H.MergeXmlNodes(newData, _dm, "/*/utils", $"/dm/utils/rev_{SBidRevision}");

      XmlNode variablesNode = newData.SelectSingleNode("//*/variables")!;
      if (variablesNode == null) return;

      foreach (XmlNode variable in variablesNode.ChildNodes)
      {
        XmlNode importedNode = _dm.ImportNode(variable, true);

        XmlElement revisionElement = _dm.CreateElement("revision");
        XmlElement rev01Element = _dm.CreateElement($"rev{SBidRevision}");
        rev01Element.InnerText = $"set{SBidRevision}";
        _ = revisionElement.AppendChild(rev01Element);

        XmlElement setElment = _dm.CreateElement($"set{SBidRevision}");

        foreach (XmlNode child in importedNode.ChildNodes)
        {
          _ = setElment.AppendChild(child.CloneNode(true));
          if (child.Name == "value")
          {
            StoreValue(variable.Name, child.InnerXml);
          }
        }

        importedNode.RemoveAll();
        _ = importedNode.AppendChild(revisionElement);
        _ = importedNode.AppendChild(setElment);

        _ = _dataNode.AppendChild(importedNode);
      }
    }
    public void SaveDataMaster()
    {
      _dm.Save(FileName);
      H.PrintLog(3, User, "DM", $"XML guardado en {FileName}");
    }
    public string GetValueString(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value!.Value;
      }
      else
      {
        if (_vm.GetVariableData(key).Source == "UTILS")
        {
          key = key.Replace('.', '/');
          int dashIndex = key.IndexOf('-');

          if (dashIndex == -1)
          {
            return _dm.SelectSingleNode($"/dm/utils/rev_{SBidRevision}/{key}")?.InnerText ?? string.Empty;
          }
          else
          {
            string attribute = key[(dashIndex + 1)..];
            key = key[..dashIndex];
            string node = $"/dm/utils/rev_{SBidRevision}/{key}";

            return _dm.SelectSingleNode(node)?.Attributes?[attribute]?.Value ?? string.Empty;
          }
        }

        {
          return string.Empty; // return empty string if the key exists in VariablesMap but has not been set in DataMaster
        }
        H.PrintLog(5, User, $"❌❌ Error ❌❌ - DM.GetValueString ", $"Key '{key}' not found in DataMaster.");
        throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
      }
    }
    public double? GetValueNumber(string key)
    {
      if (_data.ContainsKey(key))
      {
        return double.TryParse(_data[key]?.Value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double num) ? num : null;
      }
      else
      {
        H.PrintLog(5, User, $"❌❌ Error ❌❌  - DM.GetValueNumber", $"Key '{key}' not found in DataMaster.");
        throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
      }
    }
    public bool? GetValueBoolean(string key)
    {
      if (_data.ContainsKey(key))
      {
        return bool.TryParse(_data[key]?.Value.ToString(), out bool num) ? num : null;
      }
      else
      {
        H.PrintLog(5, User, $"❌❌ Error ❌❌  - DM.GetValueBoolean", $"Key '{key}' not found in DataMaster.");
        throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
      }
    }
    public XmlNode GetValueXmlNode(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        string xmlString = value?.Value ?? string.Empty;

        if (string.IsNullOrWhiteSpace(xmlString))
          return null;

        XmlDocument tempDoc = new();
        try
        {
          tempDoc.LoadXml($"<root>{xmlString}</root>");
          return tempDoc.DocumentElement.FirstChild;
        }
        catch
        {
          return null; // returns NULL when the Value has not XML format
        }
      }
      else
      {
        H.PrintLog(5, User, $"❌❌ Error ❌❌  - DM", $"Key '{key}' not found in DataMaster.");
        throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
      }
    }
    public string GetValueUnit(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value!.Unit;
      }else {
        return string.Empty; // return empty string if the key exists in VariablesMap but has not been set in DataMaster
      }
      H.PrintLog(5, User, $"❌❌ Error ❌❌ - DM.GetValueUnit ", $"Key '{key}' not found in DataMaster.");
      throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
    }
    public VariableData GetVariableData(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value;
      }
      else
      {
        H.PrintLog(5, User, $"❌❌ Error ❌❌  - DM.GetVariableData", $"Key '{key}' not found in DataMaster.");
        throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
      }
    }
    public string GetInnerText(string xpath)
    {
      XmlNode node = _dm.SelectSingleNode(xpath);
      if (node != null)
      {
        return node.InnerText;
      }
      else
      {
        throw new XmlException($"Node not found for XPath: {xpath}");
      }
    }
    private void StoreValue(string id, string value)
    {
      H.PrintLog(1, User, "StoreValue", $"variable ||{id}: {value}|| added to DataMaster data");
      VariableData varData = _vm.GetVariableData(id);
      varData.Value = value;
      _data.Add(id, varData);
    }
    private void StoreValue(string id, VariableData varData)
    {
      H.PrintLog(1, User, "StoreValue", $"variable ||{id}|| added to DataMaster data");
      _data.Add(id, varData);
    }
    private XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
    {
      XmlElement? 
        sourceElement = (XmlElement)sourceDoc.DocumentElement!.SelectSingleNode(elementName)! ??
        throw new XmlException($"Element '{elementName}' not found in the source document.");
      XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);
      return importedElement;
    }
    public void CheckMandatoryValues()
    {
      List<string> missingValues = [];

      foreach (var kvp in _data)
      {
        string variableId = kvp.Key;
        string variableValue = GetValueString(variableId);

        if (_data[kvp.Key].Mandatory && string.IsNullOrWhiteSpace(_data[kvp.Key].Value))
        {
          missingValues.Add(variableId);
        }

      }

      if (missingValues.Count > 0) 
      {
        H.PrintLog(5, User, "CheckMandatoryValues", $"❌Error❌: Mandatory values not found in DataMaster. Cannot continue with calculations. Faltan: {string.Join(", ", missingValues)}");
        throw new InvalidOperationException("MandatoryValues missing");
      }
    }
    private XmlElement CreateElement(string name, string value)
    {
      XmlElement element = _dm.CreateElement(name);
      element.InnerText = value;
      return element;
    }
    private static XmlElement CreateElement(XmlDocument doc, string name, string value)
    {
      XmlElement element = doc.CreateElement(name);
      element.InnerText = value;
      return element;
    }
  }
}
