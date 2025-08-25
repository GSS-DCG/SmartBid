using System.Globalization;
using System.Xml;

namespace SmartBid
{
  public class DataMaster
  {
    private XmlDocument _dm;
    private VariablesMap _vm;
    private Dictionary<string, VariableData> _data;
    private XmlNode? _projectDataNode;
    private XmlNode? _utilsNode;
    private XmlNode? _dataNode;

    public string FileName { get; set; }
    public string User { get; set; } = ThreadContext.CurrentThreadInfo.Value!.User;
    public Dictionary<string, VariableData> Data { get { return _data; } }
    public XmlDocument DM { get { return _dm; } }
    public int BidRevision { get; set; }
    public string sBidRevision => BidRevision.ToString("D2");



    // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
    public DataMaster(XmlDocument xmlRequest)
    {
      _vm = VariablesMap.Instance;
      BidRevision = 1;
      _dm = new XmlDocument();
      _data = new Dictionary<string, VariableData>();

      // Check if opportunityFolder exists, otherwise throw an exception
      if (GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder") == null)
      {
        H.PrintLog(5, User, "CargaXML", "⚠️ Nodo 'opportunityFolder' no encontrado.");
        throw new InvalidOperationException("El XML está incompleto: falta '//requestInfo/opportunityFolder'.");
      }

      string opportunityFolder = GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder").InnerText;
      FileName = Path.Combine(H.GetSProperty("processPath"), opportunityFolder, $"{opportunityFolder.Substring(0, 7)}_DataMaster.xml");

      // register actual revision number in _data (no need to store it in DM)
      StoreValue("revision", new VariableData("revision", "current Revision", "utils", "utils", true, true, "code", "", "", "", "", 0, [], "rev_01"));

      if (((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")!).GetAttribute("Type") == "create")
      {
        XmlDeclaration xmlDeclaration = _dm.CreateXmlDeclaration("1.0", "utf-8", null);
        _ = _dm.AppendChild(xmlDeclaration);

        XmlElement root = _dm.CreateElement("dm");
        _ = _dm.AppendChild(root);

        //Creating DataMaster structure
        XmlElement init = _dm.CreateElement("projectData");
        _projectDataNode = root.AppendChild(init);

        XmlElement utils = _dm.CreateElement("utils");
        _utilsNode = root.AppendChild(utils);

        XmlElement data = _dm.CreateElement("data");
        _dataNode = root.AppendChild(data);


        List<VariableData> varList = VariablesMap.Instance.GetVarListBySource("INIT");

        XmlDocument configDataXML = new XmlDocument();
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
            //                        _ = _dataNode.AppendChild(GetImportedElement(xmlRequest, @$"//config/{variable.ID}"));

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
        StoreValue("requestTimestap", GetImportedElement(xmlRequest, "//requestInfo/requestTimestap").InnerText);

        //Add first revision element
        _ = _utilsNode.AppendChild(_dm.CreateComment("First Revision"));
        XmlElement revision = _dm.CreateElement("rev_01");

        _ = revision.AppendChild(CreateElement("dateTime", DateTime.Now.ToString("yyMMdd_HHmm")));

        XmlElement importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo");
        _ = importedNode != null ? revision.AppendChild(_dm.ImportNode(importedNode, true)) : null;

        importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo/deliveryDocs");
        _ = importedNode != null ? revision.AppendChild(_dm.ImportNode(importedNode, true)) : null;

        importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo/inputDocs");
        _ = importedNode != null ? revision.AppendChild(_dm.ImportNode(importedNode, true)) : null;

        _ = _utilsNode.AppendChild(revision);
      }

    }

    // Constructor público con nombre de archivo ==> para cargar un DataMaster existente
    public DataMaster(string dmFileName)
    {
      _vm = VariablesMap.Instance;
      // Implementación pendiente
    }

    public void UpdateData(XmlDocument newData)
    {
      XmlNode variablesNode = newData.SelectSingleNode("/root/variables");
      if (variablesNode == null) return;

      foreach (XmlNode variable in variablesNode.ChildNodes)
      {
        XmlNode importedNode = _dm.ImportNode(variable, true);

        XmlElement revisionElement = _dm.CreateElement("revision");
        XmlElement rev01Element = _dm.CreateElement($"rev{sBidRevision}");
        rev01Element.InnerText = $"set{sBidRevision}";
        _ = revisionElement.AppendChild(rev01Element);

        XmlElement setElment = _dm.CreateElement($"set{sBidRevision}");

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
      H.PrintLog(4, User, "DM", $"XML guardado en {FileName}");
    }

    public string GetValueString(string key)
    {
      if (_data.ContainsKey(key))
      {
        return _data[key]?.Value;
      }
      else
      {
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
      if (_data.ContainsKey(key))
      {
        string xmlString = _data[key]?.Value ?? string.Empty;

        if (string.IsNullOrWhiteSpace(xmlString))
          return null;

        XmlDocument tempDoc = new XmlDocument();
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

    public VariableData GetVariableData(string key)
    {
      if (_data.ContainsKey(key))
      {
        return _data[key];
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
      H.PrintLog(1, User, "StoreValue", $"variable ||{id}|| added to DataMaster data");
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
      XmlElement sourceElement = (XmlElement)sourceDoc.DocumentElement.SelectSingleNode(elementName);
      if (sourceElement == null)
      {
        throw new XmlException($"Element '{elementName}' not found in the source document.");
      }

      XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);
      return importedElement;
    }

    public void CheckMandatoryValues()
    {
      List<string> missingValues = new List<string>();

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
