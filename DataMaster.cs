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

    // AÑADE ESTA PROPIEDAD
    public Guid InstanceId { get; }

    public string FileName { get; set; }
    public string User { get; set; } = TC.ID.Value!.User;
    public Dictionary<string, VariableData> Data { get { return _data; } } // Propiedad pública para el diccionario
    public XmlDocument DM { get { return _dm; } }
    public int BidRevision { get; set; }
    public string SBidRevision => $"rev_{BidRevision.ToString("D2")}";

    // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
    public DataMaster(XmlDocument xmlRequest, List<ToolData> targets)
    {
      // INICIALIZA InstanceId AQUÍ
      InstanceId = Guid.NewGuid();

      _vm = VariablesMap.Instance;
      BidRevision = 1;
      _dm = new();
      _data = new Dictionary<string, VariableData>();

      H.PrintLog(5, TC.ID.Value!.Time(), User, "DataMaster constructor", $"Create   using call:", xmlRequest);


      if (GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder") == null)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), User, "CargaXML", $"⚠️ Nodo 'opportunityFolder' no encontrado. ");
        throw new InvalidOperationException("El XML está incompleto: falta '//requestInfo/opportunityFolder'.");
      }

      string initialOpportunityFolder = GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder").InnerText; // Capturar temprano
      FileName = Path.Combine(H.GetSProperty("processPath"), initialOpportunityFolder, $"{initialOpportunityFolder[..7]}_DataMaster.xml");

      StoreValue("revision", new VariableData("revision", "current Revision", "utils", "utils", true, true, "code", "", "", "", "", 0, new List<string>(), SBidRevision));

      if (((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")!).GetAttribute("type") == "create")
      {
        XmlDeclaration xmlDeclaration = _dm.CreateXmlDeclaration("1.0", "utf-8", null);
        _ = _dm.AppendChild(xmlDeclaration);

        XmlElement root = _dm.CreateElement("dm");
        _ = _dm.AppendChild(root);

        XmlElement init = _dm.CreateElement("projectData");
        _projectDataNode = root.AppendChild(init)!;

        XmlElement utils = _dm.CreateElement("utils");
        _utilsNode = root.AppendChild(utils)!;

        XmlElement data = _dm.CreateElement("data");
        _dataNode = root.AppendChild(data)!;

        // --- OPERACIONES MOVIDAS HACIA ARRIBA EN EL CONSTRUCTOR ---
        // Poblar utilsData y StoreValue para "opportunityFolder" antes de la primera UpdateData
        XmlNode utilsData = _utilsNode!.AppendChild(DM.CreateElement("utilsData"))!; // _utilsNode ya debe estar inicializado
        _ = utilsData.AppendChild(CreateDmElement("dataMasterFileName", FileName));
        _ = utilsData.AppendChild(GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder"));
        StoreValue("opportunityFolder", initialOpportunityFolder); // Usar el valor capturado
        StoreValue("createdBy", GetImportedElement(xmlRequest, "//requestInfo/createdBy").InnerText);
        StoreValue("requestTimestamp", GetImportedElement(xmlRequest, "//requestInfo/requestTimestamp").InnerText);
        // --- FIN OPERACIONES MOVIDAS ---


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
          if (variable.Area == "config")
          {
            XmlElement newVar = configDataXML.CreateElement(variable.ID);
            _ = variables.AppendChild(newVar);

            XmlElement importedElement = GetImportedElement(xmlRequest, @$"//config/{variable.ID}");
            XmlAttribute unitAttribute = importedElement?.GetAttributeNode("unit")!;


            XmlNode value = newVar.AppendChild(CreateElement(configDataXML, "value", importedElement?.InnerText ?? ""));
            if (unitAttribute != null)
              ((XmlElement)value!).SetAttribute("unit", unitAttribute.Value);
            _ = newVar.AppendChild(CreateElement(configDataXML, "origin", "INIT from callXML"));
            _ = newVar.AppendChild(CreateElement(configDataXML, "note", "Variable leída de Hermes"));
          }
        }
        UpdateData(configDataXML); // Ahora, cuando UpdateData llame a GetValueString, "opportunityFolder" debería estar disponible.

        XmlElement revision = CreateRevisionElement(1, xmlRequest, targets, initialOpportunityFolder); // Pasar initialOpportunityFolder
        _utilsNode.AppendChild(revision);
      }
      else if (((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")!).GetAttribute("Type") == "modify")
      {
        // Pending implementation for new revision
      }
      else
      {
        H.PrintLog(5, TC.ID.Value!.Time(), User, "CargaXML", $"⚠️ Atributo 'Type' en '//request/requestInfo' non valid or not found. ");
        throw new InvalidOperationException("The XML is incomplete: missing or invalid 'Type' attribute in '//request/requestInfo'.");
      }

    }

    // Constructor público con nombre de archivo ==> para cargar un DataMaster existente
    public DataMaster(string dmFileName)
    {
      InstanceId = Guid.NewGuid();
      _vm = VariablesMap.Instance;
      _dm = new();
      _data = new Dictionary<string, VariableData>(); // ASEGURAR QUE SE INICIALIZA AQUÍ
      H.PrintLog(5, TC.ID.Value!.Time(), User, "DataMaster constructor", $"Create DataMaster (from file) ID: {this.InstanceId}  using file: {dmFileName}");
      // Implementación pendiente
    }

    // private Methods
    private static XmlElement CreateElement(XmlDocument doc, string name, string value)
    {
      XmlElement element = doc.CreateElement(name);
      element.InnerText = value;
      return element;
    }
    private XmlElement CreateRevisionElement(int revisionNo, XmlDocument xmlRequest, List<ToolData> targets, string opportunityFolder)
    {
      //Add first revision element
      _ = _utilsNode.AppendChild(_dm.CreateComment("First Revision"));
      XmlElement revision = _dm.CreateElement(SBidRevision);

      _ = revision.AppendChild(CreateDmElement("dateTime", DateTime.Now.ToString("yyMMdd_HHmm")));
      _ = revision.AppendChild(CreateDmElement("outputFolder", Path.Combine(H.GetSProperty("processPath"),
                                                                            GetValueString("opportunityFolder"),
                                                                            SBidRevision,
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
        XmlElement newChild = CreateDmElement("doc", target.FileName);
        newChild.SetAttribute("code", target.Code);
        newChild.SetAttribute("version", target.Version);
        newNode.AppendChild(newChild);
      }
      _ = newNode != null ? revision.AppendChild(_dm.ImportNode(newNode, true)) : null;



      //Adding node to store tools used in this revision
      newNode = H.CreateElement(_dm, "tools", "");
      newNode.SetAttribute("processedFolder", Path.Combine(H.GetSProperty("processPath"), $@"{opportunityFolder}\{SBidRevision}\TOOLS"));
      _ = newNode != null ? revision.AppendChild(_dm.ImportNode(newNode, true)) : null;
      return revision;
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
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌ - ProcessFile", $"⚠️ File: '{filePath}' is not found. ");
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

        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"Archivo '{filePath}' registered. ");
      }

      H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"All input files have been registered. ");

      return newInputDocs;
    }
    private XmlElement CreateDmElement(string name, string value)
    {
      XmlElement element = _dm.CreateElement(name);
      element.InnerText = value;
      return element;
    }
    private VariableData StoreValue(string id, string value, string origin = "", string notes = "")
    {
      // Log con InstanceId de DataMaster y HashCode del diccionario
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  - variable ||{id}: {value}|| added/updated to DataMaster data");
      VariableData? varData = _vm.GetNewVariableData(id); 

      // Log con InstanceId de VariableData clonada
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  ->  for key '{id}'. Original Value (before update): '{varData.Value}'. ()");

      varData.Value = (varData.Type != "num") ? value : value.Replace(',', '.');

      varData.Origen = origin;

      varData.Note = notes;

      if (_data.ContainsKey(id))
      {
        _data[id] = varData; // Actualizar valor existente
        H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Updated key '{id}' in _data. New VariableData  New Value: '{varData.Value}'. ()");
      }
      else
      {
        _data.Add(id, varData); // Añadir nuevo valor
        H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Added key '{id}' to _data. New VariableData  Value: '{varData.Value}'. ()");
      }
      return varData;
    }
    private void StoreValue(string id, VariableData varData)
    {
      // Log con InstanceId de DataMaster y HashCode del diccionario
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  - variable ||{id}|| added/updated to DataMaster data (from existing VariableData instance)");

      // Log con InstanceId de VariableData (la que se está pasando)
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  -> VariableData for key '{id}'. Value: '{varData.Value}'. ()");

      if (_data.ContainsKey(id))
      {
        _data[id] = varData; // Actualizar valor existente
        H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Updated key '{id}' in _data with passed VariableData. New VariableData  Value: '{varData.Value}'. ()");
      }
      else
      {
        _data.Add(id, varData); // Añadir nuevo valor
        H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Added key '{id}' to _data with passed VariableData. New VariableData  Value: '{varData.Value}'. ()");
      }
    }
    private XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
    {
      XmlElement?
        sourceElement = (XmlElement)sourceDoc.DocumentElement!.SelectSingleNode(elementName)! ??
        throw new XmlException($"Element '{elementName}' not found in the source document.");
      XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);
      return importedElement;
    }
    // public Methods
    public void UpdateData(XmlDocument newData)
    {
      // Log con InstanceId de DataMaster y HashCode del diccionario
      H.PrintLog(5, TC.ID.Value!.Time(), User, "DataMaster.UpdateData", $"Actualizando   para '{GetValueString("opportunityFolder")}' con nueva data.");
      H.MergeXmlNodes(newData, _dm, "/*/utils", $"/dm/utils/{SBidRevision}");

      XmlNode variablesNode = newData.SelectSingleNode("//*/variables")!;
      if (variablesNode == null) return;

      foreach (XmlNode variableNode in variablesNode.ChildNodes)
      {
        VariableData var = StoreValue(variableNode.Name,
                                        variableNode.SelectSingleNode("value").InnerXml,
                                        variableNode.SelectSingleNode("origin").InnerText,
                                        variableNode.SelectSingleNode("note").InnerText);

        XmlNode? variableDMNode = _dataNode.SelectSingleNode(XmlConvert.EncodeName(var.ID));
        if (variableDMNode == null)
        {
          variableDMNode = _dm.CreateElement(XmlConvert.EncodeName(var.ID));
          _dataNode.AppendChild(variableDMNode);
        }


        //Create the revision index
        XmlNode? revisionIndex = variableDMNode.SelectSingleNode("revision");
        if (revisionIndex == null)
        {
          revisionIndex = _dm.CreateElement("revision");
          variableDMNode.AppendChild(revisionIndex);

        }
        // for now: adding the value to the set element. for now creating a new set, later: seeking for the right set, if not exists create
        revisionIndex.AppendChild(CreateDmElement(SBidRevision, $"set{BidRevision.ToString("D2")}"));


        // for now: creating the set node without looking for an existing set with the same value. Later: seek for the right set, if not exists create
        XmlElement setNode = _dm.CreateElement($"set{BidRevision.ToString("D2")}");
        variableDMNode.AppendChild(setNode);
        setNode.AppendChild(CreateDmElement("value", var.Value));
        setNode.AppendChild(CreateDmElement("origin", var.Origen));
        setNode.AppendChild(CreateDmElement("Note", var.Note));
      }
    }
    public void SaveDataMaster()
    {
      _dm.Save(FileName);
      H.PrintLog(3, TC.ID.Value!.Time(), User, "DM", $"XML guardado en {FileName}.  ");
    }
    public string GetValueString(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        // Log con InstanceId de DataMaster y HashCode del diccionario, y InstanceId de VariableData
        H.PrintLog(1, TC.ID.Value!.Time(), User, "GetValueString", $"- Reading key '{key}'. Value: '{value!.Value}'");
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
            return _dm.SelectSingleNode($"/dm/utils/{SBidRevision}/{key}")?.InnerText ?? string.Empty;
          }
          else
          {
            string attribute = key[(dashIndex + 1)..];
            key = key[..dashIndex];
            string node = $"/dm/utils/{SBidRevision}/{key}";

            return _dm.SelectSingleNode(node)?.Attributes?[attribute]?.Value ?? string.Empty;
          }
        }

        H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌ - DM.GetValueString ", $"Key '{key}' not found in Dictionary");
        return string.Empty;
      }
    }
    public double? GetValueNumber(string key)
    {
      return double.TryParse(GetValueString(key), NumberStyles.Float, CultureInfo.InvariantCulture, out double num) ? num : null;
    }
    public bool? GetValueBoolean(string key)
    {
      return bool.TryParse(GetValueString(key), out bool num) ? num : null;
    }
    public List<string> GetValueList(string key, bool isNumber = false)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        string xmlString = value?.Value ?? string.Empty;
        List<string> itemList = new();

        if (string.IsNullOrWhiteSpace(xmlString))
          return null;

        XmlDocument tempDoc = new();
        try
        {
          tempDoc.LoadXml($"{xmlString}"); // Assuming the value is already a well-formed XML fragment as <l><li>item1</li><li>item2</li></l>

          foreach (XmlNode item in tempDoc.DocumentElement.ChildNodes)
          {
            if (isNumber)
            {
              //test parsing the item as a number, and use . as decimal separator
              if (double.TryParse(item.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
              {
                itemList.Add(num.ToString(CultureInfo.InvariantCulture));
              }
            }
            else
            {
              itemList.Add(item.InnerText);
            }
          }
          return itemList;
        }
        catch
        {
          H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetValueList", $"Key '{key}' found in, but value is not valid XML: '{xmlString}'.");
          return null; // returns NULL when the Value has not XML format
        }
      }
      H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM", $"Key '{key}' not found in . ");
      throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
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
          H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetValueXmlNode", $"Key '{key}' found in  , but value is not valid XML: '{xmlString}'.");
          return null; // returns NULL when the Value has not XML format
        }
      }
      else
      {
        if (_vm.GetVariableData(key).Source == "UTILS")
        {
          key = key.Replace('.', '/');
          return _dm.SelectSingleNode($"/dm/utils/{SBidRevision}/{key}");
        }
      }
      H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM", $"Key '{key}' not found in . ");
      throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
    }
    public string GetValueUnit(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value!.Unit;
      }
      else
      {
        return string.Empty; // return empty string if the key exists but it has no unit
      }
    }
    public VariableData GetVariableData(string key)
    {
      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value;
      }
      else
      {
        H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetVariableData", $"Key '{key}' not found in . ");
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
        H.PrintLog(5, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetInnerText", $"Node not found for XPath: {xpath} in . ");
        throw new XmlException($"Node not found for XPath: {xpath}");
      }
    }
    public void CheckMandatoryValues()
    {
      List<string> missingValues = [];

      foreach (var kvp in _data)
      {
        string variableId = kvp.Key;
        // La llamada a GetValueString ya loguea el InstanceId si hay un error
        if (kvp.Value.Mandatory && string.IsNullOrWhiteSpace(kvp.Value.Value)) // Acceder directamente a kvp.Value
        {
          missingValues.Add(variableId);
        }

      }

      if (missingValues.Count > 0)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), User, "CheckMandatoryValues", $"❌Error❌: Mandatory values not found in . Cannot continue with calculations. Faltan: {string.Join(", ", missingValues)}");
        throw new InvalidOperationException("MandatoryValues missing");
      }
    }
  }
}