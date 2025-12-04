using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Office.CoverPageProps;
using Microsoft.Office.Interop.Access;
using Org.BouncyCastle.Asn1.Pkcs;
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
        _ = StoreValue("opportunityFolder", initialOpportunityFolder); // Usar el valor capturado
        _ = StoreValue("createdBy", GetImportedElement(xmlRequest, "//requestInfo/createdBy").InnerText);
        _ = StoreValue("requestTimestamp", GetImportedElement(xmlRequest, "//requestInfo/requestTimestamp").InnerText);
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

            _ = StoreValue(variable.ID, element.InnerText);//Extraer el valor para almacenar en _data.
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
        _ = _utilsNode.AppendChild(revision);
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

      SaveDataMaster();

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
    private static bool TryNormalizeNumber(string original, out string corrected)
    {
      corrected = string.Empty;

      if (original == null || string.IsNullOrWhiteSpace(original))
      {
        // Empty allowed
        return true;
      }

      string s = original.Trim();

      // Handle negative numbers in accounting format: (123) -> -123
      bool parenNegative = false;
      if (s.StartsWith("(") && s.EndsWith(")"))
      {
        parenNegative = true;
        s = s.Substring(1, s.Length - 2).Trim();
      }

      // Remove common thousands separators
      s = s.Replace("'", "").Replace(" ", "").Replace("\u00A0", "").Replace("_", "");

      // Detect decimal separator
      int lastComma = s.LastIndexOf(',');
      int lastDot = s.LastIndexOf('.');
      bool hasComma = lastComma >= 0;
      bool hasDot = lastDot >= 0;

      bool hadDecimalSeparator = false;

      if (hasComma && hasDot)
      {
        int lastPos = Math.Max(lastComma, lastDot);
        char decimalSep = s[lastPos];
        hadDecimalSeparator = true;

        if (decimalSep == ',')
        {
          s = s.Replace(".", "");
          s = s.Replace(',', '.');
        }
        else
        {
          s = s.Replace(",", "");
        }
      }
      else if (hasComma)
      {
        hadDecimalSeparator = true;
        s = s.Replace(',', '.');
      }
      else if (hasDot)
      {
        hadDecimalSeparator = true;
      }

      if (parenNegative)
        s = "-" + s;

      // Validate and parse
      if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
        return false;

      // Format final corrected value
      const double EPS = 1e-12;
      if (!hadDecimalSeparator)
      {
        if (Math.Abs(num - Math.Round(num)) < EPS)
          corrected = ((long)Math.Round(num)).ToString(CultureInfo.InvariantCulture);
        else
          corrected = num.ToString(CultureInfo.InvariantCulture);
      }
      else
      {
        corrected = num.ToString(CultureInfo.InvariantCulture);
      }

      return true;
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
      _ = revision.AppendChild(ProcessInputFiles((XmlElement)xmlRequest.SelectSingleNode("//requestInfo/inputDocs")!));

      //Adding deliveryDocs from either Call or Default Delivery Docs
      newNode = H.CreateElement(_dm, "deliveryDocs", "");
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
          H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌ - ProcessFile", $"⚠️ File: '{filePath}' is not found. ");
          continue; // Saltar este documento y seguir con los demás
        }

        string hash = H.GetFileMD5(filePath); // Calculate MD5 hash for the fileName
        string lastModified = File.GetLastWriteTime(filePath).ToString("yyyy-MM-dd HH:mm:ss");

        _ = newInputDocs.AppendChild(H.CreateElement(_dm, fileType, filePath, new Dictionary<string, string>
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
      if (string.IsNullOrEmpty(name))
        throw new ArgumentException("El nombre del elemento no puede ser nulo ni vacío.", nameof(name));

      XmlElement element = _dm.CreateElement(name);

      if (string.IsNullOrEmpty(value))
      {
        element.InnerText = string.Empty;
        return element;
      }

      if (H.IsWellFormedXml(value))
      {
        element.InnerXml = value;
      }
      else
      {
        element.InnerText = value;
      }

      return element;
    }
    private VariableData StoreValue(string id, string value, string origin = "", string notes = "")
    {
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  - variable ||{id}: {value}|| added/updated to DataMaster data");
      VariableData? varData = _vm.GetNewVariableData(id);


      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  ->  for key '{id}'. Original Value (before update): '{varData.Value}'. ()");

      //make sure value is valid according to type (bool, num, xml)

      if (varData != null)
      {
        if (varData.Type == "bool")
        {
          // Normalizar valores booleanos
          value = value.Trim().ToLower();
          if (value == "1") value = "true";
          else if (value == "0") value = "false";

          if (!bool.TryParse(value, out _))
          {
            H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                $"❌❌ Error ❌❌ - Invalid boolean value '{value}' for variable '{id}'.");
          }
        }

        else if (varData.Type == "num")
        {
          if (TryNormalizeNumber(value, out string corrected))
          {
            value = corrected; // empty or normalized
          }
          else if (!string.IsNullOrEmpty(value))
          {
            H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                $"❌❌ Error ❌❌ - Invalid numeric value '{value}' for variable '{id}'.");
          }
        }

        else if (varData.Type == "list<str>" || varData.Type == "table")
        {
          if (!H.IsWellFormedXml(value))
          {
            H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                $"❌❌ Error ❌❌ - Invalid XML value for variable '{id}':\n {value}");
          }
        }

        else if (varData.Type == "list<num>")
        {
          if (!H.IsWellFormedXml(value))
          {
            H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                $"❌❌ Error ❌❌ - Invalid XML value for variable '{id}':\n {value}");
          }
          else
          {
            try
            {
              var doc = XDocument.Parse(value);
              foreach (var item in doc.Descendants("li"))
              {
                if (string.IsNullOrWhiteSpace(item.Value))
                {
                  item.Value = string.Empty;
                  continue;
                }

                if (TryNormalizeNumber(item.Value, out string corrected))
                {
                  item.Value = corrected;
                }
                else
                {
                  H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                      $"❌❌ Error ❌❌ - Invalid numeric value '{item.Value}' in list for variable '{id}'.");
                }
              }
              value = doc.ToString(SaveOptions.DisableFormatting);
            }
            catch (Exception ex)
            {
              H.PrintLog(6, TC.ID.Value!.Time(), User, "StoreValue",
                  $"❌❌ Error ❌❌ - Failed to process list<num> for variable '{id}': {ex.Message}");
            }
          }
        }

      }



      varData.Value = value;
      varData.Origen = origin;
      varData.Note = notes;

      if (_data.ContainsKey(id))
      {
        _data[id] = varData; // Actualizar valor existente
        H.PrintLog(0, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Updated key '{id}' in _data. New VariableData  New Value: '{varData.Value}'. ()");
      }
      else
      {
        _data.Add(id, varData); // Añadir nuevo valor
        H.PrintLog(0, TC.ID.Value!.Time(), User, "StoreValue", $"  -> Added key '{id}' to _data. New VariableData  Value: '{varData.Value}'. ()");
      }
      return varData;
    }
    private XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
    {
      XmlElement?
        sourceElement = (XmlElement)sourceDoc.DocumentElement!.SelectSingleNode(elementName)! ??
        throw new XmlException($"Element '{elementName}' not found in the source document.");
      XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);
      return importedElement;
    }
    private void StoreValue(string id, VariableData varData)
    {
      H.PrintLog(1, TC.ID.Value!.Time(), User, "StoreValue", $"  - variable ||{id}|| added/updated to DataMaster data (from existing VariableData instance)");

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
    private static string listToXmlList(List<string> list)
    {
      XmlDocument tempDoc = new();
      XmlElement root = tempDoc.CreateElement("l");
      _ = tempDoc.AppendChild(root);
      foreach (string item in list)
      {
        XmlElement li = tempDoc.CreateElement("li");
        li.InnerText = item;
        _ = root.AppendChild(li);
      }
      return root.InnerXml;
    }
    private void SaveDataMaster()
    {
      _dm.Save(FileName);
      H.PrintLog(3, TC.ID.Value!.Time(), User, "DM", $"XML guardado en {FileName}.  ");
    }
    // public Methods
    public void UpdateData(XmlDocument newData)
    {
      // Log con InstanceId de DataMaster y HashCode del diccionario
      H.PrintLog(5, TC.ID.Value!.Time(), User, "DataMaster.UpdateData", $"Actualizando DM y mapa de variables");
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
          _ = _dataNode.AppendChild(variableDMNode);
        }

        //Create the revision index
        XmlNode? revisions = variableDMNode.SelectSingleNode("revision");
        if (revisions == null)
        {
          revisions = _dm.CreateElement("revision");
          _ = variableDMNode.AppendChild(revisions);
        }

        XmlNode? rev = revisions.SelectSingleNode(SBidRevision);
        if (rev == null)
        {
          rev = revisions.AppendChild(CreateDmElement(SBidRevision, $"set{BidRevision.ToString("D2")}"));
        }
        string sSet = rev!.InnerText;

        XmlElement? setNode = (XmlElement)variableDMNode.SelectSingleNode(sSet);
        if (setNode == null)
        {
          setNode = _dm.CreateElement(sSet);

          _ = setNode.AppendChild(CreateDmElement("value", variableNode.SelectSingleNode("value").InnerXml));
          _ = setNode.AppendChild(CreateDmElement("origin", var.Origen));
          _ = setNode.AppendChild(CreateDmElement("Note", var.Note));
          _ = variableDMNode.AppendChild(setNode);
        }
        else
        {
          setNode.SelectSingleNode("value").InnerText = var.Value;
          _ = setNode.AppendChild(CreateDmElement("origin", $" Modified: {var.Origen}"));
          _ = setNode.AppendChild(CreateDmElement("Note", $" Modified: {var.Note}"));
        }
      }
      SaveDataMaster();
    }
    public void updateDeliveryDoc(string filePath, string origin, string? code = null, string? version = null)
    {
      XmlNode? deliveryDocsNode = (XmlElement)_dm.SelectSingleNode($"/dm/utils/{SBidRevision}/deliveryDocs");

      if (deliveryDocsNode != null)
      {
        string fileName = Path.GetFileName(filePath);

        // Check if a <doc> element with same fileName and origin exists
        XmlNode existingDoc = deliveryDocsNode.SelectSingleNode($"doc[text()='{fileName}' and @origin='{origin}']");

        if (existingDoc != null)
        {
          // Already exists with origin, do nothing
          H.PrintLog(2, TC.ID.Value!.Time(), User, "DataMaster.updateDeliveryDoc", $"DeliveryDoc: {filePath} already registered with origin in DM");
          return;
        }

        // If no existing doc, create new one
        XmlElement docElement = CreateDmElement("doc", fileName);

        if (!string.IsNullOrEmpty(origin))
          docElement.SetAttribute("origin", origin);
        if (!string.IsNullOrEmpty(code))
          docElement.SetAttribute("code", code);
        if (!string.IsNullOrEmpty(version))
          docElement.SetAttribute("version", version);

        deliveryDocsNode.AppendChild(docElement);
        H.PrintLog(2, TC.ID.Value!.Time(), User, "DataMaster.updateDeliveryDoc", $"DeliveryDoc: {filePath} from Tool registered in DM");
      }
      SaveDataMaster();
    }
    public string GetValueString(string key)
    {
      VariableData varData = _vm.GetVariableData(key);
      if (varData.Type.StartsWith("list") || varData.Type.StartsWith("table"))
      {
        throw new InvalidOperationException($"Key '{key}' is of type {varData.Type}: Cannot be retrieved as String. Use 'GetValueList' instead ");
      }

      if (_data.TryGetValue(key, out VariableData? value))
      {
        H.PrintLog(0, TC.ID.Value!.Time(), User, "GetValueString", $"- Reading key '{key}'. Value: '{value!.Value}'");
        return value!.Value;
      }
      else
      {
        if (varData.Source == "UTILS")
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

        H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌ - DM.GetValueString ", $"Key '{key}' not found in Dictionary");
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
      VariableData varData = _vm.GetVariableData(key);
      //if varType is not list or table, throw exception
      if (!(varData.Type.StartsWith("list") || varData.Type.StartsWith("table")))
      {
        throw new InvalidOperationException($"Key '{key}' is of type {varData.Type}: Cannot be retrieved as List, use GetStringValue instead.");
      }

      if (_data.TryGetValue(key, out VariableData? value))
      {

        string xmlString = value?.Value ?? string.Empty;

        if (string.IsNullOrWhiteSpace(xmlString))
          return new List<string>();

        List<string> itemList = new();
        XmlDocument tempDoc = new();
        try
        {
          tempDoc.LoadXml(xmlString); // Assuming the value is already a well-formed XML fragment as <l><li>item1</li><li>item2</li></l>

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
          H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetValueList", $"Key '{key}' found in, but value is not valid XML: '{xmlString}'.");
          return null; // returns NULL when the Value has not XML format
        }
      }
      else //if key is not found in _data it may be a UTILS variable
      {
        if (varData.Source == "UTILS")
        { 
          //find the xPath for the nodes and iterate to get the list with all values
          key = key.Replace('.', '/');
          int dashIndex = key.IndexOf('-');
          string attribute = dashIndex != -1 ? key[(dashIndex + 1)..] : string.Empty; // ensure attribute is set only if dashIndex is valid
          List<string> list = new();

          //key is the xPath, let's find all items in the xmlNode and fill the itemList with the ittem of each item or the attribute value in case "-" indicates so
          XmlNodeList nodes = _dm.SelectNodes($"/dm/utils/{SBidRevision}/{key}");

            for (int i = 0; i < nodes.Count; i++)
           {

            string item = dashIndex == -1
                ? nodes[i]?.InnerText ?? string.Empty
                : nodes[i].Attributes?[attribute]?.Value ?? string.Empty;

            list.Add(item);
            }
            return list;
          }

        H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetValueList", $"Key '{key}' not found in . ");
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
          H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetValueXmlNode", $"Key '{key}' found in  , but value is not valid XML: '{xmlString}'.");
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
      H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM", $"Key '{key}' not found in . ");
      throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
    }
    public string GetUnit(string key)
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
      VariableData varData = _vm.GetVariableData(key);

      if (_data.TryGetValue(key, out VariableData? value))
      {
        return value;
      }
      else
      {
        if (!(varData.Source == "UTILS"))
        {
          H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetVariableData", $"Key '{key}' not found in . ");
          throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
        }
        varData.Value = listToXmlList(GetValueList(key, varData.Type == "list<num>"));

        return varData;
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
        H.PrintLog(6, TC.ID.Value!.Time(), User, $"❌❌ Error ❌❌  - DM.GetInnerText", $"Node not found for XPath: {xpath} in . ");
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
        H.PrintLog(6, TC.ID.Value!.Time(), User, "CheckMandatoryValues", $"❌Error❌: Mandatory values not found in . Cannot continue with calculations. Faltan: {string.Join(", ", missingValues)}");
        throw new InvalidOperationException("MandatoryValues missing");
      }
    }
  }
}