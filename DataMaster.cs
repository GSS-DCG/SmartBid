using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SmartBid
{
    public class DataMaster
    {
        private XmlDocument _dm;
        private VariablesMap _vm;
        private Dictionary<string, XmlNode> _data;
        private XmlNode _projectDataNode;
        private XmlNode _utilsNode;
        private XmlNode _dataNode;

        public string FileName { get; set; }
        public string User { get; set; } = ThreadContext.CurrentThreadInfo.Value.User;
        public Dictionary<string, XmlNode> Data { get { return _data; } }
        public XmlDocument DM { get { return _dm; } }
        public int BidRevision { get; set; }
        public string sBidRevision => BidRevision.ToString("D2");



        // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
        public DataMaster(XmlDocument xmlRequest)
        {
            _vm = VariablesMap.Instance;
            BidRevision = 1;
            _dm = new XmlDocument();
            _data = new Dictionary<string, XmlNode>();

            // register actual revision number in _data (no need to store it in DM)
            StoreValue("revision", H.CreateElement(DM, "value", "rev_01"));

            if (((XmlElement)xmlRequest.SelectSingleNode("/request/requestInfo")).GetAttribute("Type") == "create")
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


                foreach (VariableData variable in varList)
                {

                    // Load PROJECTDATA Data
                    if (variable.Area == "projectData")
                    {
                        XmlElement element = GetImportedElement(xmlRequest, @$"//projectData/{variable.ID}");
                        _ = _projectDataNode.AppendChild(element);

                        StoreValue(variable.ID, element);
                    }

                    // Load CONFIG Init Data (creating an xml with config data variables and send it to UpdateData for inserting in DM
                    if (variable.Area == "config")
                    {
                        //tomamos el nombre de la variable
//                        _ = _dataNode.AppendChild(GetImportedElement(xmlRequest, @$"//config/{variable.ID}"));

                        // Variable 
                        XmlElement newVar = configDataXML.CreateElement(variable.ID);
                        _ = variables.AppendChild(newVar);

                        // Value node

                        XmlElement importedElement = GetImportedElement(xmlRequest, @$"//config/{variable.ID}");
                        XmlAttribute unitAttribute = importedElement?.GetAttributeNode("unit");


                        XmlNode value = newVar.AppendChild(CreateElement(configDataXML, "value", importedElement?.InnerText ?? ""));
                        if (unitAttribute != null)
                            ((XmlElement)value).SetAttribute("unit", unitAttribute.Value);
                        _ = newVar.AppendChild(CreateElement(configDataXML, "origin", "INIT from callXML"));
                        _ = newVar.AppendChild(CreateElement(configDataXML, "note", "Variable leida de Hermes"));
                    }
                }
                // Adding variable to the data dictionary
                UpdateData(configDataXML);

                // Load Utils Data
                XmlNode utilsData = _utilsNode.AppendChild(DM.CreateElement("utilsData"));

                // Check if opportunityFolder exists, otherwise throw an exception
                if (GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder") == null)
                {
                    H.PrintLog(5, User, "CargaXML", "⚠️ Nodo 'opportunityFolder' no encontrado.");
                    throw new InvalidOperationException("El XML está incompleto: falta '//requestInfo/opportunityFolder'.");
                }

                // Add opportunityFolder to dataMaster and _data dictionary
                _ = utilsData.AppendChild(GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder"));
                StoreValue("opportunityFolder", GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder"));
                StoreValue("createdBy", GetImportedElement(xmlRequest, "//requestInfo/createdBy"));

                //Add first revision element
                _ = _utilsNode.AppendChild(_dm.CreateComment("First Revision"));
                XmlElement revision = _dm.CreateElement("rev_01");

                revision.AppendChild(CreateElement("dateTime", DateTime.Now.ToString("yyMMdd_HHmm")));

                XmlElement importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo");
                _ = importedNode != null ? revision.AppendChild(_dm.ImportNode(importedNode, true)) : null;

                importedNode = (XmlElement)xmlRequest.SelectSingleNode("//bidVersion/deliveryDocs");
                _ = importedNode != null ? revision.AppendChild(_dm.ImportNode(importedNode, true)) : null;

                importedNode = (XmlElement)xmlRequest.SelectSingleNode("//bidVersion/inputDocs");
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

                XmlElement set01Element = _dm.CreateElement($"set{sBidRevision}");

                foreach (XmlNode child in importedNode.ChildNodes)
                {
                    _ = set01Element.AppendChild(child.CloneNode(true));
                    if (child.Name == "value")
                    {
                       StoreValue(variable.Name, (XmlElement)child);
                    }
                }

                importedNode.RemoveAll();
                _ = importedNode.AppendChild(revisionElement);
                _ = importedNode.AppendChild(set01Element);

                _ = _dataNode.AppendChild(importedNode);
            }
        }

        public void SaveDataMaster()
        {
            string filePath;
            if (FileName == null)
            {
                string opportunityFolder = _utilsNode["utilsData"]["opportunityFolder"]?.InnerText ?? "";
                filePath = Path.Combine(H.GetSProperty("processPath"), opportunityFolder, $"{opportunityFolder.Substring(0, 7)}_DataMaster.xml");
                FileName = filePath;
            }
            else
            {
                filePath = FileName;
            }

            _dm.Save(filePath);
            H.PrintLog(4, User, "DM", $"XML guardado en {filePath}");
        }

        public string GetValueString(string key)
        {
            if (_data.ContainsKey(key))
            {
                return _data[key]?.FirstChild.Value.ToString() ?? string.Empty;
            }
            else
            {
                throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
            }
        }

        public double? GetValueNumber(string key)
        {
            if (_data.ContainsKey(key))
            {
                return double.TryParse(_data[key]?.FirstChild.Value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double num) ? num : null;
            }
            else
            {
                H.PrintLog(5, User, "Error - DM", $"Key '{key}' not found in DataMaster.");
                throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
            }
        }

        public bool? GetValueBoolean(string key)
        {
            if (_data.ContainsKey(key))
            {
                return bool.TryParse(_data[key]?.FirstChild.Value.ToString(), out bool num) ? num : null;
            }
            else
            {
                H.PrintLog(5, User, "Error - DM", $"Key '{key}' not found in DataMaster.");
                throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
            }
        }

        public XmlNode GetValueXmlNode(string key)
        {
            if (_data.ContainsKey(key))
            {
                return _data[key];
            }
            else
            {
                H.PrintLog(5, User, "Error - DM", $"Key '{key}' not found in DataMaster.");
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

        private void StoreValue(string id, XmlElement value)
        {
            H.PrintLog(1, User, "StoreValue", $"variable ||{id}|| added to DataMaster data");
            _data.Add(id, value);
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
