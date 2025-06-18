using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Office.Interop.Word;

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


        // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
        public DataMaster(XmlDocument xmlRequest)
        {
            _vm = VariablesMap.Instance;
            BidRevision = 1;
            _dm = new XmlDocument();
            _data = new Dictionary<string, XmlNode>();

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
                configDataXML.AppendChild(configDataRoot);
                XmlElement variables = configDataXML.CreateElement("variables");
                configDataRoot.AppendChild(variables);


                foreach (VariableData variable in varList)
                    {

                    // Load Basic Data
                    if (variable.Area == "projectData") _projectDataNode.AppendChild(GetImportedElement(xmlRequest, @$"//projectData/{variable.ID}"));

                    // Load Config. Init Data (creating an xml with config data variables and send it to UpdateData for inserting in DM
                    if (variable.Area == "config") {
                        _dataNode.AppendChild(GetImportedElement(xmlRequest, @$"//config/{variable.ID}"));

                        // Variable 
                        XmlElement newVar = configDataXML.CreateElement(variable.ID);
                        variables.AppendChild(newVar);

                        // Value node
                        XmlElement value = configDataXML.CreateElement("value");

                        XmlElement importedElement = GetImportedElement(xmlRequest, @$"//config/{variable.ID}");
                        XmlAttribute unitAttribute = importedElement?.GetAttributeNode("unit");

                        if (unitAttribute != null) // Only add "unit" if it exists
                            value.SetAttribute("unit", unitAttribute.Value);

                        value.InnerText = importedElement?.InnerText ?? ""; // Avoid null reference

                        newVar.AppendChild(value);

                        // Origin node
                        XmlElement origin = configDataXML.CreateElement("origin");
                        origin.InnerText = "INIT from callXML";
                        newVar.AppendChild(origin);

                        // Note node
                        XmlElement note = configDataXML.CreateElement("note");
                        note.InnerText = "Variable leida de Hermes";
                        newVar.AppendChild(note);
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
                utilsData.AppendChild(GetImportedElement(xmlRequest, "//requestInfo/opportunityFolder"));

                //Add first revision element
                _ = _utilsNode.AppendChild(_dm.CreateComment("First Revision"));
                XmlElement revision = _dm.CreateElement("rev_01");


                XmlElement importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo");

                importedNode = (XmlElement)xmlRequest.SelectSingleNode("//requestInfo");
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

        private XmlElement CreateElement(string name, string value)
        {
            XmlElement element = _dm.CreateElement(name);
            element.InnerText = value;
            return element;
        }

        public void UpdateData(XmlDocument newData)
        {
            XmlNode variablesNode = newData.SelectSingleNode("/root/variables");
            if (variablesNode == null) return;

            foreach (XmlNode variable in variablesNode.ChildNodes)
            {
                XmlNode importedNode = _dm.ImportNode(variable, true);

                XmlElement revisionElement = _dm.CreateElement("revision");
                XmlElement rev01Element = _dm.CreateElement("rev01");
                rev01Element.InnerText = "set01";
                _ = revisionElement.AppendChild(rev01Element);

                XmlElement set01Element = _dm.CreateElement("set01");

                foreach (XmlNode child in importedNode.ChildNodes)
                {
                    _ = set01Element.AppendChild(child.CloneNode(true));
                    if (child.Name == "value")
                    {
                        _data.Add(variable.Name, child);
                    }
                }

                importedNode.RemoveAll();
                _ = importedNode.AppendChild(revisionElement);
                _ = importedNode.AppendChild(set01Element);

                _ = _dataNode.AppendChild(importedNode);
            }
        }

        public XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
        {
            XmlElement sourceElement = (XmlElement)sourceDoc.DocumentElement.SelectSingleNode(elementName);
            if (sourceElement == null)
            {
                throw new XmlException($"Element '{elementName}' not found in the source document.");
            }

            XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);
            return importedElement;
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

        public string GetStringValue(string key)
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

        public XmlNode GetXmlValue(string key)
        {
            if (_data.ContainsKey(key))
            {
                return _data[key];
            }
            else
            {
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

        private static XmlElement CreateElement(XmlDocument doc, string name, string value)
        {
            XmlElement element = doc.CreateElement(name);
            element.InnerText = value;
            return element;
        }
    }
}
