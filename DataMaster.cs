
using System.Text.RegularExpressions;
using System.Xml;

namespace SmartBid
{
    public class DataMaster
    {
        private XmlDocument _dm;
        private int _bidRevision;
        private Dictionary<string, XmlNode> _data;
        public string FileName { get; set; }

        private XmlNode _initDataNode;
        private XmlNode _utilsNode;
        private XmlNode _dataNode;

        public Dictionary<string, XmlNode> Data { get { return _data; } }
        public XmlDocument DM { get { return _dm; } }


        // Constructor público con XmlDocument ==> para crear un nuevo DataMaster
        public DataMaster(XmlDocument initData)
        {
            CreateDataMaster(initData);
        }

        // Constructor público con nombre de archivo ==> para cargar un DataMaster existente
        public DataMaster(string dmFileName)
        {
            LoadDataMaster(dmFileName);
        }

        private void CreateDataMaster(XmlDocument xmlRequest)
        {
            _bidRevision = 1;
            _dm = new XmlDocument();
            XmlDeclaration xmlDeclaration = _dm.CreateXmlDeclaration("1.0", "utf-8", null);
            _ = _dm.AppendChild(xmlDeclaration);

            XmlElement root = _dm.CreateElement("dm");
            _ = _dm.AppendChild(root);

            XmlElement init = _dm.CreateElement("initData");
            _initDataNode = root.AppendChild(init);

            XmlElement utils = _dm.CreateElement("utils");
            _utilsNode = root.AppendChild(utils);

            XmlElement data = _dm.CreateElement("data");
            _dataNode = root.AppendChild(data);

            _data = new Dictionary<string, XmlNode>();

            LoadBasicData(xmlRequest);
            LoadUtilsData(xmlRequest);
        }

        private void LoadDataMaster(string dmFileName) //Open existing DM
        {
            // Implementación pendiente
        }

        private XmlElement CreateElement(string name, string value)
        {
            XmlElement element = _dm.CreateElement(name);
            element.InnerText = value;
            return element;
        }

        private void LoadBasicData(XmlDocument initData)
        {
            List<VariableData> varList = VariablesMap.Instance.GetVarListBySource("INIT");
            foreach (VariableData variable in varList)
                _ = _initDataNode.AppendChild(GetImportedElement(initData, variable.ID));
        }

        private void LoadUtilsData(XmlDocument initData)
        {

            XmlNode utilsData = _utilsNode.AppendChild(DM.CreateElement("utilsData"));

            string projectFolder = 
                $"{DateTime.Now:yy}_" +
                $"{GetNextFolderNumber().ToString("D4")}_" +
                $"{_initDataNode["client"]?.InnerText ?? ""}_" +
                $"{_initDataNode["projectName"]?.InnerText ?? ""}";

            _ = utilsData.AppendChild(CreateElement(DM, "projectFolder", projectFolder));



            _ = _utilsNode.AppendChild(_dm.CreateComment("Primera revisión"));

            XmlElement revision = _dm.CreateElement("rev_01");

            XmlElement importedNode;

            importedNode = (XmlElement)initData.SelectSingleNode("//requestInfo");
            if (importedNode != null)
            {
                XmlNode importedSourceDocs = _dm.ImportNode(importedNode, true);
                _ = revision.AppendChild(importedSourceDocs);
            }
            importedNode = (XmlElement)initData.SelectSingleNode("//bidVersion/deliveryDocs");
            if (importedNode != null)
            {
                XmlNode importedSourceDocs = _dm.ImportNode(importedNode, true);
                _ = revision.AppendChild(importedSourceDocs);
            }

            importedNode = (XmlElement)initData.SelectSingleNode("//bidVersion/inputDocs");
            if (importedNode != null)
            {
                XmlNode importedSourceDocs = _dm.ImportNode(importedNode, true);
                _ = revision.AppendChild(importedSourceDocs);
            }

            _ = _utilsNode.AppendChild(revision);
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

                XmlElement originElement = _dm.CreateElement("origin");
                originElement.InnerText = "prepTool+21-23:24";
                _ = set01Element.AppendChild(originElement);

                importedNode.RemoveAll();
                _ = importedNode.AppendChild(revisionElement);
                _ = importedNode.AppendChild(set01Element);

                _ = _dataNode.AppendChild(importedNode);
            }
        }

        public XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
        {
            XmlElement sourceElement = (XmlElement)sourceDoc.DocumentElement.SelectSingleNode(@$"//projectData/{elementName}");
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
                string projectFolder = _utilsNode["utilsData"]["projectFolder"]?.InnerText ?? "";
                filePath = Path.Combine(H.GetSProperty("storagePath"), projectFolder, $"{projectFolder.Substring(0, 7)}_DataMaster.xml");
                FileName = filePath;
            }
            else
            {
                filePath = FileName;
            }

            _dm.Save(filePath);
            H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "DM", $"XML guardado en {filePath}");
        }

        static int GetNextFolderNumber()
        {
            string directory = H.GetSProperty("storagePath");
            if (!Directory.Exists(directory))
            {
                H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error-GetNextFolderNumber", "El directorio no existe.");
                return -1;
            }

            var folders = Directory.GetDirectories(directory)
                .Select(folder => Path.GetFileName(folder))
                .Where(name => name.Length >= 7 && name[2] == '_' && name.Substring(3, 4).All(char.IsDigit)) // Verifica formato correcto
                .OrderByDescending(name => int.Parse(name.Substring(3, 4))) // Ordena por número de proyecto
                .FirstOrDefault(); // Toma el mayor número encontrado

            return folders != null ? int.Parse(folders.Substring(3, 4)) + 1 : 1;
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
