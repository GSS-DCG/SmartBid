using System.Text.RegularExpressions;
using System.Xml;
using Windows.Devices.PointOfService;


namespace SmartBid
{
    public class DataMaster
    {
    private static DataMaster _instance;
    private static readonly object _lock = new object();

    public static DataMaster Instance
    {
        get
        {
            if (_instance == null)
            {
                throw new InvalidOperationException("DataMaster is not initialized. Call Initialize first.");
            }
            return _instance;
        }
    }

    public static void Initialize(XmlDocument initData)
    {
        lock (_lock)
        {
            if (_instance == null)
            {
                _instance = new DataMaster(initData);
            }
        }
    }

    public static void Initialize(string dmFileName)
    {
        lock (_lock)
        {
            if (_instance == null)
            {
                _instance = new DataMaster(dmFileName);
            }
        }
    }

        private XmlDocument _dm;
        private int _bidRevision;
        private Dictionary<string, object> _data;
        public string FileName { get; set; }
        public Dictionary<string, object> Data { get { return _data; } }
        private XmlNode _initDataNode;
        private XmlNode _utilsNode;
        private XmlNode _dataNode;

        private DataMaster(XmlDocument initData)
        {
            // Initialize the DataMaster with the provided XmlDocument
            // -- the DataMaster will be created
            // -- revision Number will be set to 01 
            // -- and revisionData are loaded to into the DataMaster

            CreateDataMaster(initData);

        }
    private DataMaster(string dmFileName)
        {
            // Load existingDataMaster 
            // -- the DataMaster will be read
            // -- revision Number will be set to the next one after last review

            LoadDataMaster(dmFileName);


        }
        private void CreateDataMaster(XmlDocument initData)
        {
            _bidRevision = 1; // Default revision number

            _dm = new XmlDocument();
            XmlDeclaration xmlDeclaration = _dm.CreateXmlDeclaration("1.0", "utf-8", null);
            _ = _dm.AppendChild(xmlDeclaration);

            // Root element
            XmlElement root = _dm.CreateElement("dm");
            _ = _dm.AppendChild(root);

            XmlElement init = _dm.CreateElement("initData");
            _initDataNode = root.AppendChild(init);

            XmlElement utils = _dm.CreateElement("utils");
            _utilsNode = root.AppendChild(utils);

            XmlElement data = _dm.CreateElement("data");
            _dataNode = root.AppendChild(data);

            //Dictionary to hold values of all variables
            _data = new Dictionary<string, object>();

            LoadBasicData(initData);

            LoadUtilsData(initData);
        }

        private void LoadDataMaster(string dmFileName)
        {

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
            // Crear comentario "Primera revisión"
            XmlComment revisionComment = _dm.CreateComment("Primera revisión");
            _ = _utilsNode.AppendChild(revisionComment);

            // Crear el nodo revision
            XmlElement revision = _dm.CreateElement("rev_01");

            // Seleccionar el elemento DeliveryDocs y agregarlo a revision
            XmlElement importedNode = (XmlElement)initData.SelectSingleNode("//DeliveryDocs");
            if (importedNode != null)
            {
                XmlNode importedSourceDocs = _dm.ImportNode(importedNode, true);
                _ = revision.AppendChild(importedSourceDocs);
            }

            // Seleccionar el elemento sourceDocs y agregarlo a revision
            importedNode = (XmlElement)initData.SelectSingleNode("//sourceDocs");
            if (importedNode != null)
            {
                XmlNode importedSourceDocs = _dm.ImportNode(importedNode, true);
                _ = revision.AppendChild(importedSourceDocs);
            }

            // Agregar el nodo revision al documento
            _ = _utilsNode.AppendChild(revision);
        }

        public void UpdateData(XmlDocument newData)
        {

                XmlNode variablesNode = newData.SelectSingleNode("/root/variables");
            if (variablesNode == null) return;

            foreach (XmlNode variable in variablesNode.ChildNodes)
            {

                // Import the node into the target document
                XmlNode importedNode = _dm.ImportNode(variable, true);

                // Create the <revision> element
                XmlElement revisionElement = _dm.CreateElement("revision");
                XmlElement rev01Element = _dm.CreateElement("rev01");
                rev01Element.InnerText = "set01";

                // Append <rev01> to <revision>
                revisionElement.AppendChild(rev01Element);

                // Create the <set01> element
                XmlElement set01Element = _dm.CreateElement("set01");

                // Move the child nodes under <set01>
                foreach (XmlNode child in importedNode.ChildNodes)
                {
                    set01Element.AppendChild(child.CloneNode(true));
                    if (child.Name == "value") { 
                        _data.Add(variable.Name, child.FirstChild.Value); // Populate the internal _data dictionary
                    }
                }

                // Create the <origin> element and append it to <set01>
                XmlElement originElement = _dm.CreateElement("origin");
                originElement.InnerText = "prepTool+21-23:24";
                set01Element.AppendChild(originElement);

                // Append <revision> and <set01> to the imported node
                importedNode.RemoveAll();  // Clear existing children to restructure
                importedNode.AppendChild(revisionElement);
                importedNode.AppendChild(set01Element);

                // Append the modified node to _dataNode
                _dataNode.AppendChild(importedNode);
            }
        }

        public XmlElement GetImportedElement(XmlDocument sourceDoc, string elementName)
        {
            // Select the element from the source document
            XmlElement sourceElement = (XmlElement)sourceDoc.DocumentElement.SelectSingleNode($"//initData/{elementName}");
            if (sourceElement == null)
            {
                throw new XmlException($"Element '{elementName}' not found in the source document.");
            }

            // Import the node into the new document
            XmlElement importedElement = (XmlElement)_dm.ImportNode(sourceElement, true);

            return importedElement;
        }

        public void SaveDataMaster()
        {
            string filePath;
            if (FileName == null)
            {
            int i = GetNextDataMasterNumber(Path.GetDirectoryName(H.GetSProperty("dataMaster")));
            string dmName = new string($"DataMaster_{i.ToString("D2")}.xml");
            filePath = Path.Combine(H.GetSProperty("dataMaster"), dmName);
                FileName = filePath;
            }
            else
            {
                filePath = FileName;
            }
            _dm.Save(filePath);
            Console.WriteLine($"XML guardado en {filePath}");
        }

        static int GetNextDataMasterNumber(string directory)
        {
            string pattern = @"DataMaster_(\d+)\.xml"; // Expresión regular para encontrar números
            if (!Directory.Exists(directory))
                throw new DirectoryNotFoundException($"El directorio '{directory}' no existe.");

            var files = Directory.GetFiles(directory, "DataMaster_*.xml");
            var numbers = files.Select(file =>
            {
                var match = Regex.Match(Path.GetFileName(file), pattern);
                return match.Success ? int.Parse(match.Groups[1].Value) : (int?)null;
            }).Where(n => n.HasValue).Select(n => n.Value).ToList();

            return numbers.Count > 0 ? numbers.Max() + 1 : 1; // Si no hay archivos, comienza en 1
        }

        public string GetValue (string key)
        {
            if (_data.ContainsKey(key))
            {
                return _data[key]?.ToString() ?? string.Empty;
            }
            else
            {
                throw new KeyNotFoundException($"Key '{key}' not found in DataMaster.");
            }
        }

        static void _Main(string[] args)
        {
            string callFile = Path.Combine(Path.GetFullPath(H.GetSProperty("CallFile")));
            XmlDocument xmlCall = new XmlDocument();

            if (File.Exists(callFile))
                xmlCall.Load(callFile);
            else
                throw new FileNotFoundException($"****** FILE '{callFile}' NOT FOUND ******.");


            DataMaster dataMaster = new DataMaster(xmlCall);
        }


    }

}




