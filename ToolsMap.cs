using System.Data;
using System.Runtime.InteropServices;
using System.Xml;
using ExcelDataReader;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using File = System.IO.File;
using Word = Microsoft.Office.Interop.Word;

namespace SmartBid
{

    public class ToolData
    {
        public string Resource { get; set; }
        public string ID { get; set; }
        public int Call { get; set; }
        public string Name { get; set; }
        public string FileType { get; set; }
        public string Version { get; set; }
        public string Description { get; set; }
        public string FileName { get; set; }


        // Constructor to initialize all properties
        public ToolData(string resource, string id, int call, string name, string version, string filetype, string description)
        {
            Resource = resource;
            ID = id;
            Call = call;
            Name = name;
            FileType = filetype;
            Version = version;
            Description = description;
            FileName = $"{name}_{version}.{filetype}";
        }

        // Methods
        public XmlDocument ToXMLDocument()
        {
            XmlDocument doc = new XmlDocument();
            _ = doc.AppendChild(ToXML(doc));
            return doc;
        }

        public XmlElement ToXML(XmlDocument mainDoc)
        {
            XmlElement toolElement = mainDoc.CreateElement("tool");
            toolElement.SetAttribute("resource", Resource);
            toolElement.SetAttribute("code", ID);
            toolElement.SetAttribute("call", Call.ToString());
            toolElement.SetAttribute("name", Name);
            toolElement.SetAttribute("fileType", FileType);
            toolElement.SetAttribute("version", Version);
            toolElement.SetAttribute("description", Description);
            toolElement.SetAttribute("fileName", FileName);

            return toolElement;
        }

    }

    public class ToolsMap
    {
        // Private static instance
        private static ToolsMap _instance;

        // Lock object for thread safety
        private static readonly object _lock = new object();

        // Class variables
        public List<ToolData> Tools { get; private set; } = new List<ToolData>();
        private VariablesMap _variablesMap = VariablesMap.Instance; // Instance of VariablesMap

        // Public static property to get the single instance
        public static ToolsMap Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new ToolsMap();
                    }
                    return _instance;
                }
            }
        }

        // Private constructor to prevent instantiation from outside
        private ToolsMap()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string vmFile = Path.GetFullPath(H.GetSProperty("VarMap"));
            if (!File.Exists(vmFile))
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
                _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", vmFile);
            }
            string directoryPath = Path.GetDirectoryName(vmFile);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
            string xmlFile = Path.Combine(directoryPath, "ToolsMap" + ".xml");
            DateTime fileModified = File.GetLastWriteTime(vmFile);
            DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;
            if (fileModified > xmlModified)
            {
                LoadFromXLS(vmFile);
                SaveToXml(xmlFile);
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"XML file created at: {xmlFile}");
            }
            else
            {
                LoadFromXml(xmlFile);
            }
        }

        // Methods
        private void LoadFromXml(string xmlPath)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            foreach (XmlNode node in doc.SelectNodes("//tool"))
            {
                ToolData data = new ToolData(
                    node.Attributes["resource"]?.InnerText ?? string.Empty,
                    node.Attributes["code"]?.InnerText ?? string.Empty,
                    int.TryParse(node.Attributes["call"]?.InnerText, out int callValue) ? callValue : 1, // Call
                    node.Attributes["name"]?.InnerText ?? string.Empty,
                    node.Attributes["version"]?.InnerText ?? string.Empty,
                    node.Attributes["fileType"]?.InnerText ?? string.Empty,
                    node.Attributes["description"]?.InnerText ?? string.Empty
                );

                Tools.Add(data);
            }
        }

        public void SaveToXml(string xmlPath, List<string>? varList = null)
        {
            XmlDocument doc = ToXml(varList);
            doc.Save(xmlPath);
        }

        public XmlDocument ToXml(List<string>? varList = null)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            _ = doc.AppendChild(root);
            foreach (var tool in Tools)
            {
                _ = root.AppendChild(tool.ToXML(doc));
            }
            return doc;
        }

        private void LoadFromXLS(string vmFile)
        {
            DataSet dataSet;
            using (var stream = File.Open(vmFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    using var _ = dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                }
            }

            System.Data.DataTable dataTable = dataSet.Tables["ToolMap"];

            // Iterate over the rows, stopping when column A is empty
            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                DataRow row = dataTable.Rows[i];

                if (row.IsNull(0))
                    break;

                ToolData data = new ToolData(
                    row[0]?.ToString() ?? string.Empty,  // resource 
                    row[1]?.ToString() ?? string.Empty,  // code     
                    int.TryParse(row[2]?.ToString(), out int callValue) ? callValue : 1, //Call
                    row[3]?.ToString() ?? string.Empty,  // name     
                    int.TryParse(row[4]?.ToString(), out int value) ? value.ToString("D3") : "000",  // version  
                    row[5]?.ToString() ?? string.Empty,  // filetype 
                    row[6]?.ToString() ?? string.Empty  // description
                );

                Tools.Add(data);
            }
        }

        public ToolData getToolDataByCode(string code)
        {
            return Tools.FirstOrDefault(tool => tool.ID == code);
        }

        public XmlDocument _CalculateExcel(string toolID, DataMaster dm)
        {
            // 1. Retrieve the tool data by ID
            ToolData tool = ToolsMap.Instance.getToolDataByCode(toolID);
            if (tool == null)
                throw new ArgumentException($"ToolID '{toolID}' not found.");

            // 2. Check if the file type is Excel
            if (!(tool.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) ||
                  tool.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The file is not an Excel type (.xlsx or .xlsm)");

            // 3. Create a MirrorXML instance
            MirrorXML mirror = new MirrorXML(toolID);

            // 4. Get the XXXX list from the mirror
            var variableMap = mirror.VarList;

            // 5. Build the full path to the Excel file
            string filePath = Path.Combine(
                tool.Resource == "TOOL" ? H.GetSProperty("ToolsPath") : H.GetSProperty("TemplatesPath"),
                tool.FileName
            );

            // 6. Crear copia del archivo para trabajar
            string newFilePath = Path.Combine(
                H.GetSProperty("processPath"),
                dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "",
                "TOOLS",
                tool.FileName
            );

            // Asegurar que la carpeta de destino existe antes de copiar
            _ = Directory.CreateDirectory(Path.GetDirectoryName(newFilePath)!);

            // Copiar y sobrescribir si ya existe
            File.Copy(filePath, newFilePath, true);


            // 7. Abrir el archivo en Excel Interop
            H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"Calculating tool {toolID}");
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Open(newFilePath);

            // Escribir valores en las celdas
            foreach (var entry in variableMap)
            {
                if (mirror.GetVarCallLevel(entry.Key) == tool.Call)
                {
                    string variableID = entry.Key;
                    string direction = entry.Value[1];

                    if (direction == "in")
                    {
                        string rangeName = $"{H.GetSProperty("IN_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";
                        string type = _variablesMap.GetVariableData(variableID).Type;

                        if (type != "table")
                        {
                            Excel.Range cell = null;
                            try
                            {
                                cell = workbook.Names.Item(rangeName).RefersToRange;
                                if (cell == null)
                                    throw new Exception($"Named range '{rangeName}' not found in worksheet.");
                            }
                            catch (Exception ex)
                            {
                                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Error accessing range: " + ex.Message);
                            }

                            cell.Value = dm.GetValueString(variableID);
                        }
                        else // Si es una tabla, obtenemos el XML de la tabla y lo escribimos en la hoja de Excel
                        {
                            XmlNode tableData = dm.GetValueXmlNode(variableID);
                            if (tableData.SelectSingleNode("t") != null && tableData.SelectSingleNode("t").HasChildNodes)
                            {
                                WriteTableToExcel(workbook, rangeName, tableData);
                            }
                            else
                            {
                                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"No table data found for variable '{variableID}'.");
                            }
                        }
                    }
                }
            }

            // Forzar cálculo de fórmulas en Excel
            excelApp.Calculate();

            // Guardar los cambios y cerrar Excel
            workbook.Save();
            excelApp.Calculate();

            // Crear un nuevo documento XML para salida
            XmlDocument results = new XmlDocument();
            XmlElement root = results.CreateElement("root");
            _ = results.AppendChild(root);

            XmlElement varNode = results.CreateElement("variables");
            _ = root.AppendChild(varNode);

            // Obtener la fecha y hora actual en formato dd-HH:mm
            string timestamp = DateTime.Now.ToString("dd-HH:mm");

            foreach (var entry in variableMap)
            {
                string variableID = entry.Key;
                string direction = entry.Value[1];

                if (mirror.GetVarCallLevel(variableID) == tool.Call && (direction == "out"))
                {
                    XmlElement newElement = results.CreateElement(variableID);
                    string rangeName = $"{H.GetSProperty("OUT_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";
                    Excel.Range cell = workbook.Names.Item(rangeName).RefersToRange;



                    if (cell.Value != null)
                    {

                        if (_variablesMap.GetVariableData(variableID).Type != "table")
                        {
                            // Si el tipo no es tabla, simplemente añadimos el valor
                            _ = newElement.AppendChild(H.CreateElement(results, "value", cell.Value.ToString()));
                        }
                        else // Si es una tabla, obtenemos los datos de la tabla y los añadimos
                        {
                            XmlNode tableDataXml = ReadTableFromExcel(workbook, rangeName, results);
                            if (tableDataXml != null && tableDataXml.HasChildNodes)
                            {
                                XmlElement value = results.CreateElement("value");
                                value.SetAttribute("type", "table");
                                _ = value.AppendChild(tableDataXml);
                                _ = newElement.AppendChild(value);
                            }
                            else
                            {
                                _ = newElement.AppendChild(H.CreateElement(results, "value", "No data found in table."));
                            }
                        }
                        _ = newElement.AppendChild(H.CreateElement(results, "origin", $"{toolID}+{timestamp}"));
                        _ = newElement.AppendChild(H.CreateElement(results, "note", $"Value calculated"));
                    }
                    else
                    {
                        _ = newElement.AppendChild(H.CreateElement(results, "value", VariablesMap.Instance.GetVariableData(variableID).Default));
                        _ = newElement.AppendChild(H.CreateElement(results, "origin", $"{toolID}+{timestamp}"));
                        _ = newElement.AppendChild(H.CreateElement(results, "note", $"Value calculated"));
                    }
                    _ = varNode.AppendChild(newElement);
                }

            }

            // Close Excel
            workbook.Close(false);
            excelApp.Quit();

            try
            {
                _ = Marshal.ReleaseComObject(workbook);
                _ = Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ExcelCleanup", "Error releasing Excel objects: " + ex.Message);
            }

            // Additional cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return results;
        }

        public XmlDocument CalculateExcel(string toolID, DataMaster dm)
        {
            // 1. Retrieve the tool data by ID
            ToolData tool = ToolsMap.Instance.getToolDataByCode(toolID);
            if (tool == null)
                throw new ArgumentException($"ToolID '{toolID}' not found.");

            // 2. Check if the file type is Excel
            if (!(tool.FileType.Equals("xlsx", StringComparison.OrdinalIgnoreCase) || tool.FileType.Equals("xlsm", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The file is not an Excel type (.xlsx or .xlsm)");

            // 3. Create a MirrorXML instance
            MirrorXML mirror = new MirrorXML(toolID);

            // 4. Get the XXXX list from the mirror
            var variableMap = mirror.VarList;

            // 5. Build the full path to the Excel file
            string filePath = Path.Combine(
                tool.Resource == "TOOL" ? H.GetSProperty("ToolsPath") : H.GetSProperty("TemplatesPath"),
                tool.FileName
            );

            // 6. Crear copia del archivo para trabajar
            string newFilePath = Path.Combine(
                H.GetSProperty("processPath"),
                dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "",
                "TOOLS",
                tool.FileName
            );

            // Asegurar que la carpeta de destino existe antes de copiar
            _ = Directory.CreateDirectory(Path.GetDirectoryName(newFilePath)!);

            // Copiar y sobrescribir si ya existe
            File.Copy(filePath, newFilePath, true);

            // 7. Abrir el archivo en Excel Interop
            H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "CalculateExcel", $"Calculating tool {toolID}");

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application { Visible = false };
                workbook = excelApp.Workbooks.Open(newFilePath);

                // Escribir valores en las celdas
                foreach (var entry in variableMap)
                {
                    if (mirror.GetVarCallLevel(entry.Key) == tool.Call)
                    {
                        string variableID = entry.Key;
                        string direction = entry.Value[1];

                        if (direction == "in")
                        {
                            string rangeName = $"{H.GetSProperty("IN_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";
                            string type = _variablesMap.GetVariableData(variableID).Type;

                            if (type != "table")
                            {
                                Excel.Range cell = null;
                                try
                                {
                                    cell = workbook.Names.Item(rangeName).RefersToRange;
                                    if (cell == null)
                                        throw new Exception($"Named range '{rangeName}' not found in worksheet.");
                                    cell.Value = dm.GetValueString(variableID);
                                }
                                catch (Exception ex)
                                {
                                    H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Error accessing range: " + ex.Message);
                                }
                                finally
                                {
                                    if (cell != null) _ = Marshal.ReleaseComObject(cell);
                                }
                            }
                            else
                            {
                                XmlNode tableData = dm.GetValueXmlNode(variableID);
                                if (tableData.SelectSingleNode("t") != null && tableData.SelectSingleNode("t").HasChildNodes)
                                {
                                    WriteTableToExcel(workbook, rangeName, tableData);
                                }
                                else
                                {
                                    H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"No table data found for variable '{variableID}'.");
                                }
                            }
                        }
                    }
                }

                // Forzar cálculo de fórmulas en Excel
                excelApp.Calculate();
                workbook.Save();
                excelApp.Calculate();

                // Crear un nuevo documento XML para salida
                XmlDocument results = new XmlDocument();
                XmlElement root = results.CreateElement("root");
                _ = results.AppendChild(root);
                XmlElement varNode = results.CreateElement("variables");
                _ = root.AppendChild(varNode);

                // Obtener la fecha y hora actual en formato dd-HH:mm
                string timestamp = DateTime.Now.ToString("dd-HH:mm");

                foreach (var entry in variableMap)
                {
                    string variableID = entry.Key;
                    string direction = entry.Value[1];

                    if (mirror.GetVarCallLevel(variableID) == tool.Call && direction == "out")
                    {
                        XmlElement newElement = results.CreateElement(variableID);
                        string rangeName = $"{H.GetSProperty("OUT_VarPrefix")}{((tool.Call > 1) ? $"call{tool.Call}_" : "")}{variableID}";

                        Excel.Range cell = null;
                        try
                        {
                            cell = workbook.Names.Item(rangeName).RefersToRange;
                            if (cell.Value != null)
                            {
                                if (_variablesMap.GetVariableData(variableID).Type != "table")
                                {
                                    _ = newElement.AppendChild(H.CreateElement(results, "value", cell.Value.ToString()));
                                }
                                else
                                {
                                    XmlNode tableDataXml = ReadTableFromExcel(workbook, rangeName, results);
                                    if (tableDataXml != null && tableDataXml.HasChildNodes)
                                    {
                                        XmlElement value = results.CreateElement("value");
                                        value.SetAttribute("type", "table");
                                        _ = value.AppendChild(tableDataXml);
                                        _ = newElement.AppendChild(value);
                                    }
                                    else
                                    {
                                        _ = newElement.AppendChild(H.CreateElement(results, "value", "No data found in table."));
                                    }
                                }
                                _ = newElement.AppendChild(H.CreateElement(results, "origin", $"{toolID}+{timestamp}"));
                                _ = newElement.AppendChild(H.CreateElement(results, "note", "Value calculated"));
                            }
                            else
                            {
                                _ = newElement.AppendChild(H.CreateElement(results, "value", VariablesMap.Instance.GetVariableData(variableID).Default));
                                _ = newElement.AppendChild(H.CreateElement(results, "origin", $"{toolID}+{timestamp}"));
                                _ = newElement.AppendChild(H.CreateElement(results, "note", "Value calculated"));
                            }
                        }
                        catch (Exception ex)
                        {
                            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"❌Error❌ reading range '{rangeName}': {ex.Message}");
                        }
                        finally
                        {
                            if (cell != null) _ = Marshal.ReleaseComObject(cell);
                        }

                        _ = varNode.AppendChild(newElement);
                    }
                }

                return results;
            }
            finally
            {
                // Cierre y liberación de recursos COM
                if (workbook != null)
                {
                    workbook.Close(false);
                    _ = Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    _ = Marshal.ReleaseComObject(excelApp);
                }

                // 🧹 Limpieza de recursos no administrados
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void GenerateOuputWord(string templateID, DataMaster dm)
        {
            // 1. Retrieve the tool data by ID
            ToolData tool = ToolsMap.Instance.getToolDataByCode(templateID);
            if (tool == null)
                throw new ArgumentException($"ToolID '{templateID}' not found.");

            // 2. Check if the file type is Word
            if (!tool.FileType.Equals("docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The file is not a Word file (.docx)");

            // 3. Create a MirrorXML instance
            MirrorXML mirror = new MirrorXML(templateID);

            // 4. Get the variables list from the mirror
            var variableMap = mirror.VarList;

            // 5. Build the full path to the file
            string templatePath = Path.Combine(tool.Resource == "TEMPLATE" ? H.GetSProperty("TemplatesPath") : H.GetSProperty("ToolsPath"), tool.FileName);

            // 6. Crear copia del archivo para trabajar
            string filePath = Path.Combine(H.GetSProperty("processPath"), dm.GetValueString("opportunityFolder"), "OUTPUT", tool.FileName);

            // 7. Ensure the output directory exists and copy the template file if it doesn't exist
            if (!File.Exists(filePath))
            {
                if (!Directory.Exists(Path.GetDirectoryName(filePath)))
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            }

            try
            {
                File.Copy(templatePath, filePath, true);
            }
            catch (Exception)
            {
                H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "** Error - GenerateOutputWord", $"Cannot copy template {filePath}");
                throw;
            }

            // Confeccionamos la lista de marcas a reemplazar
            Dictionary<string, XmlNode> varList = new Dictionary<string, XmlNode>();


            List<string> varNotFound = new List<string>();

            foreach (string var in mirror.VarList.Keys)
            {
                try
                {
                    varList[var] = dm.GetValueXmlNode(var);
                }
                catch (KeyNotFoundException)
                {
                    varNotFound.Add(var);
                }
            }

            if (varNotFound.Count > 0) // If there are variables not found in the DataMaster stop process and print out missing variables
            {
                throw new Exception("keysNotFoundInDataMaster: " + string.Join(", ", varNotFound));
            }


            Word.Application wordApp = null;
            Word.Document doc = null;


            // 8. Open the Word document using Interop
            try
            {
                H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOutputWord", $"Generating output - Open file: {templateID}");
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(filePath, ReadOnly: false);

                string prefix = H.GetSProperty("VarPrefix");

                foreach (Word.Field field in doc.Fields) //each mark in the word document
                {
                    string variableID = "";

                    if (field.Type == WdFieldType.wdFieldRef && field.Code.Text.Trim().StartsWith(prefix)) // when the mark is an insert mark if (a)
                    {
                        variableID = field.Code.Text.Trim().Substring(prefix.Length); // ✅ removes the prefix from the fieldName and gets the variable id.

                        Microsoft.Office.Interop.Word.Range fieldRange = field.Result; // place to inserte found


                        if (field.Code.Text.Contains(variableID))
                        {
                            //ES TABLA
                            if (varList[variableID].SelectSingleNode("t") != null && varList[variableID].SelectSingleNode("t").HasChildNodes)
                            {
                                XmlNode tableNode = varList[variableID].SelectSingleNode("t");
                                if (tableNode == null)
                                {
                                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error - GenerateOuputWord", $"No valid table data {variableID} found in XML.");
                                }

                                // 📌 Count how many rows and columns the table has
                                XmlNodeList rows = tableNode.SelectNodes("r");
                                int rowCount = rows.Count;
                                int colCount = rows[0].ChildNodes.Count; // Assumes all rows have the same number of columns

                                // 📌 Insert a dynamically sized table
                                Word.Table table = doc.Tables.Add(fieldRange, rowCount, colCount);

                                for (int i = 0; i < rowCount; i++)
                                {
                                    XmlNodeList cells = rows[i].SelectNodes("c");
                                    for (int j = 0; j < colCount; j++)
                                    {
                                        table.Cell(i + 1, j + 1).Range.Text = cells[j].InnerText;
                                    }
                                }

                                // 📌 Apply "MyStyle" formatting
                                table.set_Style(H.GetSProperty("tableStyle"));

                                // 📌 Remove the reference mark after insertion
                                field.Delete();

                                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Tabla insertada y referencia '{variableID}' eliminada correctamente.");
                            }
                            //NO ES TABLA
                            else
                            {
                                fieldRange.Text = varList[variableID].InnerText;
                                field.Unlink(); // Convierte la referencia en texto estático
                            }
                        }

                    }
                }

                doc.Save();
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Reemplazo realizado con éxito.");
            }
            catch (Exception ex)
            {
                H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "** Error - GenerateOuputWord", $"❌Error❌ con el documento {Path.GetFileName(filePath)}");
                H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Error: " + ex.Message);
            }
            finally
            {
                // 🔒 Ensure Word document and app close cleanly
                if (doc != null)
                {
                    doc.Close(SaveChanges: false); // Ya se guardó arriba
                    _ = Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    _ = Marshal.ReleaseComObject(wordApp);
                }

                // 🧹 Clean up unmanaged resources
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "GenerateOutputWord", $"Generating output {templateID} finished");
        }

        static void WriteTableToExcel(Excel.Workbook workbook, string rangeName, XmlNode doc)
        {

            try
            {
                // 📌 Parse XML Input
                var rows = doc.SelectNodes("//t/r");
                int rowCount = rows.Count;
                int colCount = rows[0].ChildNodes.Count;

                // 📌 Write data into range
                Excel.Name namedRange = workbook.Names.Item(rangeName);
                Excel.Range inputRange = namedRange.RefersToRange;

                if (inputRange.Rows.Count != rowCount || inputRange.Columns.Count != colCount)
                {
                    H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Size mismatch: Input ({rowCount}x{colCount}) vs Range ({inputRange.Rows.Count}x{inputRange.Columns.Count}).");
                    return;
                }

                for (int i = 0; i < rowCount; i++)
                    for (int j = 0; j < colCount; j++)
                        ((Excel.Range)inputRange.Cells[i + 1, j + 1]).Value = rows[i].ChildNodes[j].InnerText;

            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Error writing table to Excel: " + ex.Message);
            }
        }

        static XmlElement ReadTableFromExcel(Excel.Workbook workbook, string rangeName, XmlDocument outputDoc)
        {

            try
            {
                // 📌 Read data from range
                Excel.Name namedRange = workbook.Names.Item(rangeName);
                Excel.Range outputRange = namedRange.RefersToRange;
                int outRows = outputRange.Rows.Count;
                int outCols = outputRange.Columns.Count;

                XmlElement root = outputDoc.CreateElement("t");

                for (int i = 1; i <= outRows; i++)
                {
                    XmlElement row = outputDoc.CreateElement("r");
                    for (int j = 1; j <= outCols; j++)
                    {
                        XmlElement cell = outputDoc.CreateElement("c");
                        cell.InnerText = Convert.ToString(((Excel.Range)outputRange.Cells[i, j]).Text); // ✅ Corrected conversion
                        _ = row.AppendChild(cell);
                    }
                    _ = root.AppendChild(row);
                }

                return root; // ✅ Returning Proper XMLNode Output

            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Error reading from Excel: " + ex.Message);
                return null;
            }
        }
    }// End of class ToolsMap
}// End of namespace SmartBid
