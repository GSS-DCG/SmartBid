
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace SmartBid
{
    public class MirrorXML
    {
        private string _fileName;
        private Dictionary<string, string[]> _varList;
        private Dictionary<string, string> _basicData;

        public Dictionary<string, string[]> VarList
        {
            get => _varList;
            set => _varList = value;
        }

        public Dictionary<string, string> BasicData
        {
            get => _basicData;
            set => _basicData = value;
        }

        public string FileName
        {
            get => _fileName;
            set => _fileName = value;
        }

        public MirrorXML(string fileName)
        {
            _varList = new Dictionary<string, string[]>();
            _basicData = new Dictionary<string, string>();

            LoadData(fileName);
        }

        private void LoadData(string target)
        {
            target = Regex.Replace(target, "_Call\\d+", "");

            string? directoryPath = (ToolsMap.Instance.getToolDataByCode(target).Resource == "TOOL")
                ? Path.GetDirectoryName(H.GetSProperty("ToolsPath"))
                : (ToolsMap.Instance.getToolDataByCode(target).Resource == "TEMPLATE")
                ? Path.GetDirectoryName(H.GetSProperty("TemplatesPath"))
                : "";

            FileName = ToolsMap.Instance.getToolDataByCode(target).FileName;
            target = Path.Combine(directoryPath, FileName);

            if (!File.Exists(target))
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"{target} does not exist.");
                return;
            }

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(target);
            string fileExtension = Path.GetExtension(target);
            string xmlFile = Path.Combine(directoryPath, fileNameWithoutExtension + ".xml");
            DateTime fileModified = File.GetLastWriteTime(target);
            DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;

            if (fileModified > xmlModified)
            {
                VariablesMap varMap = VariablesMap.Instance;

                if (fileExtension.ToLower() == ".docx")
                {
                    BasicData = ExtractGSSDataFromDocx(target);
                    VarList = ExtractVariablesFromDocx(target);
                }
                else if (fileExtension.ToLower() == ".xlsx")
                {
                    BasicData = ExtractGSSDataFromXlsx(target);
                    VarList = ExtractVariablesFromXlsx(target);
                }

                foreach (string variable in VarList.Keys.ToList())
                {
                    if (!variable.Contains("\\s"))
                    {
                        string id = System.Text.RegularExpressions.Regex.Replace(variable, @"^Call[0-9]_", "");

                        if(varMap.GetNewVariableData(id) != null)
                        {
                            VarList[variable] = new string[] { varMap.GetNewVariableData(id).Source, VarList[variable][1], VarList[variable][2] }; // Fill up source data before saving to XML
                        }
                    }
                }

                CreateXMLMirror(directoryPath);
            }

            if (File.Exists(xmlFile))
            {
                XElement root = XElement.Load(xmlFile);
                root.Attributes().ToList().ForEach(attr => _basicData[attr.Name.LocalName] = attr.Value);

                foreach (var variable in root.Elements("variable"))
                {
                    VarList[variable.Value] = new string[] { variable?.Attribute("source")?.Value ?? "", variable?.Attribute("inOut")?.Value ?? "", variable?.Attribute("call")?.Value ?? "" };
                }
            }
        }

        static Dictionary<string, string> ExtractGSSDataFromDocx(string docxPath)
        {
            Dictionary<string, string> bookmarks = new Dictionary<string, string>();

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var bookmarksList = body.Descendants<BookmarkStart>();

                    foreach (var bookmark in bookmarksList)
                    {
                        string bookmarkName = bookmark.Name;
                        string bookmarkValue = "";

                        // Buscar el siguiente nodo que sea un `Run`
                        var currentElement = bookmark.NextSibling();
                        while (currentElement != null)
                        {
                            var run = currentElement as Run;
                            if (run != null)
                            {
                                var textElements = run.Descendants<Text>();
                                foreach (var textElement in textElements)
                                {
                                    bookmarkValue += textElement.Text;
                                }
                            }
                            currentElement = currentElement.NextSibling();
                        }

                        bookmarks[bookmarkName] = bookmarkValue.Trim();
                    }
                }
            }
            catch (FileNotFoundException)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"File not found: {docxPath}");
            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"An error occurred: {ex.Message}");
            }

            bookmarks = bookmarks.Where(kvp => kvp.Key.StartsWith("GSS_"))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            return bookmarks;
        }

        private Dictionary<string, string[]> ExtractVariablesFromDocx(string docxPath)
        {
            Dictionary<string, string[]> varList = new Dictionary<string, string[]>();
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var runs = body.Descendants<Run>();
                    string currentField = "";
                    bool isFieldActive = false;
                    foreach (var run in runs)
                    {
                        var fldCharBegin = run.Descendants<FieldChar>().FirstOrDefault(fc => fc.FieldCharType == FieldCharValues.Begin);
                        var fldCharEnd = run.Descendants<FieldChar>().FirstOrDefault(fc => fc.FieldCharType == FieldCharValues.End);
                        var instrText = run.Descendants<FieldCode>().FirstOrDefault();
                        if (fldCharBegin != null)
                        {
                            isFieldActive = true;
                            currentField = "";
                        }
                        if (isFieldActive && instrText != null && !string.IsNullOrEmpty(instrText.InnerText))
                        {
                            currentField += instrText.InnerText.Trim();
                        }
                        if (fldCharEnd != null)
                        {
                            isFieldActive = false;
                            if (!string.IsNullOrEmpty(currentField))
                            {
                                // Remove unwanted characters and add to the list
                                currentField = currentField.Trim();
                                varList[currentField] = new string[] { "", "in", "1" };

                                // Reset currentField for the next varName
                                currentField = "";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Error reading DOCX: {ex.Message}");
            }
            string varPrefix = H.GetSProperty("VarPrefix");
            return varList
                .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && !pair.Key.ToLower().Contains("_toc")) // Apply filter to key
                .Select(pair => new KeyValuePair<string, string[]>(pair.Key.Replace("\\* MERGEFORMAT", "").Replace("ref ", "").Replace(varPrefix, ""), pair.Value)) // Modify key, preserve value
                .DistinctBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase) // Ensure unique keys
                .ToDictionary(pair => pair.Key, pair => pair.Value); // Convert back to Dictionary

        }

        private static Dictionary<string, string> ExtractGSSDataFromXlsx(string fileName)
        {
            var gssData = new Dictionary<string, string>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                var listData = GetCellValuesFromRange(workbookPart, "GSS_DATA");

                foreach (var row in listData)
                {
                    string key = row[0];
                    string value = row[1];
                    gssData[key] = value;
                }
            }
            return gssData;
        }

        private static Dictionary<string, string[]> ExtractVariablesFromXlsx(string fileName)
        {
            Dictionary<string, string[]> varList = new Dictionary<string, string[]>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                var varNames = GetAllRangeNames(workbookPart);

                string inPrefix = H.GetSProperty("IN_VarPrefix");
                string outPrefix = H.GetSProperty("OUT_VarPrefix");

                _ = varNames.Remove("GSS_DATA"); // Remove GSS_DATA from the list
                foreach (string item in varNames)
                {
                    string varName = item;
                    string[] value = new string[3] { "", "", "1" };

                    if (varName.ToLower().StartsWith("sbin_"))
                    {
                        value[1] = "in";
                        varName = varName.Substring(5);
                    }
                    else if (varName.ToLower().StartsWith("sbout_"))
                    {
                        value[1] = "out";
                        varName = varName.Substring(6);
                    }

                    Match match = Regex.Match(varName.ToLower(), @"^call(\d)_");
                    if (match.Success) value[2] = match.Groups[1].Value; varName = varName.Substring(match.Length);

                    if (!varList.ContainsKey(varName))
                    {
                        varList.Add(new string(varName), value);
                    }
                    else
                    {
                        Console.WriteLine($"VariableData with ID '{varName}' was declared twice on the word Document.");
                    }
                }
            }
            return varList;


        }

        private static List<string> GetAllRangeNames(WorkbookPart workbookPart)
        {
            List<string> rangeNames = new List<string>();
            try
            {
                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    foreach (var definedName in definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>())
                    {
                        rangeNames.Add(definedName.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Error reading range names: {ex.Message}");
            }
            return rangeNames;
        }

        private static List<List<string>> GetCellValuesFromRange(WorkbookPart workbookPart, string rangeName)
        {
            List<List<string>> cellValues = new List<List<string>>();
            try
            {
                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var gssInputRange = definedNames.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().FirstOrDefault(dn => dn.Name == rangeName);
                    if (gssInputRange != null)
                    {
                        string[] range = gssInputRange.Text.Split('!')[1].Split(':');
                        string sheetName = gssInputRange.Text.Split('!')[0].Trim('\'');
                        Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                        if (sheet != null)
                        {
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                            string startCellReference = range[0].Replace("$", ""); // Remove dollar signs
                            string endCellReference = (range.Length == 1) ? startCellReference : range[1].Replace("$", ""); // Remove dollar signs

                            var cells = worksheetPart.Worksheet.Descendants<Cell>()
                                .Where(c => IsCellInRange(c.CellReference, startCellReference, endCellReference));

                            int numberOfColumns = GetColumnRowIndices(endCellReference).column - GetColumnRowIndices(startCellReference).column + 1;
                            int numberOfRows = cells.Count() / numberOfColumns;
                            int i = 0;
                            List<string> row = new List<string>();
                            foreach (var cell in cells)
                            {
                                i++;
                                if (i <= numberOfColumns)
                                {
                                    row.Add(GetCellValue(workbookPart, cell));
                                }
                                else
                                {
                                    cellValues.Add(row); // Add the row to the main list
                                    i = 1;
                                    row = new List<string>();
                                    row.Add(GetCellValue(workbookPart, cell));
                                }
                            }
                            cellValues.Add(row); // Add the row to the main list
                        }
                        else
                        {
                            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Sheet {sheetName} not found.");
                        }
                    }
                    else
                    {
                        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Range {rangeName} not found.");
                    }
                }
                else
                {
                    H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "No defined names found.");
                }
            }
            catch (Exception ex)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"Error reading {rangeName} range: {ex.Message}");
            }
            return cellValues;
        }

        private static bool IsCellInRange(string cellReference, string startCellReference, string endCellReference)
        {
            // Convert cell references to row and column indices
            (int startColumn, int startRow) = GetColumnRowIndices(startCellReference);
            (int endColumn, int endRow) = GetColumnRowIndices(endCellReference);
            (int cellColumn, int cellRow) = GetColumnRowIndices(cellReference);

            // Check if the cell is within the specified range
            return cellColumn >= startColumn && cellColumn <= endColumn && cellRow >= startRow && cellRow <= endRow;
        }

        private static (int column, int row) GetColumnRowIndices(string cellReference)
        {
            // Extract column letters and row numbers from the cell reference
            string columnLetters = new string(cellReference.Where(char.IsLetter).ToArray());
            string rowNumbers = new string(cellReference.Where(char.IsDigit).ToArray());

            // Convert column letters to column index (A=1, B=2, ..., Z=26, AA=27, etc.)
            int columnIndex = 0;
            foreach (char letter in columnLetters)
            {
                columnIndex = (columnIndex * 26) + letter - 'A' + 1;
            }

            // Convert row numbers to row index
            int rowIndex = int.Parse(rowNumbers);

            return (columnIndex, rowIndex);
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

        private void CreateXMLMirror(string directory)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
            var outputPath = Path.Combine(directory, fileNameWithoutExtension + ".xml");

            ToXMLDocument().Save(outputPath);
            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", $"XML file created at: {outputPath}");

            ToXMLDocument().Load(outputPath);
        }

        public XmlDocument ToXMLDocument()
        {
            XElement docElement = new XElement("doc");
            docElement.Add(new XAttribute("fileName", FileName));

            foreach (string key in BasicData.Keys)
                docElement.Add(new XAttribute(key, BasicData.TryGetValue(key, out var keyValue) ? keyValue : ""));

            docElement.Add(new XAttribute("date", DateTime.Now.ToString("MM/dd/yy_HH:mm")));

            foreach (string varName in VarList.Keys.ToList())
            {
                XElement varElement = new XElement("variable", varName);
                varElement.Add(new XAttribute("source", VarList[varName][0]));
                varElement.Add(new XAttribute("inOut", VarList[varName][1]));
                if (VarList[varName].Length == 3) varElement.Add(new XAttribute("call", VarList[varName][2]));

                docElement.Add(varElement);
            }

            // Create an XDocument with declaration
            XDocument xDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), docElement);

            // Convert XDocument to XmlDocument
            XmlDocument xmlDoc = new XmlDocument();
            using (var reader = xDoc.CreateReader())
            {
                xmlDoc.Load(reader);
            }

            return xmlDoc;
        }

        public int GetVarCallLevel(string varName)
        {
            if (VarList.ContainsKey(varName)) return int.Parse(_varList[varName][2]);
            return 0;
        }
    }
}
