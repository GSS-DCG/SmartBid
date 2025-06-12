using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Lab
{
    public class MirrorXML
    {
        private string _fileName;
        private Dictionary<string, string> _varList;
        private Dictionary<string, string> _basicData;

        public Dictionary<string, string> VarList
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
            _varList = new Dictionary<string, string>();
            _basicData = new Dictionary<string, string>();

            LoadData(fileName);
        }

        private void LoadData(string target)
        {

            string? directoryPath = target.StartsWith("T")
                ? Path.GetDirectoryName(H.GetSProperty("ToolsPath"))
                : target.StartsWith("D")
                ? Path.GetDirectoryName(H.GetSProperty("TemplatesPath"))
                : ""; 

            FileName = target;
            target = Path.Combine(directoryPath, target);

            if (!File.Exists(target))
            {
                Console.WriteLine($"{target} does not exist.");
                return;
            }

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(target);
            string fileExtension = Path.GetExtension(target);
            string xmlFile = Path.Combine(directoryPath, fileNameWithoutExtension + ".xml");
            DateTime fileModified = File.GetLastWriteTime(target);
            DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;


            if (fileModified > xmlModified)
            {
                VarMap varMap = VarMap.Instance;

                if (fileExtension == ".docx")
                {
                    BasicData = ExtractGSSDataFromDocx(target);
                    VarList = ExtractVariablesFromDocx(target);
                }
                else if (fileExtension == ".xlsx")
                {
                    ExcelReader xlsFile = new ExcelReader(target);
                    BasicData = ExtractGSSDataFromXlsx(xlsFile);
                    VarList = ExtractVariablesFromXlsx(xlsFile);
                }

                foreach (string variable in VarList.Keys.ToList())
                {
                    //var variableData = varMap.GetVariableData(variable);
                    VarList[variable] = varMap.GetVariableData(variable).Source; //Fill up source data before saving to XML

                }

                CreateXMLMirror(directoryPath);
            }

            if (File.Exists(xmlFile))
            {
                XElement root = XElement.Load(xmlFile);
                root.Attributes().ToList().ForEach(attr => _basicData[attr.Name.LocalName] = attr.Value);

                foreach (var variable in root.Elements("variable"))
                {
                    VarList[variable.Value] = variable?.Attribute("source")?.Value ?? "";
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
                Console.WriteLine($"File not found: {docxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            bookmarks = bookmarks.Where(kvp => kvp.Key.StartsWith("GSS_"))
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            return bookmarks;
        }

        private Dictionary<string, string> ExtractVariablesFromDocx(string docxPath)
        {
            Dictionary<string, string> varList = new Dictionary<string, string>();
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
                                varList[currentField] = "";

                                // Reset currentField for the next varName
                                currentField = "";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading DOCX: {ex.Message}");
            }
            return varList
                .Where(pair => !string.IsNullOrWhiteSpace(pair.Key) && !pair.Key.ToLower().Contains("_toc")) // Apply filter to key
                .Select(pair => new KeyValuePair<string, string>( // Modify key, preserve value
                    pair.Key.Replace("\\* MERGEFORMAT", "").Replace("ref ", ""),
                    pair.Value))
                .DistinctBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase) // Ensure unique keys
                .ToDictionary(pair => pair.Key, pair => pair.Value); // Convert back to Dictionary


        }

        private static Dictionary<string, string> ExtractGSSDataFromXlsx(ExcelReader file)
        {
            var gssData = new Dictionary<string, string>();
            var listData = file.GetCellValuesFromRange("GSS_DATA");

            foreach (var row in listData)
            {
                string key = row[0];
                string value = row[1];
                gssData[key] = value;
            }
            return gssData;
        }

        private static Dictionary<string, string> ExtractVariablesFromXlsx(ExcelReader file)
        {
            Dictionary<string, string> varList = new Dictionary<string, string>();
            var varNames = file.GetAllRangeNames();

            foreach (string name in varNames)
            {
                if (name.StartsWith("APG_"))
                {
                    varList[name.Trim()] = "";
                }
            }
            return varList;
        }

        private void CreateXMLMirror(string directory)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
            var outputPath = Path.Combine(directory, fileNameWithoutExtension + ".xml");

            ToXMLDocument().Save(outputPath);
            Console.WriteLine($"XML file created at: {outputPath}");

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
                varElement.Add(new XAttribute("source", VarList[varName]));
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

    }

    public class ExcelReader
    {
        private WorkbookPart workbookPart;

        // Constructor that takes a fileName argument
        public ExcelReader(string fileName)
        {

            SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false);
            this.workbookPart = document.WorkbookPart;

        }

        public List<string> GetAllRangeNames()
        {
            List<string> rangeNames = new List<string>();
            try
            {
                var definedNames = this.workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    foreach (var definedName in definedNames.Elements<DefinedName>())
                    {
                        rangeNames.Add(definedName.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading range names: {ex.Message}");
            }
            return rangeNames;
        }

        public List<List<string>> GetCellValuesFromRange(string rangeName)
        {
            List<List<string>> cellValues = new List<List<string>>();
            try
            {
                var definedNames = this.workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var gssInputRange = definedNames.Elements<DefinedName>().FirstOrDefault(dn => dn.Name == rangeName);
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
                            Console.WriteLine($"Sheet {sheetName} not found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Range {rangeName} not found.");
                    }
                }
                else
                {
                    Console.WriteLine("No defined names found.");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading {rangeName} range: {ex.Message}");
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

    }

}
