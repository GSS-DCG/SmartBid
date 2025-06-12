using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;



namespace Lab
{
    public class Program1
    {
        static void _Main(string[] args)
        {
            string rutaActual = "C:\\InSync\\Lab\\files\\";
            string nombreDocx = "ES_Informe-Demo_000.00.docx";
            string docxFile = Path.Combine(rutaActual, nombreDocx);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(nombreDocx);
            string xmlFile = Path.Combine(rutaActual, fileNameWithoutExtension + ".xml");

            Console.WriteLine(docxFile);

            if (File.Exists(docxFile))
            {
                DateTime docxModified = File.GetLastWriteTime(docxFile);
                bool shouldCreateXml = true;

                if (File.Exists(xmlFile))
                {
                    DateTime xmlCreated = File.GetCreationTime(xmlFile);
                    shouldCreateXml = docxModified > xmlCreated;
                }

                if (shouldCreateXml)
                {
                    CreateXMLdocMirror(docxFile);
                }
                else
                {
                    Console.WriteLine("El archivo XML ya está actualizado.");
                }
            }
            else
            {
                Console.WriteLine("El archivo DOCX no existe.");
            }
        }



        static List<string> ExtractAllInstrText(string docxPath, bool ifPrint = false)
        {
            List<string> fields = new List<string>();

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
                            currentField = ""; // Reiniciar acumulación al inicio de un nuevo campo
                        }

                        if (isFieldActive && instrText != null && !string.IsNullOrEmpty(instrText.InnerText))
                        {
                            currentField += instrText.InnerText.Trim(); // Concatenar fragmentos dentro del campo activo
                        }

                        if (fldCharEnd != null)
                        {
                            isFieldActive = false;
                            if (!string.IsNullOrEmpty(currentField)) // Si hay contenido acumulado, guardarlo
                            {
                                fields.Add(currentField.Trim());
                                currentField = ""; // Reiniciar acumulación
                            }
                        }
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

            // Clean and filter the fields
            fields = fields.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            fields = fields.Where(item => !item.ToLower().Contains("_toc")).ToList();
            fields = fields.Select(item => item.Replace("\\* MERGEFORMAT", "")).ToList();
            fields = fields.Select(item => item.Replace("ref ", "")).ToList();
            fields = fields.Distinct(StringComparer.OrdinalIgnoreCase).ToList();


            return fields;
        }

        static Dictionary<string, string> ExtractBookmarks(string docxPath)
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
                                var textElement = run.Descendants<Text>().FirstOrDefault();
                                if (textElement != null)
                                {
                                    bookmarkValue = textElement.Text;
                                    break; // Terminamos cuando encontramos el texto
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

            return bookmarks;
        }

        static void CreateXMLdocMirror(string docxPath)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(docxPath);
            var directory = Path.GetDirectoryName(docxPath);
            var outputPath = Path.Combine(directory, fileNameWithoutExtension + ".xml");

            var bookmarks = ExtractBookmarks(docxPath);
            var fields = ExtractAllInstrText(docxPath);

            string gssCode = bookmarks.ContainsKey("GSS_GSSCcode") ? bookmarks["GSS_GSSCcode"] : "";
            string docVersion = bookmarks.ContainsKey("GSS_DocVersion") ? bookmarks["GSS_DocVersion"] : "";
            string fecha = DateTime.Now.ToString("MM/dd/yy_HH:mm");

            XElement docElement = new XElement("doc",
                new XAttribute("fileName", fileNameWithoutExtension),
                new XAttribute("GSSCode", gssCode),
                new XAttribute("DocVersion", docVersion),
                new XAttribute("fecha", fecha)
            );


            foreach (var field in fields)
            {
                docElement.Add(new XElement("variable", field));
            }

            XDocument xmlDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), docElement);
            xmlDoc.Save(outputPath);
            Console.WriteLine($"XML file created at: {outputPath}");
        }

        static void CrearVariableMapXML(List<(string, string)> variablesList)
        {
            // Create root element
            XElement root = new XElement("root");
            // Create 'documento' element
            XElement documento = new XElement("documento");
            documento.SetAttributeValue("CodigoDocumento", "ES_DEMO_000.00");
            documento.SetAttributeValue("NombreDocumento", "Informe Demo");
            documento.SetAttributeValue("Version", "000.00");
            documento.SetAttributeValue("Idioma", "ES");
            root.Add(documento);
            // Create 'Variables' element
            XElement variables = new XElement("Variables");
            documento.Add(variables);
            // Populate variables
            foreach (var (identificador, valorPorDefecto) in variablesList)
            {
                XElement variable = new XElement("Variable", new XAttribute("IDENTIFICADOR", identificador));
                if (!string.IsNullOrEmpty(valorPorDefecto)) // Add 'VALOR_POR_DEFECTO' only if there is a value
                {
                    XElement valorElement = new XElement("VALOR_POR_DEFECTO", valorPorDefecto);
                    variable.Add(valorElement);
                }
                variables.Add(variable);
            }
            // Save XML to a file
            root.Save("output.xml");
            Console.WriteLine("XML file created successfully!");
        }
    }
}