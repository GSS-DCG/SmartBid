using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Navigation;
using iText.Kernel.Utils;
using iText.Kernel.XMP;
using iText.Kernel.XMP.Options;
using iText.Kernel.XMP.Properties;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace SmartBid
{
    class Auxiliar
    {
        #region Gestionar proceso Word
        /// <summary>
        /// Fuerza el cierre de el proceso de word en el arbol de procesos
        /// </summary>
        public static void CloseWord()
        {
            Process[] wordProcesses = Process.GetProcessesByName("WINWORD");

            if (WordAppDetection(wordProcesses))
            {
                CerrarProcesos:
                Console.WriteLine("Existen procesos de WordAbiertos desea cerrarlos (s/n): ");
                string var = Console.ReadLine();
                if (var == null)
                {
                    goto CerrarProcesos;
                }
                else if (var == "s")
                {
                    WordAppClose(wordProcesses);
                }
                else if (var == "n") { }
                else
                {
                    Console.WriteLine("Argumento no valido.");
                    goto CerrarProcesos;
                }
            }
        }

        private static bool WordAppDetection(Process[] wordProcesses)
        {
            if (wordProcesses.Length == 0)
            {
                return false;
            }

            return true;
        }

        private static void WordAppClose(Process[] wordProcesses)
        {
            foreach (Process proc in wordProcesses)
            {
                try
                {
                    H.PrintLog(2, "Main", "WordProcess:", @$"⚠️Cerrando proceso Excel (ID: {proc.Id})...");
                    proc.Kill();
                    proc.WaitForExit();
                }
                catch (Exception ex)
                {
                    H.PrintLog(2, "Main", "WordProcess:", @$"❌Error❌ al cerrar Excel: { ex.Message}");
                }
            }
        }
        #endregion

        #region Gestionar proceso Excel
        /// <summary>
        /// Fuerza el cierre de el proceso de excel en el arbol de procesos
        /// </summary>
        public static void CloseExcel()
        {
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");

            if (ExcelAppDetection(excelProcesses))
            {
            CerrarProcesos:
                Console.WriteLine("Existen procesos de Excel abiertos desea cerrarlos (s/n): ");
                string var = Console.ReadLine();

                if (var == null)
                {
                    goto CerrarProcesos;
                }
                else if (var == "s")
                {
                    ExcelAppClose(excelProcesses);
                }
                else if (var == "n") { }
                else
                {
                    Console.WriteLine("Argumento no valido.");
                    goto CerrarProcesos;
                }
            }
        }

        private static bool ExcelAppDetection(Process[] excelProcesses)
        {
            if (excelProcesses.Length == 0)
            {
                return false;
            }

            return true;
        }

        private static void ExcelAppClose(Process[] excelProcesses)
        {
            foreach (Process proc in excelProcesses)
            {
                try
                {
                    H.PrintLog(2, "Main", "ExcelProcess:", @$"⚠️Cerrando proceso excel (ID: {proc.Id})...");
                    proc.Kill();
                    proc.WaitForExit();
                }
                catch (Exception ex)
                {
                    H.PrintLog(2, "Main", "ExcelProcess:", @$"❌Error❌ al cerrar excel: {ex.Message}");
                }
            }
        }
        #endregion

        #region Word to Pdf
            
        public static bool wordToPdf(Document wordDoc, string outputPath)
        {
            try
            {
                wordDoc.ExportAsFixedFormat(outputPath, WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch
            {
                H.PrintLog(2, "Main", "wordToPdf:", @$"❌Error❌ al Convertir el archivo {wordDoc.Name} a PDF.");
                return false;
            }
        }

        #endregion

        #region Gestionar Marcadores

        public static void DeleteBookmarkLoop(string DocName, Word.Document doc, Dictionary<string, string> OpcionesHerramienta,string prefix)
        {
            string BookmarkCtrl = "";
            DocName = DocName.Replace(" ", "");
            try
            {
                List<string> marcadores = new List<string>();

                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    marcadores.Add(bookmark.Name);
                }

                foreach (string BookmarkName in marcadores)
                {
                    BookmarkCtrl = BookmarkName;

                    if (CalculateSimilarity(EliminarDiacriticos(BookmarkName.Replace(prefix, "").ToLower()), EliminarDiacriticos(OpcionesHerramienta[DocName].ToLower())) < 0.95)
                    {
                        if (!Regex.IsMatch(BookmarkName, @"^S\d+$"))
                        {
                            H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value.User, "DeleteBookmarkText:", $"Levenstein: {CalculateSimilarity(EliminarDiacriticos(BookmarkName.Replace(prefix,"").ToLower()), EliminarDiacriticos(OpcionesHerramienta[DocName].ToLower()))} " +
                                $"BookmarkName: {EliminarDiacriticos(BookmarkName.Replace(prefix, "").ToLower())} " +
                                $"OpcionHerramienta: {EliminarDiacriticos(OpcionesHerramienta[DocName].ToLower())}");

                            // Eliminar el texto dentro del marcador
                            Word.Bookmark bookmark = doc.Bookmarks[BookmarkName];
                            Word.Range range = bookmark.Range;

                            //Borrar el marcador
                            bookmark.Delete();

                            range.Text = "";
                        }
                        else
                        {
                            try
                            {
                                H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value.User, "DeleteBookmarkText:", $"Marcador de seccion {BookmarkName}");
                                // Eliminar el texto dentro del marcador
                                Word.Bookmark bookmark = doc.Bookmarks[BookmarkName];
                                Word.Range range = bookmark.Range;

                                //Borrar el marcador
                                bookmark.Delete();

                                range.Text = "";
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Eliminar el texto dentro del marcador
                        Word.Bookmark bookmark = doc.Bookmarks[BookmarkName];

                        //Borrar el marcador
                        bookmark.Delete();
                    }
                }  
            }

            catch (Exception e)
            {
                H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "DeleteBookmarkText:", @$"❌Error❌ al Eliminar los marcadores: {DocName}. Marcador: {BookmarkCtrl}");
            }
            finally
            {
                doc.Fields.Update();
            }
        }

        public static void DeleteBookmarkText(string DocName, string BookmarkName, DataMaster dm, string SubCarpeta)
        {
            CloseWord();

            string RutaBaseWordDoc = $@"{H.GetSProperty("TemplatesPath")}{DocName}";
            string RutaProcessedWordDoc = Path.Combine(H.GetSProperty("storagePath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/projectFolder")?.InnerText ?? "", $"{SubCarpeta}\\{DocName}");

            try
            {
                File.Copy(RutaBaseWordDoc, RutaProcessedWordDoc, true);
                _ = _DeleteBookmarkText(RutaProcessedWordDoc, BookmarkName);
                _ = _DeleteBookmarks(RutaProcessedWordDoc, BookmarkName);
            }
            catch (Exception e)
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "DeleteBookmarkText:", @$"❌Error❌ al Editar: {RutaProcessedWordDoc}");
            }
        }

        private static bool _DeleteBookmarks(string RutaWord, string BookmarkName)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(RutaWord, Visible: false);

            // Acceder al marcador
            Word.Bookmark bookmark = doc.Bookmarks[BookmarkName];

            //Borrar el marcador
            bookmark.Delete();

            // Guardar y cerrar el documento
            doc.Save();
            doc.Close();
            wordApp.Quit();

            return true;
        }

        private static bool _DeleteBookmarkText(string RutaWord, string BookmarkName)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(RutaWord, Visible: false);

            // Acceder al marcador
            Word.Bookmark bookmark = doc.Bookmarks[BookmarkName];
            Microsoft.Office.Interop.Word.Range range = bookmark.Range;

            // Eliminar el texto dentro del marcador
            range.Text = "";

            // Guardar y cerrar el documento
            doc.Save();
            doc.Close();
            wordApp.Quit();

            return true;
        }

        #endregion

        #region Gestion de opciones de una herramienta
        /// <summary>
        /// Obtiene los nombres de las distitnas opciones de una herramienta con prefijo "sbOP_"
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetOptionValue(Excel.Workbook workbook)
        {
            Dictionary<string, string> option = new Dictionary<string, string>();
            try
            {
                string optionPref = H.GetSProperty("OptionPrefix");
                foreach (Excel.Name name in workbook.Names)
                {
                    if (name.Name.StartsWith(optionPref))
                    {
                        Excel.Range range = name.RefersToRange;
                        if (range != null)
                        {
                            string value = range.Text.ToString();
                            string Name = name.Name.Substring(optionPref.Length);
                            option[Name] = value.Replace(" ","");
                        }
                    }
                }
            }
            catch (Exception ex) 
            {
                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "GetOptionValue", "No se han podido buscar las opciones de la herramienta: " + ex.Message);
            }
            return option;
        }


        #endregion

        #region Calculamos similitud entre texto

        /// <summary>
        /// Calcula el porcentaje de similitud entre dos cadenas de texto utilizando el algoritmo de distancia de Levenshtein.
        /// </summary>
        /// <param name="source">La primera cadena a comparar.</param>
        /// <param name="target">La segunda cadena a comparar.</param>
        /// <returns>Un valor double entre 0.0 y 1.0 que representa el porcentaje de similitud.</returns>
        public static double CalculateSimilarity(string source, string target)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
            {
                return 0.0;
            }
            if (source == target)
            {
                return 1.0;
            }
   
            int distance = LevenshteinDistance(source, target);
            double maxLength = Math.Max(source.Length, target.Length);
   
            // Normalizamos la distancia para obtener un ratio de similitud
            return 1.0 - (distance / maxLength);
        }
   
        /// <summary>
        /// Implementación del algoritmo de distancia de Levenshtein.
        /// </summary>
        private static int LevenshteinDistance(string source, string target)
        {
            int n = source.Length;
            int m = target.Length;
            int[,] d = new int[n + 1, m + 1];
   
            if (n == 0)
            {
                return m;
            }

            if (m == 0)
            {
                return n;
            }

            for (int i = 0; i <= n; d[i, 0] = i++) { }
            for (int j = 0; j <= m; d[0, j] = j++) { }

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
        
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            return d[n, m];
        }

        #endregion

        #region Eliminar diacriticos

        /// <summary>
        /// Elimina los caracteres diacríticos (acentos, diéresis, etc.) de la cadena.
        /// </summary>
        /// <returns>La cadena sin diacríticos.</returns>
        public static string EliminarDiacriticos( string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
            {
                return texto;
            }
   
            string textoNormalizado = texto.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
   
            foreach (var c in textoNormalizado)
            {
                var categoriaUnicode = CharUnicodeInfo.GetUnicodeCategory(c);
                if (categoriaUnicode != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        #endregion

        #region Dividir PDF

        public static void DividirPdfPorIndice(string rutaPdf)
        {
            string carpetaSalida = Path.GetDirectoryName(rutaPdf);
            using var pdfDoc = new PdfDocument(new PdfReader(rutaPdf));

            IList<PdfOutline> outlines = pdfDoc.GetOutlines(false).GetAllChildren();

            for (int i = 0; i < outlines.Count; i++)
            {
                var outline = outlines[i];
                string titulo = SanearNombre(outline.GetTitle());

                int paginaInicio = ObtenerNumeroPaginaDesdeDestino(pdfDoc, outline.GetDestination());
                int paginaFin = (i + 1 < outlines.Count)
                    ? ObtenerNumeroPaginaDesdeDestino(pdfDoc, outlines[i + 1].GetDestination()) - 1
                    : pdfDoc.GetNumberOfPages();

                string rutaSalida = Path.Combine(carpetaSalida, $"{titulo}.pdf");

                ExtraerPaginas(rutaPdf, rutaSalida, paginaInicio, paginaFin);
            }
        }

        private static int ObtenerNumeroPaginaDesdeDestino(PdfDocument pdfDoc, PdfDestination destino)
        {
            if (destino == null) return 1;

            PdfObject pageRef = destino.GetPdfObject();

            if (pageRef.IsIndirectReference())
            {
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    if (pdfDoc.GetPage(i).GetPdfObject().GetIndirectReference().Equals(pageRef))
                    {
                        return i;
                    }
                }
            }

            return 1; // fallback
        }

        private static void ExtraerPaginas(string rutaOrigen, string rutaDestino, int paginaInicio, int paginaFin)
        {
            using var pdfReader = new PdfReader(rutaOrigen);
            using var pdfWriter = new PdfWriter(rutaDestino);
            using var pdfDocOrigen = new PdfDocument(pdfReader);
            using var pdfDocNuevo = new PdfDocument(pdfWriter);
            var merger = new PdfMerger(pdfDocNuevo);

            merger.Merge(pdfDocOrigen, paginaInicio, paginaFin);
        }

        private static string SanearNombre(string nombre)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                nombre = nombre.Replace(c, '_');
            }
            return nombre.Trim();
        }
    }
    #endregion
}