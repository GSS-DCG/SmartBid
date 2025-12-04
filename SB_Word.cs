using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Interop.Word;


namespace SmartBid
{
  public class SB_Word
  {

    private Microsoft.Office.Interop.Word.Application? wordApp;
    private Document doc;
    private string filePath;

    public SB_Word(string filePath)
    {
      this.filePath = filePath;
      wordApp = new Microsoft.Office.Interop.Word.Application();
      doc = wordApp.Documents.Open(filePath, ReadOnly: false, Visible: false); ;
    }
    public void DeleteBookmarks(List<string> removeBkm)
    {
      string prefix = H.GetSProperty("VarPrefix").ToLower();
      string removePrefix = H.GetSProperty("RemoveBkmPrefix").ToLower();

      removeBkm = removeBkm.Select(b => (prefix + b).ToLower()).ToList();

      Dictionary<string, string> bookmarkDict = doc.Bookmarks.Cast<Bookmark>()
          .ToDictionary(b => b.Name.ToLower(), b => b.Name);

      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.DeleteBookmarks", "Lista de bookmarks:");
      foreach (var kvp in bookmarkDict)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.DeleteBookmarks", kvp.Value);
      }

      var orderedBookmarks = bookmarkDict.Keys
          .OrderByDescending(name => name.StartsWith(removePrefix))
          .ToList();

      foreach (string bookmarkName in orderedBookmarks)
      {
        if (!removeBkm.Contains(bookmarkName) && !bookmarkName.StartsWith(removePrefix))
          continue;

        if (!bookmarkDict.ContainsKey(bookmarkName))
          continue;

        try
        {
          Bookmark bookmark = doc.Bookmarks[bookmarkDict[bookmarkName]];
          Microsoft.Office.Interop.Word.Range range = bookmark.Range;

          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.DeleteBookmarks", $"removing mark: {bookmarkDict[bookmarkName]}");

          bookmark.Delete();
          range.Text = "";
        }
        catch
        {
          // Silently ignore if bookmark no longer exists
        }
      }
    }
    public void ReplaceFieldMarks(Dictionary<string, VariableData> varList, DataMaster dm)
    {
      string prefix = H.GetSProperty("VarPrefix");

      foreach (Field field in doc.Fields) //each mark in the word document
      {
        string fieldText = field.Code.Text.Trim();
        if ((field.Type == WdFieldType.wdFieldRef || field.Type == WdFieldType.wdFieldEmpty) && fieldText.StartsWith(prefix)) // when the mark is an insert mark if (a)
        {

          //string variableID = fieldText[prefix.Length..];
          string variableID = fieldText.StartsWith(prefix) ? fieldText.Substring(prefix.Length).Split('\\')[0] : fieldText;

          Microsoft.Office.Interop.Word.Range fieldRange = field.Result; // place to insert found

          if (field.Code.Text.Contains(variableID))
          {
            if (varList[variableID].Type == "table") // If the variable is a table
            {

              XmlDocument xmlDoc = new();
              try
              {
                xmlDoc.LoadXml(varList[variableID].Value);
              }
              catch (XmlException ex)
              {
                H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌Error❌❌ - GenerateOuputWord", $"Invalid XML format for (table) variable {variableID}:found text: {varList[variableID].Value}\n   {ex.Message}");
                return;
              }
              xmlDoc.LoadXml(varList[variableID].Value);

              XmlNode tableNode = xmlDoc.DocumentElement;
              if (tableNode == null)
              {
                H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌  - GenerateOuputWord", $"No valid XML data found for variable {variableID}.");
                return;
              }

              // 📌 Count how many rows and columns the table has
              XmlNodeList rows = tableNode.SelectNodes("r")!;
              int rowCount = rows.Count;
              int colCount = rows[0]!.ChildNodes.Count; // Assumes all rows have the same number of columns

              // 📌 Insert a dynamically sized table
              Table table = doc.Tables.Add(fieldRange, rowCount, colCount);

              for (int i = 0; i < rowCount; i++)
              {
                XmlNodeList cells = rows[i]!.SelectNodes("c")!;
                for (int j = 0; j < colCount; j++)
                {
                  table.Cell(i + 1, j + 1).Range.Text = cells[j]!.InnerText;
                }
              }

              // 📌 Apply "MyStyle" formatting
              table.set_Style(H.GetSProperty("tableStyle"));


              field.Unlink(); // Convierte la referencia en texto estático

              H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "GenerateOuputWord", $"Tabla insertada y referencia '{variableID}' eliminada correctamente.");
            }
            else if (varList[variableID].Type == "list<str>" || varList[variableID].Type == "list<num>")
            {
              //if list<num> isNumber set to true, if list<str> isNumber set to false
              bool isNumber = varList[variableID].Type == "list<num>";

              List<string> listData = dm.GetValueList(variableID, isNumber);

              if (fieldText.Contains('\\'))
              {
                int index = int.Parse(fieldText.Split('\\')[1]);
                if (index < listData.Count)
                {
                  fieldRange.Text = listData[index];
                  field.Unlink(); // Convierte la referencia en texto estático
                }
                else
                {
                  fieldRange.Text = "";
                  field.Unlink(); // Convierte la referencia en texto estático

                  H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.ReplaceFieldMarks", @$"⚠️ Warning ⚠️ : Index: {index} of {fieldText} was out of range. Size of Data list is: {listData.Count}.");
                  continue;
                }
              }
              else
                throw new Exception($"Variable {variableID} is of type {varList[variableID].Type}. It needs the \\x point to the value.");
            }
            else
            {
              fieldRange.Text = varList[variableID].Value;
              field.Unlink(); // Convierte la referencia en texto estático
            }
          }
        }
      }

    }
    public bool SaveAsPDF(string? filePath = null)
    {
      try
      {
        string outputPath = filePath ?? System.IO.Path.ChangeExtension(doc.FullName, ".pdf");

        doc.ExportAsFixedFormat(
            outputPath,
            WdExportFormat.wdExportFormatPDF,
            OpenAfterExport: false,
            OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint,
            Range: WdExportRange.wdExportAllDocument,
            Item: WdExportItem.wdExportDocumentContent,
            IncludeDocProps: true,
            KeepIRM: true,
            CreateBookmarks: WdExportCreateBookmarks.wdExportCreateWordBookmarks,
            DocStructureTags: true,
            BitmapMissingFonts: true,
            UseISO19005_1: false
        );
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.SaveAsPDF", "Archivo docx exportado a pdf con éxito.");
        return true;
      }
      catch
      {
        H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Word.SaveAsPDF", @$"❌Error❌ al Convertir el archivo {doc.Name} a PDF.");
        return false;
      }

    }
    public void Save()
    {
      doc.Save();
    }
    public void Close(bool saveDoc = true)
    {
      if (doc != null)
      {
        doc.Close(SaveChanges: saveDoc);
        _ = Marshal.ReleaseComObject(doc);
        doc = null;
      }
      if (wordApp != null && Marshal.IsComObject(wordApp))
      {
        wordApp.Quit();
        _ = Marshal.ReleaseComObject(wordApp);
        wordApp = null;
      }
    }

    #region Gestionar proceso Word
    /// <summary>
    /// Fuerza el cierre de el proceso de word en el arbol de procesos
    /// </summary>

    public static void CloseWord(bool forceClose = false)
    {
      Process[] wordProcesses = Process.GetProcessesByName("WINWORD");

      if (WordAppDetection(wordProcesses))
      {
        if (forceClose)
        {
          WordAppClose(wordProcesses);
          return;
        }

      CerrarProcesos:
        Console.WriteLine("Existen procesos de Word abiertos. ¿Desea cerrarlos? (s/n): ");
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
          Console.WriteLine("Argumento no válido.");
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
          H.PrintLog(2, "00:00.000", "Main", "WordProcess:", @$"⚠️Cerrando proceso Excel (ID: {proc.Id})...");
          proc.Kill();
          proc.WaitForExit();
        }
        catch (Exception ex)
        {
          H.PrintLog(2, "00:00.000", "Main", "WordProcess:", @$"❌Error❌ al cerrar Excel: {ex.Message}");
        }
      }
    }

    #endregion

  }
}
