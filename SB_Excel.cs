using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Org.BouncyCastle.Asn1.BC;
using SmartBid;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmartBid
{
  public class SB_Excel
  {
    private Microsoft.Office.Interop.Excel.Application excelApp;
    private Microsoft.Office.Interop.Excel.Workbook? workbook;
    private string filePath;

    public SB_Excel(string filePath)
    {
      this.filePath = filePath;

      excelApp = new Microsoft.Office.Interop.Excel.Application { Visible = false };
      workbook = excelApp.Workbooks.Open(filePath);
    }
    
    public bool FillUpValue(string rangeName, string value)
    {
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"Procesando celda '{rangeName}'...");

      Excel.Range? range = null;

      string rangeNameCleaned = Regex.Replace(rangeName, @"(?i)Call1_", "");

        try
        {
          if (!Regex.IsMatch(rangeName, @"(?i)Call1_"))
            range = workbook!.Names.Item(rangeName).RefersToRange;
          else
            range = workbook!.Names.Item(rangeNameCleaned).RefersToRange;
        }
        catch (Exception)
        {
          //if rangeName is type list (it ends with '/n' where n is a number then do not throw and print out a message the the value is not filled up in the spreadsheet
          if (rangeName.Contains(@"\"))
          {
            H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"List Cell: '{rangeName}' does not exist, List is longer than expected. Value has not been set.");
            return false;
          }
          else
            throw;
        }

      bool validationCheck = H.GetBProperty("ExcelValidationCheck");
      if (!validationCheck)
      {
        range.Value = value;
        H.PrintLog(0, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"Validación desactivada. Valor '{value}' escrito directamente en '{rangeName}'.");
        return true;
      }

      var validation = range.Validation;
      Excel.XlDVType validationType;

      try
      {
        validationType = (Excel.XlDVType)validation.Type;
      }
      catch
      {
        range.Value = value;
        H.PrintLog(0, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"La celda '{rangeName}' no tiene validación. Valor '{value}' escrito directamente.");
        return true;
      }

      if (validationType == Excel.XlDVType.xlValidateList)
      {
        string formula = validation.Formula1;

        string[] allowedValues;

        if (formula.StartsWith("="))
        {
          try
          {
            string rangeRef = formula[1..]; // Elimina el '='
            Excel.Range listRange = workbook.Application.Range[rangeRef];
            allowedValues = listRange.Cells.Cast<Excel.Range>()
                                .Select(static cell => cell.Value!.ToString())
                                .Where(static val => !string.IsNullOrEmpty(val))
                                .ToArray();
          }
          catch (Exception ex)
          {
            H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"❌❌ Error ❌❌  al acceder al rango de fórmula: {ex.Message}");
            return false;
          }
        }
        else
        {
          string listSeparator = workbook.Application.International[Excel.XlApplicationInternational.xlListSeparator].ToString()!;

          allowedValues = formula
            .Split([listSeparator], StringSplitOptions.RemoveEmptyEntries)
            .Select(v => v.Trim().Replace("\"", ""))
            .ToArray();
        }

        H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", "Valores permitidos:");
        foreach (var val in allowedValues)
          H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"- {val}");

        if (allowedValues.Contains(value))
        {
          range.Value = value;
          H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"Valor '{value}' escrito correctamente en '{rangeName}'.");
          return true;
        }
        else
        {
          H.PrintLog(3, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"❌❌ Error ❌❌ : El valor '{value}' no está permitido en '{rangeName}'.");
          return false;
        }
      }
      else
      {
        range.Value = value;
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.FillUpValue", $"La celda '{rangeName}' no tiene validación de lista. Valor '{value}' escrito directamente.");
        return true;
      }
    }

    public void WriteTable(string rangeName, XmlNode doc)
    {
      if (Regex.IsMatch(rangeName, @"(?i)Call1_"))
        rangeName = Regex.Replace(rangeName, @"(?i)Call1_", "");

      try
      {
        // 📌 Parse XML Input
        var rows = doc.SelectNodes("//t/r");
        int rowCount = rows!.Count;
        int colCount = rows[0]!.ChildNodes.Count;

        // 📌 Write data into range
        Excel.Name namedRange = workbook!.Names.Item(rangeName);
        Excel.Range inputRange = namedRange.RefersToRange;

        if (inputRange.Rows.Count != rowCount || inputRange.Columns.Count != colCount)
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "WriteTableToExcel", $"Size mismatch at {rangeName} in {Path.GetFileName(filePath)}: Input ({rowCount}x{colCount}) vs Range ({inputRange.Rows.Count}x{inputRange.Columns.Count}).");
          return;
        }

        for (int i = 0; i < rowCount; i++)
          for (int j = 0; j < colCount; j++)
            ((Excel.Range)inputRange.Cells[i + 1, j + 1]).Value = rows[i]!.ChildNodes[j]!.InnerText;

      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.WriteTable", $"❌❌ Error ❌❌ :  writing table to Excel: " + ex.Message);
      }
    }

    public string GetSValue(string rangeName, bool isNumber = false)
    {
      Excel.Range cell = workbook.Names.Item(rangeName).RefersToRange;
      if (cell.Value != null)
      {
        string rawValue = cell.Value.ToString();

        if (isNumber)
        {
          // Normalizar el valor: eliminar separadores de miles y convertir coma decimal a punto
          string normalized = rawValue
            .Replace("'", "")   // eliminar apóstrofes como separador de miles
            .Replace(" ", "")   // eliminar espacios
            .Replace(",", "."); // convertir coma decimal a punto

          if (double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue))
          {
            return parsedValue.ToString(CultureInfo.InvariantCulture); // devuelve con punto decimal
          }
          else
          {
            throw new FormatException($"❌ El valor '{rawValue}' del rango '{rangeName}' no puede convertirse a número.");
          }
        }

        return rawValue;
      }
      else
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.GetSValue", $"❌ Named range '{rangeName}' is empty or not found in the workbook '{filePath}'.");
        return string.Empty;
      }
    }

    public XmlElement GetTValue(string rangeName)
    {
      Excel.Range outputRange = null;
      try
      {
        if (workbook == null)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.GetTValue", $"Workbook object is null. Cannot read table from range '{rangeName}'.");
          throw new InvalidOperationException($"Workbook is not loaded. Cannot read table from range '{rangeName}'.");
        }

        // 📌 Read data from range
        Excel.Name namedRange = workbook.Names.Item(rangeName);
        outputRange = namedRange.RefersToRange;
        int outRows = outputRange.Rows.Count;
        int outCols = outputRange.Columns.Count;

        // Create a NEW XmlDocument instance to host the XML fragment
        XmlDocument tempDoc = new();
        XmlElement root = tempDoc.CreateElement("t");

        for (int i = 1; i <= outRows; i++)
        {
          XmlElement row = tempDoc.CreateElement("r");
          for (int j = 1; j <= outCols; j++)
          {
            XmlElement cell = tempDoc.CreateElement("c");
            // Safely convert cell value to string. Use .Text for displayed text.
            cell.InnerText = Convert.ToString(((Excel.Range)outputRange.Cells[i, j]).Text);
            _ = row.AppendChild(cell);
          }
          _ = root.AppendChild(row);
        }

        return root; // Return the created XmlElement, which belongs to tempDoc

      }
      catch (Exception ex)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.GetTValue", $"❌❌ Error ❌❌  reading table from Excel range '{rangeName}': {ex.Message}");
        // Re-throw or return null depending on desired error propagation
        throw; // Propagate the exception to the caller
      }
      finally
      {
        // Release COM object for the range
        if (outputRange != null)
        {
          _ = Marshal.ReleaseComObject(outputRange);
        }
        // No need to release namedRange here as it's a value type from workbook.Names.Item
      }
    }

    public void Close()
    {
      try
      {
        if (workbook != null)
        {
          workbook.Close(false); // Close without saving changes
          _ = Marshal.ReleaseComObject(workbook);
          workbook = null;
        }
        if (excelApp != null)
        {
          excelApp.Quit();
          _ = Marshal.ReleaseComObject(excelApp);
          excelApp = null;
        }
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Close", $"Workbook '{filePath}' closed successfully.");
      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Close", $"❌ Error closing workbook '{filePath}': {ex.Message}");
      }
    }

    public void Save()
    {
      try
      {
        if (workbook == null)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Save", $"Workbook object is null. Cannot save '{filePath}'. Ensure the workbook is opened in the constructor or an 'Open' method.");
          throw new InvalidOperationException($"Workbook is not loaded for file: {filePath}");
        }
        workbook.Save();
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Save", $"Workbook '{filePath}' saved successfully.");
      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Save", $"❌ Error saving workbook '{filePath}': {ex.Message}");
        throw;
      }
    }

    public void Calculate()
    {
      try
      {
        if (workbook == null)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Calculate", $"Workbook object is null. Cannot calculate or save '{filePath}'. Ensure the workbook is opened in the constructor or an 'Open' method.");
          throw new InvalidOperationException($"Workbook is not loaded for file: {filePath}");
        }

        excelApp.Calculate();

        workbook.Save();

        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Calculate", $"Workbook '{filePath}' calculated and saved successfully.");
      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "SB_Excel.Calculate", $"❌ Error calculating or saving workbook '{filePath}': {ex.Message}");
        throw;
      }
    }

    //helper to show all ranges in the workbook (no needed in the app)
    public List<string> ListNamedRanges()
    {
      var namedRanges = new List<string>();
      try
      {
        if (workbook == null)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ListNamedRanges", "❌ Workbook is null.");
          return namedRanges;
        }

        foreach (Excel.Name name in workbook.Names)
        {
          namedRanges.Add(name.Name);
          _ = Marshal.ReleaseComObject(name);
        }
      }
      catch (Exception ex)
      {
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ListNamedRanges", $"❌ Error listing named ranges: {ex.Message}");
      }

      return namedRanges;
    }

    #region Gestionar proceso Excel
    /// <summary>
    /// Fuerza el cierre de el proceso de excel en el arbol de procesos
    /// </summary>
    public static void CloseExcel(bool forceClose = false)
    {
      Process[] excelProcesses = Process.GetProcessesByName("EXCEL");

      if (ExcelAppDetection(excelProcesses))
      {
        if (forceClose)
        {
          ExcelAppClose(excelProcesses);
          return;
        }

      CerrarProcesos:
        Console.WriteLine("Existen procesos de Excel abiertos. ¿Desea cerrarlos? (s/n): ");
        string var = Console.ReadLine()!;

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
          Console.WriteLine($"SB_Excel.CloseExcel  ❌❌ Error ❌❌ : Argumento no válido.");
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
          Console.WriteLine(@$"⚠️Cerrando proceso excel (ID: {proc.Id})...");
          proc.Kill();
          proc.WaitForExit();
        }
        catch (Exception ex)
        {
          Console.WriteLine( @$"❌Error❌ al cerrar excel: {ex.Message}");
        }
      }
    }
    #endregion

  }
}



