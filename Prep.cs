using System.Diagnostics;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.EMMA;
using ExcelDataReader;

namespace SmartBid
{

  internal class PREP
  {
    private static XmlDocument inputXML = new();
    private static XmlDocument outputXML = new();
    private static bool integrated = true;

    static void _Main(string[] args)
    {
      integrated = args.Length == 0;
      string input = integrated ? ReadFromStdIn() : ReadFromFile(@"C:\InSync\PREP\prepCall.xml");

      if (string.IsNullOrWhiteSpace(input))
      {
        ShowError("No se proporcionó ningún argumento válido.");
        return;
      }

      try
      {
        inputXML.LoadXml(input);
      }
      catch (Exception ex)
      {
        ShowError($"Error al cargar XML: {ex.Message}");
        return;
      }

      outputXML = CreateOutputXml(out XmlElement utils);
      DoStuff();

      if (integrated)
      {
        Console.WriteLine(outputXML.OuterXml); // Output plano para integración
      }
      else
      {
        PrintFormattedXml(inputXML, "INPUT:");
        PrintFormattedXml(outputXML, "OUTPUT:");
      }
    }
    public static XmlDocument Run(XmlDocument input)
    {
      inputXML = input;

      outputXML = CreateOutputXml(out XmlElement utils);
      DoStuff();

      return outputXML;

    }

    private static void Print(string message)
    {
      if (!integrated)
        Console.WriteLine(message);
    }

    private static string ReadFromStdIn()
    {
      using StreamReader reader = new(Console.OpenStandardInput(), Encoding.UTF8);
      return reader.ReadToEnd().Trim();
    }

    private static string ReadFromFile(string inputFile)
    {
      Print("Running standalone\n");

      if (!File.Exists(inputFile))
      {
        ShowError($"El archivo de entrada no existe en la ruta especificada: {inputFile}");
        return string.Empty;
      }

      return File.ReadAllText(inputFile);
    }

    private static void DoStuff()
    {
      // 1. Detectar todas las áreas únicas en inputXML
      var variableNodes = inputXML.SelectNodes("//call/variables/variable");
      var areas = new HashSet<string>();

      foreach (XmlNode variable in variableNodes)
      {
        var areaAttr = variable.Attributes?["area"];
        if (areaAttr != null && !string.IsNullOrWhiteSpace(areaAttr.Value))
          _ = areas.Add(areaAttr.Value.Trim());
      }

      // List to hold all tasks, each returning an XmlDocument
      List<Task<XmlDocument>> prepCallTasks = new();

      // 2. Procesar cada área en paralelo
      foreach (string area in areas)
      {
        // Capture the current 'area' for the lambda expression
        string currentArea = area;
        prepCallTasks.Add(Task.Run(() =>
        {
          XmlDocument areaCall = CreateAreaCall(currentArea);
          return MakePrepCall(currentArea, areaCall);
        }));
      }

      // Wait for all tasks to complete
      Task.WhenAll(prepCallTasks).Wait();

      // 3. Fusionar todos los resultados después de que todas las llamadas hayan terminado
      foreach (var task in prepCallTasks)
      {
        // Retrieve the result from each completed task
        XmlDocument areaResult = task.Result;
        MergeResults(areaResult);
      }

      PlaySuccessBeep();
    }

    private static XmlDocument CreateAreaCall(string area)
    {
      XmlDocument areaCall = new();
      XmlDeclaration decl = areaCall.CreateXmlDeclaration("1.0", "UTF-8", null);
      _ = areaCall.AppendChild(decl);

      XmlElement root = areaCall.CreateElement("call");
      _ = areaCall.AppendChild(root);

      XmlElement variables = areaCall.CreateElement("out");
      foreach (XmlNode var in inputXML.SelectNodes("//variables/*"))
      {
        string varArea = var.Attributes["area"]!.Value;
        if (varArea == area)
          _ = variables.AppendChild(areaCall.ImportNode(var, true));
      }
      _ = root.AppendChild(variables);

      XmlElement inputDocs = areaCall.CreateElement("inputDocs");
      foreach (XmlElement node in inputXML.SelectNodes("//inputDocs/*"))
      {

        if (node is XmlElement doc)
        {
          string docType = doc.GetAttribute("type");
          if (docType == area)
            _ = inputDocs.AppendChild(areaCall.ImportNode(doc, true));
        }
      }
      _ = root.AppendChild(inputDocs);

      return areaCall;
    }

    private static void MergeResults(XmlDocument result)
    {
      XmlNodeList vars = result.SelectNodes("//answer/variables/*");
      XmlNodeList utils = result.SelectNodes("//answer/utils/*");

      XmlNode outputVars = outputXML.SelectSingleNode("//answer/variables");
      XmlNode outputUtils = outputXML.SelectSingleNode("//answer/utils");

      if (outputVars != null)
      {
        foreach (XmlNode var in vars)
          _ = outputVars.AppendChild(outputXML.ImportNode(var, true));
      }

      if (outputUtils != null)
      {
        foreach (XmlNode util in utils)
          _ = outputUtils.AppendChild(outputXML.ImportNode(util, true));
      }
    }

    private static XmlDocument MakePrepCall(string area, XmlDocument areaCall)
    {
      string prepFolder = H.GetSProperty("prepFolder");
      string prepExe = $"Prep_{area}.exe";
      string prepPy = $"Prep_{area}.py";
      string prepToolPath = "";
      string arguments = "";

      var stopwatch = Stopwatch.StartNew();

      if (File.Exists(Path.Combine(prepFolder, prepExe)))
      {
        prepToolPath = Path.Combine(prepFolder, prepExe);
      }
      else if (File.Exists(Path.Combine(prepFolder, prepPy)))
      {
        prepToolPath = @"C:\Users\martin.molina\AppData\Local\Programs\Python\Python313\python.exe";
        arguments = $"\"{Path.Combine(prepFolder, prepPy)}\" 00";
      }
      else
      {
        ShowError($"No se encontró la herramienta para el área '{area}': {prepToolPath}");
        return new XmlDocument(); // Retorna vacío si no se encuentra el ejecutable
      }

      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "MakePrepCall", $"Ejecutando {Path.GetFileName(prepToolPath)} {arguments}");
      H.PrintLog(1, ThreadContext.CurrentThreadInfo.Value!.User, "MakePrepCall", $"Call sent to {area}\n");
      H.PrintXML(1, areaCall);

      ProcessStartInfo psi = new()
      {
        FileName = prepToolPath,
        Arguments = arguments,
        RedirectStandardInput = true,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true,
        StandardInputEncoding = Encoding.UTF8
      };

      string output;
      string error;

      using (Process process = new() { StartInfo = psi })
      {
        _ = process.Start();

        using (StreamWriter writer = process.StandardInput)
        {
          writer.Write(areaCall.OuterXml); // No usar WriteLine
        }

        output = process.StandardOutput.ReadToEnd();
        error = process.StandardError.ReadToEnd();

        if (!string.IsNullOrEmpty(error))
        {
          Console.WriteLine("STDERR:");
          Console.WriteLine(error);
        }

        process.WaitForExit();
      }

      if (!string.IsNullOrWhiteSpace(error))
      {
        Console.WriteLine($"⚠️ Error desde {area}Prep.exe:\n{error}");
      }

      try
      {
        XmlDocument result = new();
        result.LoadXml(output);

        stopwatch.Stop();
        Print($"Finished processing area: {area} in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
        return result;
      }
      catch (Exception ex)
      {
        ShowError($"Error al parsear la respuesta XML de {area}Prep.exe: {ex.Message}");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "MakePrepCall", $"Output recibido:\n{output}");

        return new XmlDocument(); // Retorna vacío si falla el parseo
      }
    }

    private static void ShowError(string message)
    {
      Console.WriteLine($"❌Error❌: {message}");
      Console.Beep(220, 1000);
      Thread.Sleep(5000);
    }

    private static void PlaySuccessBeep()
    {
      Console.Beep(880, 200);
    }

    private static XmlDocument CreateOutputXml(out XmlElement utils)
    {
      XmlDocument outputXML = new();
      XmlDeclaration xmlDeclaration = outputXML.CreateXmlDeclaration("1.0", "UTF-8", null);
      _ = outputXML.AppendChild(xmlDeclaration);

      XmlElement root = outputXML.CreateElement("answer");
      root.SetAttribute("result", "OK");
      _ = outputXML.AppendChild(root);

      XmlElement variables = outputXML.CreateElement("variables");
      _ = root.AppendChild(variables);

      utils = outputXML.CreateElement("utils");
      _ = root.AppendChild(utils);

      return outputXML;
    }

    private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, string value, string origin)
    {
      XmlElement xmlVar = outputXML.CreateElement(ID);
      _ = (outputXML.SelectSingleNode("//answer/variables")?.AppendChild(xmlVar));

      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "value", value, ("type", type)));
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "origin", $"{origin}-{DateTime.Now:yyMMdd-HHmm}"));
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "note", "Calculated Value"));

      return xmlVar;
    }

    private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, XmlElement value, string origin)
    {
      XmlElement xmlVar = outputXML.CreateElement(ID);
      _ = (outputXML.SelectSingleNode("//answer/variables")?.AppendChild(xmlVar));

      _ = xmlVar.AppendChild(CreateElementWithXML(outputXML, "value", value.OuterXml, ("type", type)));
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "origin", $"{origin}-{DateTime.Now:yyMMdd-HHmm}"));
      _ = xmlVar.AppendChild(CreateElementWithText(outputXML, "note", "Calculated Value"));

      return xmlVar;
    }

    private static XmlElement CreateElementWithXML(XmlDocument doc, string name, string value, params (string, string)[] attributes)
    {
      XmlElement element = doc.CreateElement(name);
      foreach (var (key, val) in attributes)
        element.SetAttribute(key, val);
      element.InnerXml = value;
      return element;
    }

    private static XmlElement CreateElementWithText(XmlDocument doc, string name, string text, params (string, string)[] attributes)
    {
      XmlElement element = doc.CreateElement(name);
      foreach (var (key, val) in attributes)
        element.SetAttribute(key, val);
      element.InnerText = text;
      return element;
    }

    private static void PrintFormattedXml(XmlDocument xmlDoc, string label)
    {
      if (integrated) return;  //Don't print anything if integrated

      Console.WriteLine(label);
      using StringWriter sw = new();
      using XmlTextWriter writer = new(sw) { Formatting = Formatting.Indented };
      xmlDoc.WriteTo(writer);
      Console.WriteLine(sw.ToString());
      Console.WriteLine();
    }

    private static int GetRandomNumber()
    {
      Random random = new();
      int randomNumber = random.Next(0, 1001);
      return randomNumber;
    }

    private static XmlElement CreateElement(XmlDocument doc, string name, string value)
    {
      XmlElement element = doc.CreateElement(name);
      element.InnerText = value;
      return element;
    }

    public static Dictionary<string, string[]> GetValuesFromFile(string excelPath)
    {
      // Leer Excel: columna A = variable ID, columna B = valor
      Dictionary<string, string[]> excelValues = [];

      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
      using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
      using (var reader = ExcelReaderFactory.CreateReader(stream))
      {
        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
          ConfigureDataTable = _ => new ExcelDataTableConfiguration() { UseHeaderRow = false }
        });

        var table = result.Tables[0];
        for (int i = 0; i < table.Rows.Count; i++)
        {
          string key = table.Rows[i][0]?.ToString()?.Trim();
          string[] value = [table.Rows[i][1]?.ToString()?.Trim(), Path.GetFileName(excelPath)];
          if (!string.IsNullOrEmpty(key))
            excelValues[key] = value ?? ["", Path.GetFileName(excelPath)];
        }
      }


      return excelValues;
    }
  }
}