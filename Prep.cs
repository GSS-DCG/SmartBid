using System.Diagnostics;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.EMMA;
using ExcelDataReader;

namespace SmartBid
{
  internal class PREP // Ya no es static si tuviera el modificador. Si no lo tiene, es de instancia por defecto.
  {
    private XmlDocument _inputXML = new();
    private XmlDocument _outputXML = new();
    private bool _integrated = true;

    // Constructor para inicializar el modo integrado/standalone
    public PREP()
    {
    }

    // El método Run ahora es de instancia (eliminar 'static')
    public XmlDocument Run(XmlDocument input)
    {
      _inputXML = input; // Asigna a la instancia
      _outputXML = CreateOutputXml(out XmlElement utils); // Crea para la instancia
      DoStuff();
      return _outputXML;
    }


    private string ReadFromStdIn() // YA NO ES STATIC
    {
      using StreamReader reader = new(Console.OpenStandardInput(), Encoding.UTF8);
      return reader.ReadToEnd().Trim();
    }

    private string ReadFromFile(string inputFile) // YA NO ES STATIC
    {
      if (!File.Exists(inputFile))
      {
        ShowError($"El archivo de entrada no existe en la ruta especificada: {inputFile}");
        return string.Empty;
      }
      return File.ReadAllText(inputFile);
    }

    private void DoStuff() // YA NO ES STATIC
    {
      var variableNodes = _inputXML.SelectNodes("//call/variables/variable"); // Acceso a campo de instancia
      var areas = new HashSet<string>();

      foreach (XmlNode variable in variableNodes)
      {
        var areaAttr = variable.Attributes?["area"];
        if (areaAttr != null && !string.IsNullOrWhiteSpace(areaAttr.Value))
          areas.Add(areaAttr.Value.Trim());
      }

      List<Task<XmlDocument>> prepCallTasks = new();

      foreach (string area in areas)
      {
        string currentArea = area;
        prepCallTasks.Add(Task.Run(() =>
        {
          // Asegurarse de que CreateAreaCall acceda al _inputXML de esta instancia o se le pase.
          // En este caso, CreateAreaCall también se hará de instancia.
          XmlDocument areaCall = CreateAreaCall(currentArea);
          return MakePrepCall(currentArea, areaCall); // MakePrepCall ya devuelve un XmlDocument local
        }));
        Thread.Sleep(100);
      }

      Task.WhenAll(prepCallTasks).Wait();

      foreach (var task in prepCallTasks)
      {
        XmlDocument areaResult = task.Result;
        MergeResults(areaResult); // MergeResults modificará _outputXML de esta instancia
      }

      PlaySuccessBeep();
    }

    // YA NO ES STATIC. Accede a _inputXML de la instancia.
    private XmlDocument CreateAreaCall(string area)
    {
      XmlDocument areaCall = new();
      XmlDeclaration decl = areaCall.CreateXmlDeclaration("1.0", "UTF-8", null);
      areaCall.AppendChild(decl);

      XmlElement root = areaCall.CreateElement("call");
      areaCall.AppendChild(root);
      XmlElement variables = areaCall.CreateElement("variables");
      root.AppendChild(variables);

      XmlElement variablesOut = areaCall.CreateElement("out");
      // Usa el _inputXML de la instancia
      foreach (XmlNode var in _inputXML.SelectNodes("//variables/*"))
      {
        string varArea = var.Attributes["area"]!.Value;
        if (varArea == area)
          variablesOut.AppendChild(areaCall.ImportNode(var, true));
      }
      variables.AppendChild(variablesOut);

      XmlElement inputDocs = areaCall.CreateElement("inputDocs");
      // Usa el _inputXML de la instancia
      foreach (XmlElement node in _inputXML.SelectNodes($"//inputDocs/{area}")!)
      {
        inputDocs.AppendChild(areaCall.ImportNode(node, deep: true));
      }

      root.AppendChild(inputDocs);
      return areaCall;
    }

    // YA NO ES STATIC. Modifica _outputXML de la instancia.
    private void MergeResults(XmlDocument result)
    {
      XmlNodeList vars = result.SelectNodes("//answer/variables/*");
      XmlNodeList utils = result.SelectNodes("//answer/utils/*");

      XmlNode outputVars = _outputXML.SelectSingleNode("//answer/variables"); // Acceso a campo de instancia
      XmlNode outputUtils = _outputXML.SelectSingleNode("//answer/utils");     // Acceso a campo de instancia

      if (outputVars != null)
      {
        foreach (XmlNode var in vars)
          outputVars.AppendChild(_outputXML.ImportNode(var, true)); // Importa en el _outputXML de la instancia
      }

      if (outputUtils != null)
      {
        foreach (XmlNode util in utils)
          outputUtils.AppendChild(_outputXML.ImportNode(util, true)); // Importa en el _outputXML de la instancia
      }
    }

    // YA NO ES STATIC. No accede a campos de instancia de PREP, pero es llamado por DoStuff de instancia.
    // Puede ser static o de instancia; por consistencia y para que sea callable por DoStuff de instancia,
    // lo convertimos a instancia.
    private XmlDocument MakePrepCall(string area, XmlDocument areaCall)
    {
      // ... (el código existente de MakePrepCall, que es bastante autónomo.
      // Asegúrate de que las llamadas a Print y ShowError usen la versión de instancia)
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
        prepToolPath = @"python.exe";
        arguments = $"\"{Path.Combine(prepFolder, prepPy)}\" 00";
      }
      else
      {
        ShowError($"No se encontró la herramienta para el área '{area}': {prepToolPath}");
        return new XmlDocument(); // Retorna vacío si no se encuentra el ejecutable
      }

      H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "MakePrepCall", $"Ejecutando {Path.GetFileName(prepToolPath)} {arguments}");
      H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "MakePrepCall", $"Call sent to {area}\n", areaCall);

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
        process.Start();

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

        return result;
      }
      catch (Exception ex)
      {
        ShowError($"Error al parsear la respuesta XML de {area}Prep.exe: {ex.Message}");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "MakePrepCall", $"Output recibido:\n{output}");

        return new XmlDocument(); // Retorna vacío si falla el parseo
      }
    }


    private void ShowError(string message) // YA NO ES STATIC
    {
      Console.WriteLine($"❌Error❌: {message}");
      Console.Beep(220, 1000);
      Thread.Sleep(5000);
    }

    private void PlaySuccessBeep() // YA NO ES STATIC
    {
      Console.Beep(880, 200);
    }

    // YA NO ES STATIC. Crea un XmlDocument para la instancia actual.
    private XmlDocument CreateOutputXml(out XmlElement utils)
    {
      XmlDocument outputXML = new();
      XmlDeclaration xmlDeclaration = outputXML.CreateXmlDeclaration("1.0", "UTF-8", null);
      outputXML.AppendChild(xmlDeclaration);

      XmlElement root = outputXML.CreateElement("answer");
      root.SetAttribute("result", "OK");
      outputXML.AppendChild(root);

      XmlElement variables = outputXML.CreateElement("variables");
      root.AppendChild(variables);

      utils = outputXML.CreateElement("utils");
      root.AppendChild(utils);

      return outputXML;
    }

    //private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, string value, string origin) { /* ... */ return null; } // Código original
    //private static XmlElement CreateXmlVariable(XmlDocument outputXML, string ID, string type, XmlElement value, string origin) { /* ... */ return null; } // Código original
    //private static XmlElement CreateElementWithXML(XmlDocument doc, string name, string value, params (string, string)[] attributes) { /* ... */ return null; } // Código original
    //private static XmlElement CreateElementWithText(XmlDocument doc, string name, string text, params (string, string)[] attributes) { /* ... */ return null; } // Código original
    //private void PrintFormattedXml(XmlDocument xmlDoc, string label) { /* ... */ } // YA NO ES STATIC
    //private int GetRandomNumber() { /* ... */ return 0; } // YA NO ES STATIC
    //private XmlElement CreateElement(XmlDocument doc, string name, string value) { /* ... */ return null; } // YA NO ES STATIC
    //public static Dictionary<string, string[]> GetValuesFromFile(string excelPath) { /* ... */ return null; } // Código original (puede permanecer static)

  }
}