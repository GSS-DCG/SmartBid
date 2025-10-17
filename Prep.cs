using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml;

namespace SmartBid
{
  internal class PREP
  {
    private XmlDocument _inputXML = new();
    private XmlDocument _outputXML = new();
    private bool _integrated = true;

    // Constructor para inicializar el modo integrado/standalone
    public PREP()
    {
    }

    public XmlDocument Run(XmlDocument input)
    {
      _inputXML = input; // Asigna a la instancia
      _outputXML = CreateOutputXml(out XmlElement utils); // Crea para la instancia
      DoStuff();
      AddDefaultValues();
      return _outputXML;
    }

    private void DoStuff()
    {
      var variableNodes = _inputXML.SelectNodes("//call/variables/variable"); // Acceso a campo de instancia
      var areas = new HashSet<string>();


      //find every area in prep variables
      foreach (XmlNode variable in variableNodes)
      {
        var areaAttr = variable.Attributes?["area"];
        if (areaAttr != null && !string.IsNullOrWhiteSpace(areaAttr.Value))
          _ = areas.Add(areaAttr.Value.Trim());
      }

      List<Task<XmlDocument>> prepCallTasks = new();

      // Make calls for each area in parallel
      foreach (string area in areas)
      {
        string currentArea = area;
        prepCallTasks.Add(Task.Run(() =>
        {
          XmlDocument areaCall = CreateAreaCall(currentArea);
          return MakePrepCall(currentArea, areaCall); // MakePrepCall ya devuelve un XmlDocument local
        }));
        Thread.Sleep(100);
      }

      Task.WhenAll(prepCallTasks).Wait();

      // Merge results from all areas
      foreach (var task in prepCallTasks)
      {
        XmlDocument areaResult = task.Result;
        MergeResults(areaResult);
      }

      PlaySuccessBeep();
    }

    private XmlDocument CreateAreaCall(string area)
    {
      XmlDocument areaCall = new();
      XmlDeclaration decl = areaCall.CreateXmlDeclaration("1.0", "UTF-8", null);
      _ = areaCall.AppendChild(decl);

      XmlElement root = areaCall.CreateElement("call");
      _ = areaCall.AppendChild(root);
      XmlElement variables = areaCall.CreateElement("variables");
      _ = root.AppendChild(variables);

      XmlElement variablesOut = areaCall.CreateElement("out");
      // Usa el _inputXML de la instancia
      foreach (XmlNode var in _inputXML.SelectNodes("//variables/*"))
      {
        string varArea = var.Attributes["area"]!.Value;
        if (varArea == area)
          _ = variablesOut.AppendChild(areaCall.ImportNode(var, true));
      }
      _ = variables.AppendChild(variablesOut);

      XmlElement inputDocs = areaCall.CreateElement("inputDocs");
      // Usa el _inputXML de la instancia
      foreach (XmlElement node in _inputXML.SelectNodes($"//inputDocs/{area}")!)
      {
        _ = inputDocs.AppendChild(areaCall.ImportNode(node, deep: true));
      }

      _ = root.AppendChild(inputDocs);
      return areaCall;
    }

    private void MergeResults(XmlDocument result)
    {
      XmlNodeList vars = result.SelectNodes("//answer/variables/*");
      XmlNodeList utils = result.SelectNodes("//answer/utils/*");

      XmlNode outputVars = _outputXML.SelectSingleNode("//answer/variables");
      XmlNode outputUtils = _outputXML.SelectSingleNode("//answer/utils");

      if (outputVars != null)
      {
        foreach (XmlNode var in vars)
        {
          _ = outputVars.AppendChild(_outputXML.ImportNode(var, true));
          //Creating a new xmldocument from the var xmlnode to printout
          XmlDocument doc = new();
          doc.AppendChild(doc.ImportNode(var, true));
          H.PrintLog(1, TC.ID.Value!.Time(), TC.ID.Value!.User, "Prep.MergeResults", $"Merged variable: {var.Name}", doc);
        }
        
      }

      if (outputUtils != null)
      {
        foreach (XmlNode util in utils)
          _ = outputUtils.AppendChild(_outputXML.ImportNode(util, true));
      }
    }


    private void AddDefaultValues()
    {
      // Log
      H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "Prep.AddDefaultValues",
                 "Executing Add Default Values to Prep Variables (ID-based, element-per-variable schema)");

      // 1) Input variables: //call/variables/variable
      var inputVars = _inputXML.SelectNodes("//call/variables/variable");
      if (inputVars == null || inputVars.Count == 0) return;

      // 2) Ensure output container: //answer/variables
      var outputVarsNode = EnsureAnswerVariablesContainer(_outputXML);
      if (outputVarsNode == null) return;

      foreach (XmlElement inVar in inputVars)
      {
        // --- Read from input (ID is the ONLY matching key) ---
        string id = inVar.GetAttribute("id");          // REQUIRED for matching
        if (string.IsNullOrWhiteSpace(id)) continue;     // Skip variables without ID

        string type = inVar.GetAttribute("type");        // e.g., num, code, route, text, ...
                                                         // varName is informational only; we ignore it for matching/output structure
                                                         // string varName = inVar.GetAttribute("varName");
                                                         // string area    = inVar.GetAttribute("area");   // not used in output by your schema

        // Default value (inner text of <default>)
        var defaultNode = inVar.SelectSingleNode("default") as XmlElement;
        string defaultVal = defaultNode?.InnerText?.Trim() ?? string.Empty;

        // Unit: prefer @unit on <variable> then <default @unit>
        string unit = FirstNonEmpty(
                          inVar.GetAttribute("unit"),
                          defaultNode?.GetAttribute("unit")
                      ) ?? string.Empty;

        // Allowable spec: from <allowableRange>/<value>
        string allowSpec = GetAllowSpecFromAllowableRange(inVar);

        // 3) Find/create output variable by element name == ID
        XmlElement? outVar = FindOutputVariableById(_outputXML, id);

        if (outVar == null)
        {
          // Create <{id}>
          outVar = _outputXML.CreateElement(id);
          _ = outputVarsNode.AppendChild(outVar);

          // <value type="..." [unit="..."]>default</value>
          var valueEl = EnsureChildElement(outVar, "value");
          if (!string.IsNullOrEmpty(type)) valueEl.SetAttribute("type", type);
          if (!string.IsNullOrEmpty(unit)) valueEl.SetAttribute("unit", unit);

          if (!string.IsNullOrEmpty(defaultVal))
          {
            valueEl.InnerText = defaultVal;
            SetChildText(outVar, "origin", "PREP_DEFAULT");
            SetChildText(outVar, "note", "Default value applied");
          }
          else
          {
            // Keep empty value node; origin/note not added if we didn't apply anything
            if (string.IsNullOrEmpty(valueEl.InnerText))
              valueEl.InnerText = string.Empty;
          }
          continue;
        }

        // Ensure <value>
        var valEl = EnsureChildElement(outVar, "value");

        // Ensure type/unit on <value> (don’t overwrite if exists)
        if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(valEl.GetAttribute("type")))
          valEl.SetAttribute("type", type);

        if (!string.IsNullOrEmpty(unit) && string.IsNullOrEmpty(valEl.GetAttribute("unit")))
          valEl.SetAttribute("unit", unit);

        // Fill default if empty; otherwise validate against allowSpec
        string currentValue = valEl.InnerText?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(currentValue))
        {
          if (!string.IsNullOrEmpty(defaultVal))
          {
            valEl.InnerText = defaultVal;
            SetChildText(outVar, "origin", "PREP_DEFAULT");
            SetChildText(outVar, "note", "Default value applied");
          }
        }
        else if (!string.IsNullOrEmpty(allowSpec))
        {
          if (!ValueSatisfiesAllowSpec(currentValue, allowSpec, type))
          {
            if (!string.IsNullOrEmpty(defaultVal))
            {
              valEl.InnerText = defaultVal;
              SetChildText(outVar, "origin", "PREP_DEFAULT");
              SetChildText(outVar, "note", "Corrected to default due to invalid value");
            }
            else
            {
              SetChildText(outVar, "origin", "PREP_DEFAULT");
              SetChildText(outVar, "note", "Value out of allowableRange and no default available");
            }
          }
        }
      }
    }


    // Helpers: output container and node creation

    /// <summary>
    /// Ensures //answer/variables exists and returns the variables element.
    /// </summary>
    private static XmlElement? EnsureAnswerVariablesContainer(XmlDocument doc)
    {
      // Prefer root <answer>, else any <answer> in doc
      var answer = doc.SelectSingleNode("/answer") as XmlElement
                   ?? doc.SelectSingleNode("//answer") as XmlElement;

      if (answer == null)
      {
        // If the document is empty, create <answer> root
        if (doc.DocumentElement == null)
        {
          answer = doc.CreateElement("answer");
          _ = doc.AppendChild(answer);
        }
        else
        {
          // As a last resort, replace document with a clean <answer>
          doc.RemoveAll();
          answer = doc.CreateElement("answer");
          _ = doc.AppendChild(answer);
        }
      }

      var variables = answer.SelectSingleNode("variables") as XmlElement;
      if (variables == null)
      {
        variables = doc.CreateElement("variables");
        _ = answer.AppendChild(variables);
      }
      return variables;
    }

    /// <summary>
    /// Finds the output variable by element name == id, using a direct path first,
    /// then a robust local-name() fallback for edge cases.
    /// </summary>
    private static XmlElement? FindOutputVariableById(XmlDocument outDoc, string id)
    {
      if (string.IsNullOrWhiteSpace(id)) return null;

      // Fast path (id as a straight element name)
      var direct = outDoc.SelectSingleNode($"//answer/variables/{id}") as XmlElement;
      if (direct != null) return direct;

      // Robust fallback: match by local-name() == id (handles odd characters / namespaces)
      string lit = ToXPathLiteral(id);
      var ln = outDoc.SelectSingleNode($"//answer/variables/*[local-name()={lit}]") as XmlElement;
      return ln;
    }

    /// <summary>
    /// Builds a safe XPath string literal, even if it contains both ' and ".
    /// </summary>
    private static string ToXPathLiteral(string value)
    {
      if (!value.Contains("'")) return $"'{value}'";
      if (!value.Contains("\"")) return $"\"{value}\"";

      // Contains both quotes => concat('a', "\"", 'b', ...)
      var parts = value.Split('\'');
      var sb = new StringBuilder("concat(");
      for (int i = 0; i < parts.Length; i++)
      {
        if (i > 0) _ = sb.Append(",\"'\",");
        _ = sb.Append('\'').Append(parts[i]).Append('\'');
      }
      _ = sb.Append(')');
      return sb.ToString();
    }

    private static XmlElement EnsureChildElement(XmlElement parent, string childName)
    {
      var child = parent.SelectSingleNode(childName) as XmlElement;
      if (child == null)
      {
        child = parent.OwnerDocument!.CreateElement(childName);
        _ = parent.AppendChild(child);
      }
      return child;
    }

    private static void SetChildText(XmlElement parent, string childName, string text)
    {
      var child = EnsureChildElement(parent, childName);
      child.InnerText = text ?? string.Empty;
    }


    // Helpers: input reading and validation

    private static string? FirstNonEmpty(params string?[] candidates)
    {
      return candidates?.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));
    }

    /// <summary>
    /// Collects all &lt;allowableRange&gt;/&lt;value&gt; texts and joins them with ';'
    /// (e.g., "-90:+90; 180").
    /// </summary>
    private static string GetAllowSpecFromAllowableRange(XmlElement inVar)
    {
      var nodes = inVar.SelectNodes("allowableRange/value");
      if (nodes == null || nodes.Count == 0) return string.Empty;

      var parts = nodes
          .OfType<XmlElement>()
          .Select(x => (x.InnerText ?? string.Empty).Trim())
          .Where(x => x.Length > 0)
          .ToArray();

      return string.Join(";", parts);
    }

    /// <summary>
    /// Validates currentValue against allowSpec:
    /// - ';' separates items
    /// - each item can be a literal or a range "min:max" (inclusive)
    /// - numeric parsing tolerant to '+' and ',' decimal; InvariantCulture
    /// - for numeric types (num/int/double/float/decimal), compares numerically; otherwise string match (case-insensitive)
    /// </summary>
    private static bool ValueSatisfiesAllowSpec(string currentValue, string allowSpec, string type)
    {
      if (string.IsNullOrWhiteSpace(currentValue) || string.IsNullOrWhiteSpace(allowSpec))
        return true;
      
      //For now no validation for lists
      if (type != null && (type.Equals("list<num>") || type.Equals("list<str>")))
        return true; 

      static string NormNum(string s) => s.Trim().Replace(',', '.'); // allow comma decimals

      bool isNumericType =
          string.Equals(type, "num", StringComparison.OrdinalIgnoreCase) ||
          string.Equals(type, "int", StringComparison.OrdinalIgnoreCase) ||
          string.Equals(type, "double", StringComparison.OrdinalIgnoreCase) ||
          string.Equals(type, "float", StringComparison.OrdinalIgnoreCase) ||
          string.Equals(type, "decimal", StringComparison.OrdinalIgnoreCase);

      var items = allowSpec
          .Split(';')
          .Select(x => x.Trim())
          .Where(x => x.Length > 0)
          .ToArray();

      if (items.Length == 0) items = new[] { allowSpec.Trim() };

      bool TryParseDouble(string s, out double d) =>
          double.TryParse(NormNum(s), NumberStyles.Any, CultureInfo.InvariantCulture, out d);

      bool curIsNum = TryParseDouble(currentValue, out double curNum);

      foreach (var item in items)
      {
        if (item.Contains(":"))
        {
          // Range "min:max" (inclusive)
          var parts = item.Split(':');
          if (parts.Length != 2) continue;

          if (TryParseDouble(parts[0], out var min) &&
              TryParseDouble(parts[1], out var max))
          {
            if (curIsNum && curNum >= min && curNum <= max)
              return true;
          }
        }
        else
        {
          // Literal
          if (isNumericType || (curIsNum && TryParseDouble(item, out _)))
          {
            if (TryParseDouble(item, out var litNum) && curIsNum)
            {
              if (Math.Abs(curNum - litNum) < 1e-9)
                return true;
            }
          }
          else
          {
            if (string.Equals(currentValue.Trim(), item.Trim(), StringComparison.OrdinalIgnoreCase))
              return true;
          }
        }
      }

      return false;
    }

    
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

    private XmlDocument CreateOutputXml(out XmlElement utils)
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


  }
}