using System.Globalization;
using System.Text;
using System.Xml;
using SmartBid;
using Outlook = Microsoft.Office.Interop.Outlook;

public static class H //Helper class for reading properties from an XML file 
{
  private static string PROPERTIES_FILEPATH;
  private static Dictionary<string, object> propertyCache = new Dictionary<string, object>();

  static H()
  {
    string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

    Console.WriteLine($"base directory: {baseDirectory}");

    if (baseDirectory.Equals("C:\\InSync\\SmartBid\\bin\\Debug\\net8.0-windows10.0.22000.0\\", StringComparison.OrdinalIgnoreCase))
    {
      PROPERTIES_FILEPATH = "C:/InSync/SmartBid/properties.xml";
    }
    else
    {
      PROPERTIES_FILEPATH = Path.Combine(baseDirectory, "properties.xml");
    }
    H.PrintLog(2, "Helper", "Helper", $"PROPERTIES_FILEPATH: {PROPERTIES_FILEPATH}");
  }

  public static string GetSProperty(string name)
  {
    if (propertyCache.ContainsKey(name))
    {
      return propertyCache[name].ToString();
    }

    if (!File.Exists(Path.GetFullPath(PROPERTIES_FILEPATH)))
    {
      PrintLog(2, "Helper", "GetSProperty", $" ****** FILE: {PROPERTIES_FILEPATH} NOT FOUND. ******\n Review properties.xml file location or update PROPERTIES_FILEPATH in H.cs Class file\n\n");
      _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", PROPERTIES_FILEPATH);
    }

    XmlDocument xmlDoc = new XmlDocument();
    xmlDoc.Load(PROPERTIES_FILEPATH);
    XmlNode node = xmlDoc.SelectSingleNode($"/root/{name}");
    string value = node != null ? node.InnerText : string.Empty;
    propertyCache[name] = value;
    return value;
  }

  public static int GetNProperty(string name)
  {
    string stringValue = GetSProperty(name);
    if (int.TryParse(stringValue, out int result))
    {
      propertyCache[name] = result;
      return result;
    }
    return 0;
  }

  public static bool GetBProperty(string name)
  {
    string stringValue = GetSProperty(name);
    if (bool.TryParse(stringValue, out bool result))
    {
      propertyCache[name] = result;
      return result;
    }
    return false;
  }

  public static void PrintXML(int level, XmlDocument xmlDoc)
  {

    if (GetNProperty("printLevel") <= level)
    {
      StringWriter sw = new StringWriter();
      XmlTextWriter writer = new XmlTextWriter(sw) { Formatting = Formatting.Indented };
      xmlDoc.WriteTo(writer);
      Console.WriteLine(sw.ToString()); // Print formatted XML
    }
    return; // Only print if log level is sufficient
  }

  public static void PrintLog(int level = 2, string user = "", string eventLog = "info", string message = "")
  {
    if (GetNProperty("printLevel") <= level)
    {
      Console.WriteLine($"{level} - user: {user} >> {eventLog}: {message}");
    }
    if (GetNProperty("logLevel") <= level)
    {
      DBtools.LogMessage(level, user, eventLog, message);
    }
  }

  public static XmlElement CreateElement(XmlDocument doc, string name, string value)
  {
    XmlElement element = doc.CreateElement(name);
    element.InnerText = value;
    return element;
  }

  public static bool MailTo(List<string> email, string subject, string text)
  {
    try
    {
      var outlookApp = new Outlook.Application();

      var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

      mailItem.Subject = $"Prueba C#"; //Asunto del correo

      if (email.Count > 0)
      {
        foreach (string _email in email)
        {
          _ = mailItem.Recipients.Add(_email);
        }

        mailItem.Body = $"Correo enviado desde C#"; //Texto dentro del correo

        //mailItem.Attachments.Add(""); //Ruta de archivo adjunto

        mailItem.Send();

        PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "_EnviarMail:", $"Correo enviado a:\n  {string.Join("\n  ", email)}"

        );
      }
    }
    catch (Exception ex)
    {
      PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "_EnviarMail:", $"❌Error❌ al enviar el correo.");
      return false;
    }
    return true;
  }

  public static string EliminarDiacriticos(string texto)
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
        _ = stringBuilder.Append(c);
      }
    }

    return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
  }

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



}


