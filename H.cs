using System.Globalization;
using System.Text;
using System.Xml;
using SmartBid;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SmartBid
{


  public static class H //Helper class for reading properties from an XML file 
  {
    private static string PROPERTIES_FILEPATH;
    private static Dictionary<string, object> propertyCache = [];

    static H()
    {
      string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

      if (baseDirectory.Equals("C:\\InSync\\SmartBid\\bin\\Debug\\net8.0-windows10.0.22000.0\\", StringComparison.OrdinalIgnoreCase))
      {
        PROPERTIES_FILEPATH = "C:/InSync/SmartBid/properties.xml";
      }
      else
      {
        PROPERTIES_FILEPATH = Path.Combine(baseDirectory, "properties.xml");
      }
      H.PrintLog(2, "Helper", "Helper", $"base directory: {baseDirectory}");
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

      XmlDocument xmlDoc = new();
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
        StringWriter sw = new();
        XmlTextWriter writer = new(sw) { Formatting = Formatting.Indented };
        xmlDoc.WriteTo(writer);
        Console.WriteLine(sw.ToString()); // Print formatted XML
      }
      return; // Only print if log level is sufficient
    }
    public static void SaveXML(int level, XmlDocument xmlDoc, string fileName)
    {
      if (GetNProperty("printLevel") <= level)
      {
        using var writer = new XmlTextWriter(fileName, System.Text.Encoding.UTF8);
        writer.Formatting = Formatting.Indented;
        xmlDoc.WriteTo(writer);
      }
      return; // Only save if log level is sufficient
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

    public static void MergeXmlNodes(XmlDocument sourceDoc, XmlDocument targetDoc, string sourcePath, string targetPath)
    {
      XmlNode sourceParent = sourceDoc.SelectSingleNode(sourcePath);
      XmlNode targetParent = targetDoc.SelectSingleNode(targetPath);

      if (sourceParent == null || targetParent == null)
        return;

      foreach (XmlNode sourceChild in sourceParent.ChildNodes)
      {
        // Buscar nodo con el mismo nombre en el destino
        XmlNode targetChild = targetParent.SelectSingleNode(sourceChild.Name);

        if (targetChild == null)
        {
          // Si no existe, importar y añadir el nodo completo
          XmlNode imported = targetDoc.ImportNode(sourceChild, true);
          targetParent.AppendChild(imported);
        }
        else
        {
          // Si existe, añadir los hijos del nodo fuente al nodo destino
          foreach (XmlNode subNode in sourceChild.ChildNodes)
          {
            XmlNode importedSubNode = targetDoc.ImportNode(subNode, true);
            targetChild.AppendChild(importedSubNode);
          }
        }
      }
    }


    public static bool MailTo(List<string> email, string subject, string body = "", string? attachmentPath = null)
    {
      // Asegura logs sin depender de ThreadContext
      string userForLog = ThreadContext.CurrentThreadInfo?.Value?.User ?? "SYSTEM";

      try
      {
        if (email == null || email.Count == 0)
        {
          PrintLog(5, userForLog, "MailTo", "⚠️ No recipients supplied.");
          PrintLog(5, userForLog, "Subject", $"{subject}");
          return false;
        }

        var outlookApp = new Outlook.Application();
        var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

        // Usar los parámetros recibidos (antes estaban hardcodeados)
        mailItem.Subject = subject ?? string.Empty;
        mailItem.Body = body ?? string.Empty;

        foreach (string _email in email)
        {
          if (!string.IsNullOrWhiteSpace(_email))
            _ = mailItem.Recipients.Add(_email);
        }

        // Adjuntar archivo opcional si existe
        if (!string.IsNullOrWhiteSpace(attachmentPath) && File.Exists(attachmentPath))
        {
          // olByValue es el uso más común para adjuntar
          mailItem.Attachments.Add(attachmentPath,
              Outlook.OlAttachmentType.olByValue,
              Type.Missing, Type.Missing);
        }

        // Intenta resolver contactos locales de Outlook (opcional)
        _ = mailItem.Recipients.ResolveAll();

        mailItem.Send();

        PrintLog(5, userForLog, "MailTo",
            $"Correo enviado a:\n {string.Join("\n ", email)}\n" +
            (string.IsNullOrWhiteSpace(attachmentPath) ? "" : $"Adjunto: {attachmentPath}"));
        return true;
      }
      catch (Exception ex)
      {
        PrintLog(2, userForLog, "MailTo", $"❌ Error al enviar el correo: {ex.Message}");
        return false;
      }
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

}
