using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using Microsoft.VisualBasic.ApplicationServices;
using SmartBid;
using static SmartBid.TC; // Esto ya debería estar ahí
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
      H.PrintLog(2, "00:00.000", "SYSTEM", "Helper", $"base directory: {baseDirectory}");
      H.PrintLog(2, "00:00.000", "SYSTEM", "Helper", $"PROPERTIES_FILEPATH: {PROPERTIES_FILEPATH}");
    }

    public static string GetSProperty(string name)
    {
      if (propertyCache.ContainsKey(name))
      {
        return propertyCache[name].ToString();
      }

      if (!File.Exists(Path.GetFullPath(PROPERTIES_FILEPATH)))
      {
        PrintLog(2, "", "SYSTEM", "GetSProperty", $" ****** FILE: {PROPERTIES_FILEPATH} NOT FOUND. ******\n Review properties.xml file location or update PROPERTIES_FILEPATH in H.cs Class file\n\n");
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



    public static void PrintLog(
        int level,
        string timer,
        string user,
        string eventLog,
        string message,
        XmlDocument? xmlDoc = null) // Made xmlDoc nullable
    {

      // Trim user to text before '@' (recipients local-part). If no '@', keep as-is.
      string displayUser = user ?? string.Empty;
      int atIdx = displayUser.IndexOf('@');
      if (atIdx > 0) displayUser = displayUser.Substring(0, atIdx);

      int? currentCallId = TC.ID?.Value?.CallId;
      string contextPrefix = currentCallId.HasValue ? $"[CALL:{currentCallId.Value:D4}]" : "[MAIN]"; // Formato D4 para 4 dígitos
      string? indentedXml = xmlDoc != null ? xmlDoc.OuterXml : null;

      if (GetNProperty("printLevel") <= level)
      {
        // Seleccionar color según nivel
        switch (level)
        {
          case 0:
          case 1:
            Console.ForegroundColor = ConsoleColor.Gray;
            break;
          case 2:
          case 3:
            Console.ForegroundColor = ConsoleColor.Cyan;
            break;
          case 4:
            Console.ForegroundColor = ConsoleColor.Yellow;
            break;
          default:
            Console.ForegroundColor = ConsoleColor.White;
            break;
        }

        Console.WriteLine($"{contextPrefix} {level} {timer} = {displayUser} >> {eventLog}: {message}");
        Console.ResetColor();

        if (xmlDoc != null)
        {
          // Convertir el XmlDocument a texto plano
          string rawXml = xmlDoc.OuterXml;

          // Leer el XML como texto
          using StringReader stringReader = new(rawXml);
          using XmlReader xmlReader = XmlReader.Create(stringReader);

          using StringWriter sw = new();
          XmlWriterSettings settings = new()
          {
            Indent = true,
            IndentChars = "  ", // indentación por nivel
            NewLineHandling = NewLineHandling.Replace,
            NewLineChars = Environment.NewLine
          };

          using XmlWriter writer = XmlWriter.Create(sw, settings);
          writer.WriteNode(xmlReader, true);
          writer.Flush();

          // Añadir indentación global de 4 espacios
          indentedXml = string.Join(Environment.NewLine,
              sw.ToString()
                .Split(new[] { Environment.NewLine }, StringSplitOptions.None)
                .Select(line => "    " + line));

          Console.ForegroundColor = ConsoleColor.Green;
          Console.WriteLine(indentedXml);
          Console.ResetColor();
        }
      }
      if (GetNProperty("logLevel") <= level)
      {
        DBtools.LogMessage(level, user, eventLog, message, currentCallId, indentedXml);
      }
    }

    public static XmlElement CreateElement(XmlDocument doc, string name, string value, Dictionary<string, string>? attributes = null) // Made attributes nullable
    {
      XmlElement element = doc.CreateElement(name);
      element.InnerText = value;
      if (attributes != null)
      {
        foreach (var attr in attributes)
        {
          element.SetAttribute(attr.Key, attr.Value);
        }
      }
      return element;
    }

    public static void MergeXmlNodes(XmlDocument sourceDoc, XmlDocument targetDoc, string sourcePath, string targetPath)
    {
      XmlNode? sourceParent = sourceDoc.SelectSingleNode(sourcePath); // Made nullable
      XmlNode? targetParent = targetDoc.SelectSingleNode(targetPath); // Made nullable

      if (sourceParent == null || targetParent == null)
        return;

      foreach (XmlNode sourceChild in sourceParent.ChildNodes)
      {
        // Buscar nodo con el mismo nombre en el destino
        XmlNode? targetChild = targetParent.SelectSingleNode(sourceChild.Name); // Made nullable

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


    public static bool MailTo(List<string> recipients, string subject, string body = "", string? attachmentPath = null, bool isHtml = false)
    {
      // Asegura logs sin depender de TC.ID.Value si no está establecido (ej. llamada desde el Main thread)
      string userForLog = TC.ID?.Value?.User ?? "SYSTEM";
      int? currentCallId = TC.ID?.Value?.CallId; // Captura el CallId para el log
      string contextPrefix = currentCallId.HasValue ? $"[CALL:{currentCallId.Value:D4}]" : "[MAIN]";


      try
      {
        if (recipients == null || recipients.Count == 0)
        {
          // MODIFICADO: Usar H.PrintLog para registrar el mensaje, que ya maneja el CallId
          H.PrintLog(5, "00:00.000", userForLog, "MailTo", $"⚠️ No recipients supplied for subject: {subject}.");
          return false;
        }

        var outlookApp = new Outlook.Application();
        var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

        mailItem.Subject = subject ?? string.Empty;

        if (isHtml)
          mailItem.HTMLBody = body ?? string.Empty;
        else
          mailItem.Body = body ?? string.Empty;

        foreach (string _email in recipients)
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

        // MODIFICADO: Usar H.PrintLog para registrar el mensaje
        H.PrintLog(5, "00:00.000", userForLog, "MailTo",
            $"Correo enviado a:\n {string.Join("\n ", recipients)}\n" +
            (string.IsNullOrWhiteSpace(attachmentPath) ? "" : $"Adjunto: {attachmentPath}"));
        return true;
      }
      catch (Exception ex)
      {
        // MODIFICADO: Usar H.PrintLog para registrar el error
        H.PrintLog(2, "00:00.000", userForLog, "MailTo", $"❌ Error al enviar el correo: {ex.Message}");
        return false;
      }
    }

    public static string GetFileMD5(string path)
    {
      using var stream = File.OpenRead(path);
      using var md5 = MD5.Create();
      byte[] hash = md5.ComputeHash(stream);
      return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
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