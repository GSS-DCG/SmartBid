using System.Xml;
using Windows.System;

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
        H.PrintLog(2, "Helper", "myEvent", $"PROPERTIES_FILEPATH: {PROPERTIES_FILEPATH}");
    }

    public static string GetSProperty(string name)
    {
        if (propertyCache.ContainsKey(name))
        {
            return propertyCache[name].ToString() as string;
        }

        if (!File.Exists(Path.GetFullPath(PROPERTIES_FILEPATH)))
        {
            PrintLog(2, "Helper", "myEvent", $" ****** FILE: {PROPERTIES_FILEPATH} NOT FOUND. ******\n Review properties.xml file location or update PROPERTIES_FILEPATH in H.cs Class file\n\n");
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

    public static void PrintXML(XmlDocument xmlDoc)
    {

        if (GetNProperty("printLevel") <= 2)
        {
            StringWriter sw = new StringWriter();
            XmlTextWriter writer = new XmlTextWriter(sw) { Formatting = Formatting.Indented };
            xmlDoc.WriteTo(writer);
            Console.WriteLine(sw.ToString()); // Print formatted XML
        }
        return; // Only print if log level is sufficient
    }

    public static void PrintLog (int level = 2, string user = "", string eventLog = "info", string message = "")
    {
        if (GetNProperty("printLevel") <= level)
        {
            Console.WriteLine($"{level} - user: {user}  >>{eventLog}: {message}");
        }
        if (GetNProperty("logLevel") <= level)
        {
            DBtools.LogMessage( level, user, eventLog, message);
        }

    }

    public static XmlElement CreateElement(XmlDocument doc, string name, string value)
    {
        XmlElement element = doc.CreateElement(name);
        element.InnerText = value;
        return element;
    }
}
