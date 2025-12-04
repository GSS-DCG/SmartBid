using System.Data;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using ExcelDataReader;
using DataTable = System.Data.DataTable;
using File = System.IO.File;

namespace SmartBid
{

  public class VariableData
  {
    // AÑADE ESTA PROPIEDAD DENTRO DE TU CLASE VariableData
    public string ID { get; set; }
    public string VarName { get; set; }
    public string Area { get; set; }
    public string Source { get; set; }
    public bool Critic { get; set; }
    public bool Mandatory { get; set; }
    public string Type { get; set; }
    public string Unit { get; set; }
    public string Default { get; set; }
    public string Description { get; set; }
    public string Prompt { get; set; }
    public string Note { get; set; }
    public string Origen { get; set; }
    public string? InOut { get; set; }
    //public int? Call { get; set; }
    public int Deep { get; set; }
    public List<string> AllowableRange { get; set; }
    public string Value { get; set; }

    // Constructor completo de VariableData (ajusta a tu implementación real)
    public VariableData(
      string id, string varName, string area, string source, bool critic, bool mandatory, string type,
      string unit = "", string defaultValue = "", string description = "", string prompt = "",
      int deep = 0, List<string>? allowableRange = null, string value = ""
    )
    {
      ID = id;
      VarName = varName;
      Area = area;
      Source = source;
      Critic = critic;
      Mandatory = mandatory;
      Type = type;
      Unit = unit;
      Default = defaultValue;
      Description = description;
      Prompt = prompt;
      Deep = deep;
      AllowableRange = allowableRange ?? new List<string>();
      Value = value;
    }

    public VariableData Clone()
    {
      return new VariableData(
        this.ID, this.VarName, this.Area, this.Source, this.Critic, this.Mandatory,
        this.Type, this.Unit, this.Default, this.Description, this.Prompt,
        this.Deep, new List<string>(this.AllowableRange), this.Value
      );
    }
    public XmlDocument ToXMLDocument()
    {
      XmlDocument doc = new();
      _ = doc.AppendChild(ToXML(doc));
      return doc;
    }

    public XmlElement ToXML(XmlDocument mainDoc)
    {
      XmlElement varElem = mainDoc.CreateElement("variable");
      varElem.SetAttribute("source", Source);
      if (!string.IsNullOrEmpty(Area))
        varElem.SetAttribute("area", Area);
      varElem.SetAttribute("id", ID);
      if (!string.IsNullOrEmpty(Type))
        varElem.SetAttribute("type", Type);
      if (!string.IsNullOrEmpty(Unit))
        varElem.SetAttribute("unit", Unit);
      varElem.SetAttribute("varName", VarName);
      varElem.SetAttribute("critic", Critic.ToString().ToLowerInvariant());
      varElem.SetAttribute("mandatory", Mandatory.ToString().ToLowerInvariant());
      varElem.SetAttribute("deep", Deep.ToString());

      // description
      var descriptionElem = mainDoc.CreateElement("description");
      descriptionElem.InnerText = Description;
      _ = varElem.AppendChild(descriptionElem);

      // allowableRange
      if (AllowableRange.Count > 0)
      {
        var rangeElem = mainDoc.CreateElement("allowableRange");
        foreach (var val in AllowableRange)
        {
          var valElem = mainDoc.CreateElement("value");
          valElem.InnerText = val;
          _ = rangeElem.AppendChild(valElem);
        }
        _ = varElem.AppendChild(rangeElem);
      }

      // default
      if (!string.IsNullOrEmpty(Default))
      {
        var defaultElem = mainDoc.CreateElement("default");
        if (H.IsWellFormedXml(Default))
        {
          defaultElem.InnerXml = Default;
        }
        else
        {
          defaultElem.InnerText = Default;
        }
        _ = varElem.AppendChild(defaultElem);
      }

      // prompt
      if (!string.IsNullOrEmpty(Prompt))
      {
        var promptElem = mainDoc.CreateElement("prompt");
        promptElem.InnerText = Prompt;
        _ = varElem.AppendChild(promptElem);
      }

      return varElem;
    }

  }


  public class VariablesMap
  {
    // Fields
    private static VariablesMap? _instance;
    private static readonly object _lock = new();

    // Properties
    public static VariablesMap Instance
    {
      get
      {
        if (_instance == null)
        {
          lock (_lock)
          {
            if (_instance == null)
            {
              _instance = new VariablesMap();
            }
          }
        }
        return _instance;
      }
    }

    public List<VariableData> Variables { get; private set; } = [];

    // Constructor
    private VariablesMap()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
      string vmFile = Path.GetFullPath(H.GetSProperty("VarMap"));
      if (!File.Exists(vmFile))
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
        _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", vmFile);
      }
      H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap", $" Utilizando el mapa de Variables:{vmFile} \n");

      string directoryPath = Path.GetDirectoryName(vmFile);
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
      string xmlFile = Path.Combine(directoryPath, "VariablesMap" + ".xml");
      DateTime fileModified = File.GetLastWriteTime(vmFile);
      DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;
      if (fileModified > xmlModified)
      {
        LoadFromXLS(vmFile);
        SaveToXml(xmlFile);
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap", $"XML file created at: {xmlFile}");
      }
      else
      {
        LoadFromXml(xmlFile);
      }
    }

    // Methods
    private void LoadFromXml(string xmlPath)
    {
      XmlDocument doc = new();
      doc.Load(xmlPath);
      foreach (XmlNode node in doc.SelectNodes("//variable")!)
      {
        VariableData data;
        try
        {
          data = new(
          node.Attributes["id"]!.InnerText ?? string.Empty,
          node.Attributes["varName"]?.InnerText ?? string.Empty,
          node.Attributes["area"]?.InnerText ?? string.Empty,
          node.Attributes["source"]!.InnerText ?? string.Empty,
          Convert.ToBoolean(H.ParseBoolean(node.Attributes["critic"]!.InnerText ?? "false")),
          Convert.ToBoolean(H.ParseBoolean(node.Attributes["mandatory"]!.InnerText ?? "false")),
          node.Attributes["type"]!.InnerText ?? string.Empty,
          node.Attributes["unit"]?.InnerText ?? string.Empty
          );

        }
        catch (Exception)
        {
          H.PrintLog(6,
              TC.ID.Value!.Time(),
              TC.ID.Value!.User,
              "VariablesMap.LoadFromXml",
              $"❌❌ Error ❌❌ Alguno de los valores obligatorios de la variable (id, source, critic, mandatory or type) {node.Attributes["id"]?.InnerText} no han sido declarados en VarMap");
          throw;
        }
        data.Default = H.IsWellFormedXml(node.SelectSingleNode("default")?.InnerXml)? node.SelectSingleNode("default")?.InnerXml: node.SelectSingleNode("default")?.InnerText ?? "";
        data.Description = node.SelectSingleNode("description")?.InnerText ?? "";
        data.Prompt = node.SelectSingleNode("prompt")?.InnerText ?? "";
        data.Deep = Convert.ToInt16(node.Attributes["deep"]?.InnerText ?? "0");
        var rangeNode = node.SelectSingleNode("allowableRange");
        if (rangeNode != null)
        {
          var values = new List<string>();
          foreach (XmlNode val in rangeNode.SelectNodes("value"))
          {
            values.Add(val.InnerText.Trim());
          }
          data.AllowableRange = values;
        }
        Variables.Add(data);
      }
    }

    public void SaveToXml(string xmlPath, List<string>? varList = null)
    {
      XmlDocument doc = new();
      XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
      _ = doc.AppendChild(xmlDeclaration);
      _ = doc.AppendChild(ToXml(doc, varList));
      doc.Save(xmlPath);
    }

    public XmlElement ToXml(XmlDocument doc, List<string>? varList = null)
    {
      XmlElement variables = doc.CreateElement("variables");
      foreach (var variable in Variables)
      {
        // If varList is provided, only append variables whose ID is in varList
        if (varList == null || varList.Contains(variable.ID))
        {
          _ = variables.AppendChild(variable.ToXML(doc));
        }
      }
      return variables;
    }

    private void LoadFromXLS(string vmFile)
    {
      DataSet dataSet;
      using (var stream = File.Open(vmFile, FileMode.Open, FileAccess.Read))
      {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
          using var _ = dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
          {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration()
            {
              UseHeaderRow = true
            }
          });
        }
      }
      DataTable dataTable = dataSet.Tables["VarMap"]!;

      for (int i = 0; i < dataTable.Rows.Count; i++)
      {
        DataRow row = dataTable.Rows[i];

        if (row.IsNull("ID"))
          break;

        try
        {
          VariableData data = new(
          row["ID"]!.ToString()!,
          row["NAME"]!.ToString()!,
          row["AREA"]!.ToString()!,
          row["SOURCE"]!.ToString()!,
          Convert.ToBoolean(H.ParseBoolean(row["CRITICAL"]!.ToString())),
          Convert.ToBoolean(H.ParseBoolean(row["MANDATORY"]!.ToString())),
          row["DATA TYPE"]?.ToString()!,
          row["UNIT"]?.ToString()!,
          row["DEFAULT VALUE"]?.ToString()!,
          row["DESCRIPTION"]?.ToString()!,
          row["PROMPT"]?.ToString()!
          );

          if (!row.IsNull("ALLOWABLE VALUE"))
          {
            var rangeList = row["ALLOWABLE VALUE"]!.ToString()!
                .Split(';')
                .Select(v => v.Trim())
                .ToList();

            data.AllowableRange = rangeList;
          }
          Variables.Add(data);
        }
        catch (Exception ex)
        {
          H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap.LoadFromXLS",
$"❌❌ Error ❌❌ -- Unexpected title in VariableMap/VarMap table \n{ex}");
          throw;
        }

      }

      //Check alloable variable types
      List<string> alloableVarTypes = [.. H.GetSProperty("allowableVarTypes").Split(';')];

      var nonAllowableTypeVars = Variables
          .Where(var => !alloableVarTypes.Contains(var.Type))
          .ToList();


      if (nonAllowableTypeVars.Count > 0)
      {
        foreach (var variable in nonAllowableTypeVars)
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap.LoadFromXLS",
          $"Variable type '{variable.Type}' declared for variable {variable.ID} is not valid.");
        }
        throw new ArgumentException("One or more variables have invalid types.");
      }

      // Check for non-normalized IDs
      var nonNormalizedIDs = Variables
          .Where(v => v.ID != v.ID.Normalize(NormalizationForm.FormD))
          .Select(v => v.ID)
          .ToList();

      if (nonNormalizedIDs.Count > 0)
      {
        foreach (var id in nonNormalizedIDs)
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "VariablesMap.LoadFromXLS",
              $"Variable ID '{id}' is not normalized.");
        }
        throw new ArgumentException("One or more variable IDs are not normalized.");
      }
    }

    public bool IsVariableExists(string id)
    {
      //trim id
      return Variables.Any(variable => variable.ID.Trim().ToLower() == id.Trim().ToLower());
    }

    public List<string> GetVarIDList()
    {
      List<string> varList = [];
      foreach (var variable in Variables)
      {
        varList.Add(variable.ID);
      }
      return varList;
    }

    public List<VariableData> GetAllVarList()
    {
      List<VariableData> varList = [.. Variables];
      return varList;
    }

    public List<VariableData> GetVarListBySource(string source)
    {
      return [.. Variables.Where(variable => variable.Source == source)];
    }

    public VariableData GetVariableData(string id)
    {
      return  Variables.FirstOrDefault(variable => variable.ID == id); 
    }

    public VariableData? GetNewVariableData(string id)
    {
      var varData = Variables.FirstOrDefault(variable => variable.ID == id);

      if (!id.Contains("\\s"))
      {
        if (varData == null)
        {
          Console.WriteLine($"VariableData with ID '{id}' was not found in Variable Map");

          return null;
        }

        return varData.Clone(); // Using the Clone method to return a new instance
      }
      return null;
    }

    public List<string> GetCalculationPath(List<string> varList)
    {
      foreach (string var in varList)
      {
      }
      return [];
    }

  }
}