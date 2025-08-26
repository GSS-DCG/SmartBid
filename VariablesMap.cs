using System.Data;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Office2010.Excel;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using SmartBid;
using DataTable = System.Data.DataTable;
using File = System.IO.File;

namespace SmartBid
{

  public class VariableData(
      string id,
      string varName,
      string area,
      string source,
      bool critic,
      bool mandatory,
      string type,
      string unit = "",
      string defaultValue = "",
      string description = "",
      string prompt = "",
      int deep = 0,
      List<string>? allowableRange = null,
      string value = "")
  {
    #region Properties
    public string ID { get; set; } = id;
    public string VarName { get; set; } = varName;
    public string Area { get; set; } = area;
    public string Source { get; set; } = source;
    public bool Critic { get; set; } = critic;
    public bool Mandatory { get; set; } = mandatory;
    public string Type { get; set; } = type;
    public string Unit { get; set; } = unit;
    public string Default { get; set; } = defaultValue;
    public string Description { get; set; } = description;
    public string Prompt { get; set; } = prompt;
    public string? InOut { get; set; }
    public int? Call { get; set; }
    public int Deep { get; set; } = deep;
    public List<string> AllowableRange { get; set; } = allowableRange ?? [];
    public string Value { get; set; } = value;

    #endregion
    #region Constructor

    #endregion

    #region Methods

    public VariableData Clone()
    {

      return new VariableData(
          this.ID,
          this.VarName,
          this.Area,
          this.Source,
          this.Critic,
          this.Mandatory,
          this.Type,
          this.Unit,
          this.Default,
          this.Description,
          this.Prompt,
          this.Deep,
          [.. this.AllowableRange] // Ensure a new list instance
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
      varElem.SetAttribute("id", ID);
      varElem.SetAttribute("varName", VarName);
      varElem.SetAttribute("source", Source);
      varElem.SetAttribute("critic", Critic.ToString());
      varElem.SetAttribute("mandatory", Mandatory.ToString());
      varElem.SetAttribute("deep", Deep.ToString());

      if (!string.IsNullOrEmpty(Area))
        varElem.SetAttribute("area", Area);
      if (!string.IsNullOrEmpty(Type))
        varElem.SetAttribute("type", Type);
      if (!string.IsNullOrEmpty(Unit))
        varElem.SetAttribute("unit", Unit);

      //Add description
      XmlElement descriptionElem = mainDoc.CreateElement("description");
      descriptionElem.InnerText = Description;
      _ = varElem.AppendChild(descriptionElem);

      // Add allowableRange if present
      var ranges = AllowableRange;
      if (ranges != null && ranges.Count > 0)
      {
        XmlElement rangeElem = mainDoc.CreateElement("allowableRange");
        foreach (var val in ranges)
        {
          XmlElement valElem = mainDoc.CreateElement("value");
          valElem.InnerText = val;
          _ = rangeElem.AppendChild(valElem);
        }
        _ = varElem.AppendChild(rangeElem);
      }

      // Add default if present
      if (!string.IsNullOrEmpty(Default))
      {
        XmlElement defaultElem = mainDoc.CreateElement("default");
        defaultElem.InnerText = Default;
        _ = varElem.AppendChild(defaultElem);
      }

      // Add promt if present
      if (!string.IsNullOrEmpty(Prompt))
      {
        XmlElement promptElem = mainDoc.CreateElement("prompt");
        promptElem.InnerText = Prompt;
        _ = varElem.AppendChild(promptElem);
      }

      return varElem;
    }

    #endregion
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
        H.PrintLog(2, "VarMap", "VariablesMap", $" ****** FILE: {vmFile} NOT FOUND. ******\n Review value 'VarMap' in properties.xml it should point to the location of the Variables Map file.\n\n");
        _ = new FileNotFoundException("PROPERTIES FILE NOT FOUND", vmFile);
      }
      H.PrintLog(2, "VarMap", "VariablesMap", $" Utilizando el mapa de Variables:{vmFile} \n");

      string directoryPath = Path.GetDirectoryName(vmFile);
      string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
      string xmlFile = Path.Combine(directoryPath, "VariablesMap" + ".xml");
      DateTime fileModified = File.GetLastWriteTime(vmFile);
      DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;
      if (fileModified > xmlModified)
      {
        LoadFromXLS(vmFile);
        SaveToXml(xmlFile);
        H.PrintLog(2, "VarMap", "VariablesMap", $"XML file created at: {xmlFile}");
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
        VariableData data = new(
            node.Attributes["id"]!.InnerText ?? string.Empty,
            node.Attributes["varName"]?.InnerText ?? string.Empty,
            node.Attributes["area"]!.InnerText ?? string.Empty,
            node.Attributes["source"]!.InnerText ?? string.Empty,
            Convert.ToBoolean(node.Attributes["critic"]!.InnerText ?? "false"),
            Convert.ToBoolean(node.Attributes["mandatory"]!.InnerText ?? "false"),
            node.Attributes["type"]!.InnerText ?? string.Empty,
            node.Attributes["unit"]?.InnerText ?? string.Empty
        );

        data.Default = node.SelectSingleNode("default")?.InnerText ?? "";
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
      XmlDocument doc = ToXml(varList);
      doc.Save(xmlPath);
    }

    public XmlDocument ToXml(List<string>? varList = null)
    {
      XmlDocument doc = new();
      XmlElement root = doc.CreateElement("root");
      _ = doc.AppendChild(root);
      foreach (var variable in Variables)
      {
        // If varList is provided, only append variables whose ID is in varList
        if (varList == null || varList.Contains(variable.ID))
        {
          _ = root.AppendChild(variable.ToXML(doc));
        }
      }
      return doc;
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
          Convert.ToBoolean(row["CRITICAL"]!.ToString()), 
          Convert.ToBoolean(row["MANDATORY"]!.ToString()),
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
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "VariablesMap.LoadFromXLS",
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
          H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "VariablesMap.LoadFromXLS",
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
          H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "VariablesMap.LoadFromXLS",
              $"Variable ID '{id}' is not normalized.");
        }
        throw new ArgumentException("One or more variable IDs are not normalized.");
      }
    }

    public bool IsVariableExists(string id)
    {
      return Variables.Any(variable => variable.ID == id);
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
      return Variables.FirstOrDefault(variable => variable.ID == id);
    }

    public VariableData? GetNewVariableData(string id)
    {
      var varData = Variables.FirstOrDefault(variable => variable.ID == id);

      if (!id.Contains("\\s"))
      {
        if (varData == null)
        {
          Console.WriteLine($"VariableData with ID '{id}' was not found");

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