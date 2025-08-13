using System.Data;
using System.Text;
using System.Xml;
using ExcelDataReader;
using SmartBid;
using File = System.IO.File;


public class VariableData
{
  #region Properties
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
  public string InOut { get; set; }
  public int Call { get; set; }
  public int Deep { get; set; }
  public List<string> AllowableRange { get; set; }
  public string Value { get; set; } // Current value of the variable

  #endregion

  #region Constructor
  public VariableData(
      string ID,
      string varName,
      string area,
      string source,
      bool critic,
      bool mandatory,
      string type,
      string unit = "",
      string defaultValue = "",
      string description = "",
      int deep = 0,
      List<string> allowableRange = null,
      string value = "")
  {
    this.ID = ID;
    this.VarName = varName;
    this.Area = area;
    this.Source = source;
    this.Critic = critic;
    this.Mandatory = mandatory;
    this.Type = type;
    this.Unit = unit;
    this.Default = defaultValue;
    this.Description = description;
    this.AllowableRange = allowableRange ?? new List<string>();
    this.Value = value;
  }

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
        this.Deep,
        new List<string>(this.AllowableRange) // Ensure a new list instance
    );
  }
  public XmlDocument ToXMLDocument()
  {
    XmlDocument doc = new XmlDocument();
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

    return varElem;
  }

  #endregion
}

public class VariablesMap
{
  // Fields
  private static VariablesMap _instance;
  private static readonly object _lock = new object();

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

  public List<VariableData> Variables { get; private set; } = new List<VariableData>();

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
    XmlDocument doc = new XmlDocument();
    doc.Load(xmlPath);
    foreach (XmlNode node in doc.SelectNodes("//variable"))
    {
      VariableData data = new VariableData(
          node.Attributes["id"]?.InnerText ?? string.Empty,
          node.Attributes["varName"]?.InnerText ?? string.Empty,
          node.Attributes["area"]?.InnerText ?? string.Empty,
          node.Attributes["source"]?.InnerText ?? string.Empty,
          Convert.ToBoolean(node.Attributes["critic"]?.InnerText ?? "false"),
          Convert.ToBoolean(node.Attributes["mandatory"]?.InnerText ?? "false"),
          node.Attributes["type"]?.InnerText ?? string.Empty,
          node.Attributes["unit"]?.InnerText ?? string.Empty
      );

      data.Default = node.SelectSingleNode("default")?.InnerText ?? "";
      data.Description = node.SelectSingleNode("description")?.InnerText ?? "";
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
    XmlDocument doc = new XmlDocument();
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
    DataTable dataTable = dataSet.Tables["VarMap"];
    // Iterate over the rows from row 3 until column A is empty
    for (int i = 2; i < dataTable.Rows.Count; i++)
    {
      DataRow row = dataTable.Rows[i];
      if (row.IsNull(0))
        break;
      VariableData data = new VariableData(
          row[0]?.ToString(), // ID
          row[1]?.ToString(), // varName
          row[2]?.ToString(), // area
          row[3]?.ToString(), // Source 
          Convert.ToBoolean(row[19]?.ToString()), // critic
          Convert.ToBoolean(row[20]?.ToString()), // mandatory
          row[22]?.ToString(), // type
          row[23]?.ToString(), // unit
          row[24]?.ToString(), // defaultValue
          row[25]?.ToString()  // description
      );
      if (!row.IsNull(22))
      {
        var rangeList = new List<string>();
        foreach (var value in row[21].ToString().Split(';')) //AllowableRange
        {
          rangeList.Add(value.Trim());
        }
        data.AllowableRange = rangeList;
      }
      Variables.Add(data);
    }

    //Check alloable variable types
    List<string> alloableVarTypes = H.GetSProperty("allowableVarTypes").Split(';').ToList();

    var nonAllowableTypeVars = Variables
        .Where(var => !alloableVarTypes.Contains(var.Type))
        .ToList();


    if (nonAllowableTypeVars.Count > 0)
    {
      foreach (var variable in nonAllowableTypeVars)
      {
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "VariablesMap.LoadFromXLS",
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
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "VariablesMap.LoadFromXLS",
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
    List<string> varList = new List<string>();
    foreach (var variable in Variables)
    {
      varList.Add(variable.ID);
    }
    return varList;
  }

  public List<VariableData> GetAllVarList()
  {
    List<VariableData> varList = new List<VariableData>();
    foreach (var variable in Variables)
    {
      varList.Add(variable);
    }
    return varList;
  }

  public List<VariableData> GetVarListBySource(string source)
  {
    return Variables.Where(variable => variable.Source == source).ToList();
  }

  public VariableData GetVariableData(string id)
  {
    return Variables.FirstOrDefault(variable => variable.ID == id);

  }

  public VariableData GetNewVariableData(string id)
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
    return new List<string>();
  }

}
