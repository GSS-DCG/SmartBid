using System.Data;
using System.Xml;
using ExcelDataReader;
using File = System.IO.File;

public class VarMap
{
    public List<VariableData> Variables { get; private set; } = new List<VariableData>();

    private void LoadFromXml(string xmlPath)
    {
        XmlDocument doc = new XmlDocument();
        doc.Load(xmlPath);

        foreach (XmlNode node in doc.SelectNodes("//variable"))
        {
            VariableData data = new VariableData
            (
                // Set attributes
                node.Attributes["ID"]?.InnerText,
                node.Attributes["varName"]?.InnerText,
                node.Attributes["area"]?.InnerText,
                node.Attributes["PrepTool"]?.InnerText,
                node.Attributes["critic"]?.InnerText,
                node.Attributes["mandatory"]?.InnerText,
                node.Attributes["type"]?.InnerText,
                node.Attributes["unit"]?.InnerText
            );

            // Set child elements
            data.SetDefault(node.SelectSingleNode("default")?.InnerText ?? "");
            data.SetDescription(node.SelectSingleNode("description")?.InnerText ?? "");

            // Parse allowableRange
            var rangeNode = node.SelectSingleNode("allowableRange");
            if (rangeNode != null)
            {
                var values = new List<string>();
                foreach (XmlNode val in rangeNode.SelectNodes("value"))
                {
                    values.Add(val.InnerText.Trim());
                }
                data.SetAllowableRange(values);
            }

            Variables.Add(data);
        }
    }

    private void SaveToXml(string xmlPath)
    {
        XmlDocument doc = new XmlDocument();
        XmlElement root = doc.CreateElement("root");
        _ = doc.AppendChild(root);

        foreach (var variable in Variables)
        {
            XmlElement varElem = doc.CreateElement("variable");

            varElem.SetAttribute("ID", variable.GetID());
            varElem.SetAttribute("varName", variable.GetVarName());
            varElem.SetAttribute("PrepTool", variable.GetPrepTool());
            varElem.SetAttribute("critic", variable.GetCritic());
            varElem.SetAttribute("mandatory", variable.GetMandatory());
            if (!string.IsNullOrEmpty(variable.GetArea()))
                varElem.SetAttribute("area", variable.GetArea());
            if (!string.IsNullOrEmpty(variable.GetType()))
                varElem.SetAttribute("type", variable.GetType());
            if (!string.IsNullOrEmpty(variable.GetUnit()))
                varElem.SetAttribute("unit", variable.GetUnit());

            // Add allowableRange if present
            var ranges = variable.GetAllowableRange();
            if (ranges != null && ranges.Count > 0)
            {
                XmlElement rangeElem = doc.CreateElement("allowableRange");
                foreach (var val in ranges)
                {
                    XmlElement valElem = doc.CreateElement("value");
                    valElem.InnerText = val;
                    _ = rangeElem.AppendChild(valElem);
                }
                _ = varElem.AppendChild(rangeElem);
            }

            // Add default if present
            if (!string.IsNullOrEmpty(variable.GetDefault()))
            {
                XmlElement defaultElem = doc.CreateElement("default");
                defaultElem.InnerText = variable.GetDefault();
                _ = varElem.AppendChild(defaultElem);
            }

            // Add description
            XmlElement descElem = doc.CreateElement("description");
            descElem.InnerText = variable.GetDescription() ?? "";
            _ = varElem.AppendChild(descElem);

            _ = root.AppendChild(varElem);
        }

        doc.Save(xmlPath);
        Console.WriteLine("**********XML create**********\n");
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

        DataTable dataTable = dataSet.Tables[0];

        // Iterate over the rows from row 3 until column A is empty
        for (int i = 2; i < dataTable.Rows.Count; i++)
        {
            DataRow row = dataTable.Rows[i];
            if (row.IsNull(0))
                break;

            VariableData data = new VariableData
            (
            row[0]?.ToString(),				        // ID           
            row[1]?.ToString(),	    		        // varName      
            row[2]?.ToString(),				        // area         
            row[3]?.ToString(),			            // PrepTool     
            row[19]?.ToString(),			        // critic       
            row[20]?.ToString(),			        // mandatory    
            row[22]?.ToString(),				    // type         
            row[23]?.ToString(),				    // unit         
            row[25]?.ToString(),			        // desciption   
            row[24]?.ToString().Replace(';', ',')	// defaultValue 
            );

            if (!row.IsNull(21))
            {
                var rangeList = new List<string>();
                foreach (var value in row[21].ToString().Split(';'))
                {
                    rangeList.Add(value.Trim());
                }
                data.SetAllowableRange(rangeList);
            }


            Variables.Add(data);




        }

        Console.WriteLine("XML file generated.");
    }

    public List<string> GetVarList()
    {
        List<string> varList = new List<string>();
        foreach (var variable in Variables)
        {
            varList.Add(variable.GetID());
        }
        return varList;
    }

    public VarMap(string vmFile)
    {
        // Register legacy encoding support
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


        if (!File.Exists(vmFile))
        {
            Console.WriteLine($"{vmFile} does not exist.");
            return;
        }

        string directoryPath = Path.GetDirectoryName(vmFile);
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(vmFile);
        string xmlFile = Path.Combine(directoryPath, fileNameWithoutExtension + ".xml");

        DateTime fileModified = File.GetLastWriteTime(vmFile);
        DateTime xmlModified = File.Exists(xmlFile) ? File.GetLastWriteTime(xmlFile) : default;

        if (fileModified > xmlModified)
        {
            LoadFromXLS(vmFile);
            SaveToXml(xmlFile);
            Console.WriteLine("Reading from Excel and creating the XML");
        }
        else
        {
            LoadFromXml(xmlFile);
            Console.WriteLine("The XML file is already up to date.");
        }

    }

    public static void Main(string[] args)
    {
        VarMap myClass = new VarMap("C:\\InSync\\Lab\\files\\VariableMap.xlsx");

        Console.WriteLine("Readed varialbes ID");
        myClass.GetVarList().ForEach(variable => Console.WriteLine(variable)); // No need for `ToList()`
    }



}


//public class VariableData
//{
//    private string ID;
//    private string varName;
//    private string area;
//    private string PrepTool;
//    private string critic;
//    private string mandatory;
//    private string type;
//    private string unit;
//    private string defaultValue;
//    private string description;
//    private List<string> allowableRange;

//    // Constructor opcional
//    public VariableData(
//        string ID,
//        string varName,
//        string area,
//        string PrepTool,
//        string critic,
//        string mandatory,
//        string type,
//        string unit = "",
//        string defaultValue = "",
//        string description = "",
//        List<string> allowableRange = null)
//    {
//        this.ID = ID;
//        this.varName = varName;
//        this.area = area;
//        this.PrepTool = PrepTool;
//        this.critic = critic;
//        this.mandatory = mandatory;
//        this.type = type;
//        this.unit = unit;
//        this.defaultValue = defaultValue;
//        this.description = description;
//        this.allowableRange = allowableRange ?? new List<string>();
//    }

//    public string GetID() => ID;
//    public void SetID(string value) => ID = value;

//    public string GetVarName() => varName;
//    public void SetVarName(string value) => varName = value;

//    public string GetArea() => area;
//    public void SetArea(string value) => area = value;

//    public string GetPrepTool() => PrepTool;
//    public void SetPrepTool(string value) => PrepTool = value;

//    public string GetCritic() => critic;
//    public void SetCritic(string value) => critic = value;

//    public string GetMandatory() => mandatory;
//    public void SetMandatory(string value) => mandatory = value;

//    public string GetType() => type;
//    public void SetType(string value) => type = value;

//    public string GetUnit() => unit;
//    public void SetUnit(string value) => unit = value;

//    public string GetDefault() => defaultValue;
//    public void SetDefault(string value) => defaultValue = value;

//    public string GetDescription() => description;
//    public void SetDescription(string value) => description = value;

//    public List<string> GetAllowableRange() => allowableRange;
//    public void SetAllowableRange(List<string> value) => allowableRange = value;
//}


public class VariableData
{
    public string ID { get; set; }
    public string VarName { get; set; }
    public string Area { get; set; }
    public string PrepTool { get; set; }
    public string Critic { get; set; }
    public string Mandatory { get; set; }
    public string Type { get; set; }
    public string Unit { get; set; }
    public string Default { get; set; }
    public string Description { get; set; }
    public List<string> AllowableRange { get; set; }

    // Constructor opcional
    public VariableData(
        string ID,
        string varName,
        string area,
        string PrepTool,
        string critic,
        string mandatory,
        string type,
        string unit = "",
        string defaultValue = "",
        string description = "",
        List<string> allowableRange = null)
    {
        this.ID = ID;
        this.VarName = varName;
        this.Area = area;
        this.PrepTool = PrepTool;
        this.Critic = critic;
        this.Mandatory = mandatory;
        this.Type = type;
        this.Unit = unit;
        this.Default = defaultValue;
        this.Description = description;
        this.AllowableRange = allowableRange ?? new List<string>();
    }
}
