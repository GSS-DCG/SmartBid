
using System.Globalization;
using System.Xml;
using MySql.Data.MySqlClient;


public static class DBConnectionFactory
{
    private static readonly string ConnectionString = H.GetSProperty("DBConnextion");

    public static MySqlConnection CreateConnection()
    {
        var conn = new MySqlConnection(ConnectionString);
        conn.Open();
        return conn;
    }
}


static class DBtools
{
    public static int InsertCallStart(XmlDocument doc)
    {
        using (var conn = DBConnectionFactory.CreateConnection())
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                INSERT INTO callsTracker (
                    CD_Date, CD_ProjectName, CD_Client, CD_Location_Country, CD_Location_City,
                    CD_ProjectSize, CD_CreatedBy, CD_InputFolder,CD_Status, CD_Request
                ) VALUES (
                    @Date, @ProjectName, @Client, @Country, @City,
                    @Size, @CreatedBy, @InputFolder, @Status, @Request
                );
                SELECT LAST_INSERT_ID();";

            XmlNode requestDataNode = doc.SelectSingleNode(@"request/requestInfo");
            XmlNode projectDataNode = doc.SelectSingleNode(@"request/projectData");
            XmlNode locationNode = doc.SelectSingleNode(@"request/projectData/location");

            _ = cmd.Parameters.AddWithValue("@ProjectName", projectDataNode["projectName"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@Client", projectDataNode["client"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@Country", locationNode?["country"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@City", locationNode?["city"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@Size", double.TryParse(projectDataNode["projectSize"]?.InnerText, out double sizeVal) ? sizeVal : 0);
            _ = cmd.Parameters.AddWithValue("@CreatedBy", requestDataNode["createdBy"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@InputFolder", requestDataNode["inputFolder"]?.InnerText ?? "");
            _ = cmd.Parameters.AddWithValue("@Status", "In progress");
            _ = cmd.Parameters.AddWithValue("@Request", requestDataNode?.Attributes["Type"]?.Value);
            _ = cmd.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            try
            {
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
            catch (Exception ex)
            {
                H.PrintLog(2, "DBtools", "myEvent", "Error inserting call start: " + ex.Message);
                return -1;
            }
        }
    }

    public static void UpdateCallRegistry(int callID, string status, string result)
    {
        using (var conn = DBConnectionFactory.CreateConnection())
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = @"
                UPDATE callsTracker SET
                    CD_Status = @Status,
                    CD_result = @Result
                WHERE CD_ID = @CallID;";

            _ = cmd.Parameters.AddWithValue("@Status", status);
            _ = cmd.Parameters.AddWithValue("@Result", result);
            _ = cmd.Parameters.AddWithValue("@CallID", callID);

            try
            {
                _ = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                H.PrintLog(2, "DBtools", "myEvent", "Error updating call registry: " + ex.Message);
            }
        }
    }

    public static int InsertNewProjectWithBid(XmlDocument dm)
    {
        using (var conn = DBConnectionFactory.CreateConnection())
        using (var transaction = conn.BeginTransaction())
        {
            try
            {
                // Extract XML nodes
                XmlNode initDataNode = dm.SelectSingleNode(@"dm/initData");
                XmlNode locationNode = dm.SelectSingleNode(@"dm/initData/location");
                XmlNode geoCoordinatesNode = dm.SelectSingleNode(@"dm/initData/location/coordinates");
                XmlNode utilNode = dm.SelectSingleNode(@"dm/utils/rev_01/requestInfo");
                XmlNodeList inputDocs = dm.SelectNodes(@"dm/utils/rev_01/inputDocs/doc");
                XmlNodeList deliveryDocs = dm.SelectNodes(@"dm/utils/rev_01/deliveryDocs/doc");

                // Insert project
                var cmdProject = conn.CreateCommand();
                cmdProject.Transaction = transaction;
                cmdProject.CommandText = @"
                    INSERT INTO projects (
                        Pro_ProjectName, Pro_Location_Country, Pro_Location_Region, Pro_Location_Province,
                        Pro_Location_City, Pro_Location_PostalCode,
                        Pro_Location_Coordinates_Latitude, Pro_Location_Coordinates_Longitude,
                        Pro_ProjectSize, Pro_Product, Pro_KAM, PRO_CreatedBy
                    ) VALUES (
                        @ProjectName, @Country, @Region, @Province, @City, @PostalCode,
                        @Latitude, @Longitude, @Size, @Product, @KAM, @CreatedBy
                    );
                    SELECT LAST_INSERT_ID();";



                _ = cmdProject.Parameters.AddWithValue("@ProjectName", initDataNode["projectName"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@Country", locationNode?["country"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@Region", locationNode?["region"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@Province", locationNode?["province"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@City", locationNode?["city"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@PostalCode", locationNode?["postalCode"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@Latitude", decimal.TryParse(geoCoordinatesNode?["latitude"]?.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal lat) ? lat : 0);
                _ = cmdProject.Parameters.AddWithValue("@Longitude", decimal.TryParse(geoCoordinatesNode?["longitude"]?.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal lng) ? lng : 0);
                _ = cmdProject.Parameters.AddWithValue("@Size", decimal.TryParse(initDataNode?["projectSize"]?.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal size) ? size : 0);
                _ = cmdProject.Parameters.AddWithValue("@Product", initDataNode["product"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@KAM", initDataNode["kam"]?.InnerText ?? "");
                _ = cmdProject.Parameters.AddWithValue("@CreatedBy", utilNode["createdBy"]?.InnerText ?? "");

                int projectId = Convert.ToInt32(cmdProject.ExecuteScalar());

                // Insert bid version
                var cmdBid = conn.CreateCommand();
                cmdBid.Transaction = transaction;
                cmdBid.CommandText = @"
                    INSERT INTO bidversion (
                        BV_Version, BV_Date, BV_CreatedBy, BV_Status, project_id
                    ) VALUES (
                        @Version, @Date, @CreatedBy, @Status, @ProjectId
                    );
                    SELECT LAST_INSERT_ID();";

                _ = cmdBid.Parameters.AddWithValue("@Version", 1);
                _ = cmdBid.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _ = cmdBid.Parameters.AddWithValue("@CreatedBy", utilNode["createdBy"]?.InnerText ?? "");
                _ = cmdBid.Parameters.AddWithValue("@Status", "In progress");
                _ = cmdBid.Parameters.AddWithValue("@ProjectId", projectId);

                int bidVersionId = Convert.ToInt32(cmdBid.ExecuteScalar());

                // Insert input docs
                foreach (XmlNode docNode in inputDocs)
                {
                    var cmdInput = conn.CreateCommand();
                    cmdInput.Transaction = transaction;
                    cmdInput.CommandText = "INSERT INTO inputdocs (ID_FileName, ID_BV_ID) VALUES (@FileName, @BV_ID);";
                    _ = cmdInput.Parameters.AddWithValue("@FileName", docNode?.InnerText ?? "UnnamedInput");
                    _ = cmdInput.Parameters.AddWithValue("@BV_ID", bidVersionId);
                    _ = cmdInput.ExecuteNonQuery();
                }

                // Insert delivery docs
                foreach (XmlNode docNode in deliveryDocs)
                {
                    var cmdDelivery = conn.CreateCommand();
                    cmdDelivery.Transaction = transaction;
                    cmdDelivery.CommandText = "INSERT INTO deliverydocs (DD_Code, DD_BV_ID) VALUES (@Code, @BV_ID);";
                    _ = cmdDelivery.Parameters.AddWithValue("@Code", docNode?.InnerText ?? "UnnamedDelivery");
                    _ = cmdDelivery.Parameters.AddWithValue("@BV_ID", bidVersionId);
                    _ = cmdDelivery.ExecuteNonQuery();
                }

                transaction.Commit();
                return projectId;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                H.PrintLog(2, "DBtools", "myEvent", "Error during insert: " + ex.Message);
                return -1;
            }
        }
    }

    public static void LogMessage(int level, string user, string eventLog, string message)
    {
            try
            {
                using var conn = DBConnectionFactory.CreateConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO log (Log_time, Log_level, Log_event, Log_user, Log_message) VALUES (@Timestamp, @Level, @Event, @User, @Message)";
                _ = cmd.Parameters.AddWithValue("@Timestamp", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                _ = cmd.Parameters.AddWithValue("@Level", level);
                _ = cmd.Parameters.AddWithValue("@Event", eventLog);
                _ = cmd.Parameters.AddWithValue("@User", user);
                _ = cmd.Parameters.AddWithValue("@Message", message);
                _ = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine("DBtools", "Error-LogMessage", "Error during insert: " + ex.Message);
            }
    }

}
