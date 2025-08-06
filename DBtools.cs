using System.Xml;
using MySql.Data.MySqlClient;
using SmartBid;


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
                    CD_Date, CD_Request, CD_opportunityFolder, CD_opportunityID, CD_Client, CD_LocationCountry, 
                    CD_ProjectSize, CD_CreatedBy, CD_Status
                ) VALUES (
                    @Date, @Request, @OpportunityFolder, @OportunityID, @Client, @Country, 
                    @Size, @CreatedBy, @Status
                );
                SELECT LAST_INSERT_ID();";

      XmlNode requestDataNode = doc.SelectSingleNode(@"request/requestInfo");
      XmlNode projectDataNode = doc.SelectSingleNode(@"request/projectData");

      _ = cmd.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
      _ = cmd.Parameters.AddWithValue("@Request", requestDataNode?.Attributes["Type"]?.Value);
      _ = cmd.Parameters.AddWithValue("@OpportunityFolder", requestDataNode["opportunityFolder"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@OportunityID", projectDataNode["opportunityID"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@Client", projectDataNode?["client"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@Country", projectDataNode?["locationCountry"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@Size", projectDataNode["peakPower"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@CreatedBy", requestDataNode["createdBy"]?.InnerText ?? "");
      _ = cmd.Parameters.AddWithValue("@Status", "In progress");

      try
      {
        return Convert.ToInt32(cmd.ExecuteScalar());
      }
      catch (Exception ex)
      {
        H.PrintLog(5, "DBtools", "Error - InsertCallStart", "Error inserting callsTracker registry: " + ex.Message);
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
        H.PrintLog(5, "DBtools", "Error - UpdateCallRegistry", "Error updating callsTracker registry: " + ex.Message);
      }
    }
  }

  public static int InsertNewProjectWithBid(DataMaster dm)
  {
    XmlDocument dataMaster = dm.DM;
    using (var conn = DBConnectionFactory.CreateConnection())
    using (var transaction = conn.BeginTransaction())
    {
      try
      {
        // Extract XML nodes
        XmlNodeList inputDocs = dataMaster.SelectNodes(@"dm /utils/rev_01/inputDocs/doc");
        XmlNodeList deliveryDocs = dataMaster.SelectNodes(@"dm/utils/rev_01/deliveryDocs/doc");

        // Insert project
        var cmdProject = conn.CreateCommand();
        cmdProject.Transaction = transaction;
        cmdProject.CommandText = @"
                    INSERT INTO opportunities (
                        OPO_Opportunity_ID, OPO_Opportunity_Folder, OPO_Project_Name, OPO_Owner, OPO_Client, OPO_PeakPower, 
                        OPO_Location_Country, OPO_Location_Coordinates_Latitude, OPO_Location_Coordinates_Longitude,
                        OPO_Product, OPO_KAM, OPO_CreatedBy
                    ) VALUES (
                        @OportunityID, @OportunityFolder, @ProjectName, @Owner, @Client, 
                        @PeakPower, @Country, @Latitude, @Longitude, 
                        @Product, @KAM, @CreatedBy
                    );
                    SELECT LAST_INSERT_ID();";

        //_ = cmdProject.Parameters.AddWithValue("@OportunityID", dm.GetInnerText(@"dm/projectData/opportunityID"));
        _ = cmdProject.Parameters.AddWithValue("@OportunityID", dm.GetValueString("opportunityID"));
        _ = cmdProject.Parameters.AddWithValue("@OportunityFolder", dm.GetValueString("opportunityFolder"));
        _ = cmdProject.Parameters.AddWithValue("@ProjectName", dm.GetValueString("projectName"));
        _ = cmdProject.Parameters.AddWithValue("@Owner", dm.GetValueString("owner"));
        _ = cmdProject.Parameters.AddWithValue("@Client", dm.GetValueString("client"));
        _ = cmdProject.Parameters.AddWithValue("@PeakPower", dm.GetValueNumber("peakPower") ?? 0);
        _ = cmdProject.Parameters.AddWithValue("@Country", dm.GetValueString("locationCountry"));
        _ = cmdProject.Parameters.AddWithValue("@Latitude", dm.GetValueNumber("locationCoordinatesLatitude") ?? 0);
        _ = cmdProject.Parameters.AddWithValue("@Longitude", dm.GetValueNumber("locationCoordinatesLongitude") ?? 0);
        _ = cmdProject.Parameters.AddWithValue("@Product", dm.GetValueString("product"));
        _ = cmdProject.Parameters.AddWithValue("@KAM", dm.GetValueString("kam"));
        _ = cmdProject.Parameters.AddWithValue("@CreatedBy", dm.GetValueString("createdBy"));

        int projectId = Convert.ToInt32(cmdProject.ExecuteScalar());

        // Insert bid version
        var cmdBid = conn.CreateCommand();
        cmdBid.Transaction = transaction;
        cmdBid.CommandText = @"
                    INSERT INTO bidversion (
                        BV_Version, BV_Date, BV_CreatedBy, BV_Status, OPO_ID
                    ) VALUES (
                        @Version, @Date, @CreatedBy, @Status, @ProjectId
                    );
                    SELECT LAST_INSERT_ID();";

        _ = cmdBid.Parameters.AddWithValue("@Version", 1);
        _ = cmdBid.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        _ = cmdBid.Parameters.AddWithValue("@CreatedBy", dm.GetValueString("createdBy"));
        _ = cmdBid.Parameters.AddWithValue("@Status", "In progress");
        _ = cmdBid.Parameters.AddWithValue("@ProjectId", projectId);

        int bidVersionId = Convert.ToInt32(cmdBid.ExecuteScalar());

        // Insert input docs
        foreach (XmlNode docNode in inputDocs)
        {
          var cmdInput = conn.CreateCommand();
          cmdInput.Transaction = transaction;
          cmdInput.CommandText = "INSERT INTO inputdocs (ID_FileType, ID_FileName, ID_FileCheckSum, ID_FileLastModified, ID_BV_ID) VALUES (@FileType, @FileName, @ID_FileCheckSum, @ID_FileLastModified, @BV_ID);";

          _ = cmdInput.Parameters.AddWithValue("@FileType", docNode.Attributes["type"]?.Value ?? "");
          _ = cmdInput.Parameters.AddWithValue("@FileName", docNode?.InnerText ?? "UnnamedInput");
          _ = cmdInput.Parameters.AddWithValue("@ID_FileCheckSum", docNode.Attributes["hash"]?.Value ?? "");

          // Convertir la fecha al formato DateTime de MySQL
          DateTime parsedDate;
          if (DateTime.TryParse(docNode.Attributes["lastModified"]?.Value, out parsedDate))
            _ = cmdInput.Parameters.AddWithValue("@ID_FileLastModified", parsedDate.ToString("yyyy-MM-dd HH:mm:ss"));  // Formato correcto para MySQL DATETIME
          else
            _ = cmdInput.Parameters.AddWithValue("@ID_FileLastModified", DBNull.Value); // Manejo de error si la fecha no es válida
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
        H.PrintLog(5, "DBtools", "Error - InsertNewProjectWithBid", "Error during insert: " + ex.Message);
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
      Console.WriteLine("DBtools", "Error - LogMessage", "Error during insert: " + ex.Message);
    }
  }

  public static void InsertFileHash(string fileName, string type, string hash, string lastModified)
  {
    try
    {
      using var conn = DBConnectionFactory.CreateConnection();
      using var cmd = conn.CreateCommand();
      cmd.CommandText = "INSERT INTO inputFileHashs " +
        "       (IFH_time,   IFH_fileName, IFH_type, IFH_Checksum, IFH_lastModify) " +
        "VALUES (@Timestamp, @fileName,    @type,    @hash,    @lastModify)";
      _ = cmd.Parameters.AddWithValue("@Timestamp", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
      _ = cmd.Parameters.AddWithValue("@fileName", fileName);
      _ = cmd.Parameters.AddWithValue("@type", type);
      _ = cmd.Parameters.AddWithValue("@hash", hash);
      _ = cmd.Parameters.AddWithValue("@lastModify", lastModified);
      _ = cmd.ExecuteNonQuery();
    }
    catch (Exception ex)
    {
      H.PrintLog(5, "DBtools", "Error - InsertFileHash", "Error during insert inputFileHashs: " + ex.Message);
    }
  }


}

