using System.Xml;
using MySql.Data.MySqlClient;
using SmartBid;

class DBtools
{
    
    private MySqlConnection conn;

    public  DBtools()
    {
        string DBConnextion = H.GetSProperty("DBConnextion");
        conn = new MySqlConnection(DBConnextion);
        
            try
            {
                conn.Open();
                Console.WriteLine("Connection successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        
    }

    public int InsertCallStart(XmlDocument doc)
    {
        string query = @"
        INSERT INTO callsTracker (
            CD_ProjectName, CD_Client, CD_Location_Country, CD_Location_City,
            CD_ProjectSize, CD_CreatedBy, CD_Status, CD_Request
        ) VALUES (
            @ProjectName, @Client, @Country, @City,
            @Size, @CreatedBy, @Status, @Request
        );
        SELECT LAST_INSERT_ID();"; // Obtiene el ID del registro insertado

        using (MySqlCommand cmd = new MySqlCommand(query, conn))
        {
            XmlNode requestDataNode = doc.SelectSingleNode(@"dm/requestInfo ");
            XmlNode projectDataNode = doc.SelectSingleNode(@"dm/projectData");
            XmlNode locationNode = doc.SelectSingleNode(@"dm/projectData/Location");

            cmd.Parameters.AddWithValue("@CreatedBy", requestDataNode["createdBy"]?.InnerText ?? "");
            cmd.Parameters.AddWithValue("@ProjectName", projectDataNode["ProjectName"]?.InnerText ?? "");
            cmd.Parameters.AddWithValue("@Client", projectDataNode["Client"]?.InnerText ?? "");
            cmd.Parameters.AddWithValue("@Country", locationNode?["Country"]?.InnerText ?? "");
            cmd.Parameters.AddWithValue("@City", locationNode?["City"]?.InnerText ?? "");
            cmd.Parameters.AddWithValue("@Size", double.TryParse(projectDataNode["ProjectSize"].InnerText, out double sizeVal) ? sizeVal : 0);
            cmd.Parameters.AddWithValue("@Status", "In progress");
            cmd.Parameters.AddWithValue("@Request", requestDataNode?.Attributes["Type"]?.Value);

            try
            {
                object result = cmd.ExecuteScalar(); // Ejecuta la consulta y devuelve el ID
                int insertedId = Convert.ToInt32(result); // Convierte el resultado a entero

                Console.WriteLine($"Datos insertados correctamente en callsTracker con ID {insertedId}");
                return insertedId; // Devuelve el ID del nuevo registro
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al insertar datos: " + ex.Message);
                return -1; // Devuelve -1 si ocurre un error
            }
        }
    }

    public void UpdateCallRegistry(int callID, string status, string result)
    {
        string query = @$"
        UPDATE callsTracker SET
            CD_Status = @Status,
            CD_result = @Result
        WHERE CD_ID = @CallID;";
        using (MySqlCommand cmd = new MySqlCommand(query, conn))
        {
            cmd.Parameters.AddWithValue("@Status", status);
            cmd.Parameters.AddWithValue("@Result", result);
            cmd.Parameters.AddWithValue("@CallID", callID); 
            try
            {
                int rowsAffected = cmd.ExecuteNonQuery();
                Console.WriteLine($"Datos actualizados correctamente en callsTracker. Filas afectadas: {rowsAffected}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al actualizar datos: " + ex.Message);
            }
        }
    }

    public int InsertNewProjectWithBid(XmlDocument dm)
    {
        using (MySqlTransaction transaction = conn.BeginTransaction())
        {
            try
            {
                XmlNode initDataNode = dm.SelectSingleNode(@"dm/initData");
                XmlNode locationNode = dm.SelectSingleNode(@"dm/initData/Location");
                XmlNode geoCoordinatesNode = dm.SelectSingleNode(@"dm/initData/Location/Coordinates");
                XmlNode utilNode = dm.SelectSingleNode(@"dm/util");
                XmlNode projectDataNode = dm.SelectSingleNode(@"dm/data");
                XmlNodeList inputDocs = dm.SelectNodes(@"dm/utils/rev_01/inputDocs/doc");
                XmlNodeList deliveryDocs = dm.SelectNodes(@"dm/utils/rev_01/deliveryDocs/doc");

                // 1. Insert project
                string insertProject = @"
                INSERT INTO projects (
                    Pro_ProjectName, Pro_Location_Country, Pro_Location_Region, Pro_Location_Province,
                    Pro_Location_City, Pro_Location_PostalCode,
                    Pro_Location_Coordinates_Latitude, Pro_Location_Coordinates_Longitude,
                    Pro_ProjectSize, Pro_Product, Pro_KAM, PRO_CreatedBy
                ) VALUES (
                    @ProjectName, @Country, @Province, @City, @PostalCode,
                    @Latitude, @Longitude, @Size, @Product, @KAM, @CreatedBy
                );
                SELECT LAST_INSERT_ID();";
                MySqlCommand cmdProject = new MySqlCommand(insertProject, conn, transaction);
                cmdProject.Parameters.AddWithValue("@ProjectName", initDataNode["ProjectName"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@Country", locationNode?["Country"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@Region", locationNode?["Region"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@Province", locationNode?["Province"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@City", locationNode?["City"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@PostalCode", locationNode?["PostalCode"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@Latitude", decimal.TryParse(geoCoordinatesNode?["Latitude"]?.InnerText, out decimal lat) ? lat : 0);
                cmdProject.Parameters.AddWithValue("@Longitude", decimal.TryParse(geoCoordinatesNode?["Longitude"]?.InnerText, out decimal lng) ? lng : 0);
                cmdProject.Parameters.AddWithValue("@Size", int.TryParse(initDataNode["ProjectSize"]?.InnerText, out int size) ? size : 0);
                cmdProject.Parameters.AddWithValue("@Product", initDataNode["Product"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@KAM", initDataNode["KAM"]?.InnerText ?? "");
                cmdProject.Parameters.AddWithValue("@CreatedBy", utilNode["createdBy"]?.InnerText ?? "");
                int projectId = Convert.ToInt32(cmdProject.ExecuteScalar());

                // 2. Insert bidversion
                string insertBid = @"
    INSERT INTO bidversion (
        BV_Version, BV_Date, BV_CreatedBy, BV_Status, project_id
    ) VALUES (
        @Version, @Date, @CreatedBy, @Status, @ProjectId
    );
    SELECT LAST_INSERT_ID();";

                MySqlCommand cmdBid = new MySqlCommand(insertBid, conn, transaction);
                cmdBid.Parameters.AddWithValue("@Version", 1); // Assuming version 1 for the first bid
                cmdBid.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                cmdBid.Parameters.AddWithValue("@CreatedBy", utilNode["createdBy"]?.InnerText ?? "");
                cmdBid.Parameters.AddWithValue("@Status", "In progress");
                cmdBid.Parameters.AddWithValue("@ProjectId", projectId);

                int bidVersionId = Convert.ToInt32(cmdBid.ExecuteScalar());

                // 3. Insert inputdocs
                foreach (XmlNode docNode in inputDocs)
                {
                    string file = docNode["doc"]?.InnerText ?? "UnnamedInput";
                    string insertInput = "INSERT INTO inputdocs (ID_FileName, ID_BV_ID) VALUES (@FileName, @BV_ID);";
                    MySqlCommand cmdInput = new MySqlCommand(insertInput, conn, transaction);
                    cmdInput.Parameters.AddWithValue("@FileName", file);
                    cmdInput.Parameters.AddWithValue("@BV_ID", bidVersionId);
                    cmdInput.ExecuteNonQuery();
                }

                // 4. Insert deliverydocs
                foreach (XmlNode docNode in deliveryDocs)
                {
                    string doc = docNode["doc"]?.InnerText ?? "UnnamedDelivery";
                    string insertDelivery = "INSERT INTO deliverydocs (DD_Code, DD_BV_ID) VALUES (@Code, @BV_ID);";
                    MySqlCommand cmdDelivery = new MySqlCommand(insertDelivery, conn, transaction);
                    cmdDelivery.Parameters.AddWithValue("@Code", doc);
                    cmdDelivery.Parameters.AddWithValue("@BV_ID", bidVersionId);
                    cmdDelivery.ExecuteNonQuery();
                }

                transaction.Commit();
                Console.WriteLine($"Project {projectId} and BidVersion {bidVersionId} inserted successfully.");
                return projectId;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                Console.WriteLine("Error during insert: " + ex.Message);
                return -1;
            }
        }
    }

    private void QueryDatabase(string query)
    {
        // Ensure the connection is open
        if (conn.State != System.Data.ConnectionState.Open)
        {
            conn.Open();
        }
        // Example query
        query = "SELECT * FROM DeliveryDocs";
        using (MySqlCommand cmd = new MySqlCommand(query, conn))
        {
            using (MySqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    Console.WriteLine(reader["DD_Code"].ToString());
                }
            }
        }
    }
}