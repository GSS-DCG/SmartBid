using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml;

namespace SmartBid
{
  class SB_Main
  {
    private static ConcurrentQueue<string> _fileQueue = new();
    private static AutoResetEvent _eventSignal = new(false);
    private static FileSystemWatcher? watcher;
    private static bool _stopRequested = false;

    static void Main()
    {
      Console.OutputEncoding = System.Text.Encoding.UTF8;
      Console.WriteLine(
        "                                                                             \r\n" +
        " ██████\\                                     ██\\     ███████\\  ██\\       ██\\ \r\n" +
        "██  __██\\                                    ██ |    ██  __██\\ \\__|      ██ |\r\n" +
        "██ /  \\__|██████\\████\\   ██████\\   ██████\\ ██████\\   ██ |  ██ |██\\  ███████ |\r\n" +
        "\\██████\\  ██  _██  _██\\  \\____██\\ ██  __██\\\\_██  _|  ███████\\ |██ |██  __██ |\r\n" +
        " \\____██\\ ██ / ██ / ██ | ███████ |██ |  \\__| ██ |    ██  __██\\ ██ |██ /  ██ |\r\n" +
        "██\\   ██ |██ | ██ | ██ |██  __██ |██ |       ██ |██\\ ██ |  ██ |██ |██ |  ██ |\r\n" +
        "\\██████  |██ | ██ | ██ |\\███████ |██ |       \\████  |███████  |██ |\\███████ |\r\n" +
        " \\______/ \\__| \\__| \\__| \\_______|\\__|        \\____/ \\_______/ \\__| \\_______|\r\n\n\n");

      string path = H.GetSProperty("callsPath");
      H.PrintLog(5, "Main", "Main", $"Usando Varmap: {H.GetSProperty("VarMap")}");

      watcher = new FileSystemWatcher
      {
        Path = path,
        Filter = "*.*",
        NotifyFilter = NotifyFilters.FileName
      };

      watcher.Created += (sender, e) =>
      {
        H.PrintLog(5, "Main", "Main", $"Evento detectado: {e.FullPath}");

        if (Regex.IsMatch(Path.GetFileName(e.FullPath), @"^call_\d+\.xml$", RegexOptions.IgnoreCase))
        {
          _fileQueue.Enqueue(e.FullPath);
          _ = _eventSignal.Set();
        }
      };


      SB_Word.CloseWord(H.GetBProperty("closeWord"));
      SB_Excel.CloseExcel(H.GetBProperty("closeExcel"));

      watcher.EnableRaisingEvents = true;
      H.PrintLog(5, "Main", "Main", $"Observando el directorio: {path}");
      H.PrintLog(5, "Main", "Main", "Presiona 'Q' para salir...");


      // Procesamiento en un hilo separado
      _ = Task.Run(ProcessFiles);

      //Autorun si está configurado
      Thread.Sleep(400); // Espera para asegurar que el watcher esté listo
      if (!string.IsNullOrEmpty(H.GetSProperty("autorun")))
      {
        H.PrintLog(5, "Main", "Main", $"Ejecutando Autorun: {H.GetSProperty("autorun")}\n" +
          $"Para ejectutar normalmente eliminar el valor en la propiedad 'autorun' en properties.xml\n\n");

        Process.Start(Path.Combine(H.GetSProperty("callsPath"), H.GetSProperty("autorun")));
      }

      // Monitor de entrada para salir con 'Q'
      while (true)
      {
        if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
        {
          H.PrintLog(5, "Main", "Main", "Salida solicitada... deteniendo el watcher.");
          watcher.EnableRaisingEvents = false; // Detiene la detección de archivos nuevos
          _stopRequested = true;
          break;
        }
        Thread.Sleep(1000); // Reduce la carga de la CPU
      }

      H.PrintLog(5, "Main", "Main", "Todos los archivos han sido procesados. Programa terminado.");
    }

    static void ProcessFiles()
    {
      while (!_stopRequested || !_fileQueue.IsEmpty) // Sigue procesando hasta vaciar la cola
      {
        _ = _eventSignal.WaitOne();

        while (_fileQueue.TryDequeue(result: out string? filePath))
        {
          // Procesamiento paralelo de cada archivo
          _ = Task.Run(() =>
          {
            ThreadContext.CurrentThreadInfo.Value = null;
            ProcessFile(filePath);
          });

        }
        Thread.Sleep(250); // Reduce la carga de la CPU
      }
    }

    static void ProcessFile(string filePath)
    {

      XmlDocument xmlCall = new();
      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      {
        xmlCall.Load(stream);
      }

      // ✅ Inicializar el contexto lógico (seguro para ejecución en paralelo)
      string userName = xmlCall.SelectSingleNode(@"request/requestInfo/createdBy")?.InnerText ?? "UnknownUser";
      ThreadContext.CurrentThreadInfo.Value = new ThreadContext.ThreadInfo(userName);

      H.PrintLog(5, userName, "ProcessFile", $"Procesando archivo: {filePath}");

      int callID = DBtools.InsertCallStart(xmlCall); // Report starting process to DB

      List<ToolData> targets = Calculator.GetDeliveryDocs(xmlCall); // Get the delivery docs from the call

      DataMaster dm = CreateDataMaster(xmlCall, targets); //Create New DataMaster
      try
      {
        // checks that all files declare exits and stores the checksum of the fileName for comparison
        ProcessInputFiles(dm, 1);

        //Stores de call fileName in case configuration says so
        StoreCallFile(H.GetBProperty("storeXmlCall"), filePath, Path.GetDirectoryName(dm.FileName));

        Calculator calculator = new(dm, targets);

        calculator.RunCalculations();

        ReturnRemoveFiles(dm); // Returns or removes files depending on configuration

        DBtools.UpdateCallRegistry(callID, "DONE", "OK");

        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--***************************************--");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--****||PROJECT: {dm.GetValueString("opportunityFolder")} DONE||****--");
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--***************************************--");

        //Auxiliar.DeleteBookmarkText("ES_Informe de corrosión_Rev0.0.docx", "Ruta_05", dm, "OUTPUT");


        List<string> emailRecipients = [];

        // Add KAM email if configured to do so
        if (H.GetBProperty("mailKAM"))
          emailRecipients.Add(dm.GetValueString("kam"));

        // Add CreatedBy email if configured to do so
        if (H.GetBProperty("mailCreatedBy"))
          emailRecipients.Add(dm.GetValueString("createdBy"));

        _ = H.MailTo(emailRecipients, "Mail de Prueba", "Enviado desde SmartBid");

      }
      catch (Exception ex)
      {
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌--");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--❌❌ Error al procesar {dm.GetValueString("opportunityFolder")}❌❌");

        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"🧨 Excepción: {ex.GetType().Name}");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"📄 Mensaje: {ex.Message}");
        H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"🧭 StackTrace:\n{ex.StackTrace}");
        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"--❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌--");
      }


    }

    private static DataMaster CreateDataMaster(XmlDocument xmlCall, List<ToolData> targets) //Creamos el datamaster
    {
      H.PrintXML(2, xmlCall); //Print the XML call for debugging

      //Instantiating the DataMaster class with the XML string 
      DataMaster dm = new(xmlCall, targets);

      //Creating the projectFolder in the storage directory
      string projectFolder = Path.Combine(H.GetSProperty("processPath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");

      if (!Directory.Exists(projectFolder))
        _ = Directory.CreateDirectory(projectFolder);

      dm.SaveDataMaster();

      return dm;
    }
    private static void ProcessInputFiles(DataMaster dm, int revision)
    {
      string inputPath = Path.Combine(H.GetSProperty("oppsPath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");



      foreach (XmlElement doc in dm.DM.SelectNodes(@$"dm/utils/rev_{revision.ToString("D2")}/inputDocs/doc"))
      {
        string fileType = doc.GetAttribute("type");
        string fileName = doc.InnerText;
        string inputFilesTimeStamp = dm.DM.SelectSingleNode(@$"dm/utils/rev_{revision.ToString("D2")}/inputDocs")?.Attributes?["timeStamp"]?.Value ?? DateTime.Now.ToString("yyMMdd");
        string filePath = Path.Combine(inputPath, "1.DOC", inputFilesTimeStamp, fileType, fileName);

        if (!File.Exists(filePath))
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, $"❌❌ Error ❌❌  - ProcessFile", $"⚠️ El archivo '{filePath}' no existe.");
          continue; // Saltar este documento y seguir con los demás
        }

        string hash = CalcularMD5(filePath); // Calculate MD5 hash for the fileName
        string lastModified = File.GetLastWriteTime(filePath).ToString("yyyy-MM-dd HH:mm:ss");

        doc.SetAttribute("hash", hash); // Set the hash attribute in the XML
        doc.SetAttribute("lastModified", lastModified); // Set the hash attribute in the XML

        DBtools.InsertFileHash(filePath, doc.GetAttribute("type"), hash, lastModified); // Store the fileName hash in the database

        H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"Archivo '{filePath}' registered");

      }
      H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "ProcessFile", $"All input files have been registered'.");
    }
    private static string CalcularMD5(string path)
    {
      using var stream = File.OpenRead(path);
      using var md5 = MD5.Create();
      byte[] hash = md5.ComputeHash(stream);
      return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
    private static void StoreCallFile(bool store, string callFile, string oppFolder)
    {
      if (store)
      {
        try
        {
          string fileName = $"{DateTime.Now:yyMMdd-HHmmss}_{Path.GetFileName(callFile)}";
          string targetDir = Path.Combine(H.GetSProperty("processPath"), "", "calls");

          File.Move(callFile, Path.Combine(targetDir, oppFolder, fileName));

          H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value!.User, "StoreCallFile", $"Call File '{callFile}' moved to '{targetDir}'.");
        }
        catch (Exception ex)
        {
          H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value!.User, $"❌❌ Error ❌❌  - StoreCallFile", $"❌Error❌ al mover '{callFile}': {ex.Message}");
        }
      }
      else
        File.Delete(callFile);
    }
    private static void ReturnRemoveFiles(DataMaster dm)
    {
      string revisionDateStamp = dm.GetInnerText(@"dm/utils/rev_01/dateTime")[..6];
      string projectFolder = dm.GetInnerText(@"dm/utils/utilsData/opportunityFolder");

      string processedToolsPath = Path.Combine(H.GetSProperty("processPath"), projectFolder, "TOOLS");
      string processedOutputsPath = Path.Combine(H.GetSProperty("processPath"), projectFolder, "OUTPUT");
      string oppsToolsPath = Path.Combine(H.GetSProperty("oppsPath"), projectFolder, @$"2.ING\{revisionDateStamp}\TOOLS");
      string oppsDeliveriesPath = Path.Combine(H.GetSProperty("oppsPath"), projectFolder, @$"2.ING\{revisionDateStamp}");


      if (H.GetBProperty("returnTools"))
        foreach (string file in Directory.GetFiles(processedToolsPath))
        {
          _ = Directory.CreateDirectory(oppsToolsPath); // Crea si no existe
          File.Copy(file, Path.Combine(oppsToolsPath, Path.GetFileName(file)), overwrite: true);
        }

      if (!H.GetBProperty("storeTools"))
        foreach (string file in Directory.GetFiles(processedToolsPath))
          File.Delete(file);


      if (H.GetBProperty("returnDeliveries"))
        foreach (string file in Directory.GetFiles(processedOutputsPath))
        {
          _ = Directory.CreateDirectory(oppsDeliveriesPath); // Crea si no existe
          File.Copy(file, Path.Combine(oppsDeliveriesPath, Path.GetFileName(file)), overwrite: true);
        }

      if (!H.GetBProperty("storeDeliveries"))
        foreach (string file in Directory.GetFiles(processedOutputsPath))
          File.Delete(file);

      if (H.GetBProperty("returnDataMaster"))
        File.Copy(dm.FileName, Path.Combine(H.GetSProperty("oppsPath"), projectFolder, Path.GetFileName(dm.FileName)), overwrite: true);
    }

  }

  static class ThreadContext
  {
    public class ThreadInfo
    {
      public int ThreadId { get; }
      public string User { get; }
      public ThreadInfo(string user)
      {
        ThreadId = Environment.CurrentManagedThreadId;
        User = user;
      }
    }

    // ✅ Ahora usamos AsyncLocal en lugar de ThreadLocal
    public static AsyncLocal<ThreadInfo> CurrentThreadInfo = new();
  }


}
