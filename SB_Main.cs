using System.Collections.Concurrent;
using System.Diagnostics;
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

    // >>> NEW: listener management for two extra directories
    private static CancellationTokenSource _cts = new();
    private static Task? _listener1Task, _listener2Task;

    static void Main()
    {
      Console.OutputEncoding = System.Text.Encoding.UTF8;
      Console.WriteLine(
@"

     ██████\                                     ██\     ███████\  ██\       ██\ 
    ██  __██\                                    ██ |    ██  __██\ \__|      ██ |
    ██ /  \__|██████\████\   ██████\   ██████\ ██████\   ██ |  ██ |██\  ███████ |
    \██████\  ██  _██  _██\  \____██\ ██  __██\\_██  _|  ███████\ |██ |██  __██ |
     \____██\ ██ / ██ / ██ | ███████ |██ |  \__| ██ |    ██  __██\ ██ |██ /  ██ |
    ██\   ██ |██ | ██ | ██ |██  __██ |██ |       ██ |██\ ██ |  ██ |██ |██ |  ██ |
    \██████  |██ | ██ | ██ |\███████ |██ |       \████  |███████  |██ |\███████ |
     \______/ \__| \__| \__| \_______|\__|        \____/ \_______/ \__| \_______|
     

");
      string path = H.GetSProperty("callsPath");
      H.PrintLog(5, "00:00.000", "Main", "Main", $"Usando Varmap: {H.GetSProperty("VarMap")}");

      watcher = new FileSystemWatcher
      {
        Path = path,
        Filter = "*.*",
        NotifyFilter = NotifyFilters.FileName
      };
      watcher.Created += (sender, e) =>
      {
        H.PrintLog(5, "00:00.000", "Main", "Main", $"   ***    \nEvento detectado: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}");
        H.PrintLog(5, "00:00.000", "Main", "Main", $"{e.FullPath}");
        if (Regex.IsMatch(Path.GetFileName(e.FullPath), @"^call_.*\.xml$", RegexOptions.IgnoreCase))
        {
          _fileQueue.Enqueue(e.FullPath);
          _ = _eventSignal.Set();
        }
      };

      SB_Word.CloseWord(H.GetBProperty("closeWord"));
      SB_Excel.CloseExcel(H.GetBProperty("closeExcel"));

      watcher.EnableRaisingEvents = true;

      H.PrintLog(5, "00:00.000", "Main", "Main", $"Observando el directorio: {path}");

      // >>> NEW: Start two independent directory listeners
      string dir1 = H.GetSProperty("callsPathTemp");   // Directorio de entrada del Hermes
      string dir2 = H.GetSProperty("callsPathDWG");   // Directorio de escucha de la devolución del DWG

      _listener1Task = StartDirectoryListener(
        path: dir1,
        onNewFile: file =>
        {
          if (Regex.IsMatch(Path.GetFileName(file), @"^call_.*\.xml$", RegexOptions.IgnoreCase))
            DoStuff1(file);
          else
            H.PrintLog(5, "00:00.000", "SYSTEM", "Listener1", $"⚠️ Archivo ignorado (no call_*.xml): {file}");
        },
        token: _cts.Token,
        name: "Listener1"
      );

      _listener2Task = StartDirectoryListener(
        path: dir2,
        onNewFile: file =>
        {
          string ext = Path.GetExtension(file);
          if (ext.Equals(".dwg", StringComparison.OrdinalIgnoreCase) ||
              ext.Equals(".dxf", StringComparison.OrdinalIgnoreCase))
            DoStuff2(file);
          else
            H.PrintLog(5, "00:00.000", "SYSTEM", "Listener1", $"⚠️ Archivo ignorado (no call_*.xml): {file}");
        },
        token: _cts.Token,
        name: "Listener2"
      );


      H.PrintLog(5, "00:00.000", "SYSTEM", "Main", $"Listening: {dir1}");
      H.PrintLog(5, "00:00.000", "SYSTEM", "Main", $"Listening: {dir2}");

      H.PrintLog(5, "00:00.000", "Main", "Main", "Presiona 'Q' para salir...");

      // Procesamiento en un hilo separado
      _ = Task.Run(ProcessFiles);

      // Autorun si está configurado
      Thread.Sleep(400); // Espera para asegurar que el watcher esté listo
      if (!string.IsNullOrEmpty(H.GetSProperty("autorun")))
      {
        H.PrintLog(5, "00:00.000", "Main", "Main", $"Ejecutando Autorun: {H.GetSProperty("autorun")}\n" +
            $"Para ejectutar normalmente eliminar el valor en la propiedad 'autorun' en properties.xml\n\n");
        _ = Process.Start(H.GetSProperty("autorun"));
      }

      // Monitor de entrada para salir con 'Q'
      while (true)
      {
        if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
        {
          H.PrintLog(5, "00:00.000", "Main", "Main", "Salida solicitada... deteniendo el watcher.");
          watcher.EnableRaisingEvents = false; // Detiene la detección de archivos nuevos
          _stopRequested = true;

          // >>> NEW: stop the two extra listeners cleanly
          _cts.Cancel();
          try
          {
            _ = Task.WhenAll(
                _listener1Task ?? Task.CompletedTask,
                _listener2Task ?? Task.CompletedTask
            ).Wait(2000); // small grace period
          }
          catch { /* ignored on cancellation */ }

          break;
        }
        Thread.Sleep(1000); // Reduce la carga de la CPU
      }

      H.PrintLog(5, "00:00.000", "Main", "Main", "Todos los archivos han sido procesados. Programa terminado.");
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
            TC.ID.Value = null;
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
      TC.ID.Value = new TC.ThreadInfo(userName);

      // si autorun esperar 4 segundos a ejecutar
      if (H.GetBProperty("autorun"))
        Thread.Sleep(4000);

      H.PrintLog(5, "00:00.000", userName, "ProcessFile", $"Procesando archivo: {filePath}");

      int callID = DBtools.InsertCallStart(xmlCall); // Report starting process to DB
      List<ToolData> targets = Calculator.GetDeliveryDocs(xmlCall); // Get the delivery docs from the call
      DataMaster dm = CreateDataMaster(xmlCall, targets); //Create New DataMaster

      try
      {


        //Stores de call fileName in case configuration says so
        StoreCallFile(H.GetBProperty("storeXmlCall"), filePath, Path.GetDirectoryName(dm.FileName));

        Calculator calculator = new(dm, targets);
        calculator.RunCalculations();

        if (xmlCall.SelectSingleNode("/request/requestInfo")!.Attributes!["type"]?.Value == "create" && H.GetBProperty("createOppsFoldersStructure"))
          createOppsFoldersStructure(dm.GetInnerText("dm/projectData/opportunityID"), dm.GetInnerText("dm/utils/utilsData/opportunityFolder"));

        ReturnRemoveFiles(dm); // Returns or removes files depending on configuration

        DBtools.UpdateCallRegistry(callID, "DONE", "OK");

        string project = dm.GetValueString("opportunityFolder");

        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"-- ****{new string('*', project.Length + 17)}**** --");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"-- **** PROJECT: {project} DONE  **** --");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"-- ****{new string('*', project.Length + 17)}**** --");

        List<string> emailRecipients = [];
        // Add KAM email if configured to do so
        if (H.GetBProperty("mailKAM"))
          emailRecipients.Add(dm.GetValueString("kam"));
        // Add CreatedBy email if configured to do so
        if (H.GetBProperty("mailCreatedBy"))
          emailRecipients.Add(dm.GetValueString("createdBy"));


        string bodyHtml = $@"
          <p>Enviado desde SmartBid</p>

          <p>Le informamos de que la generación de los documentos técnicos relacionados con el proyecto <b>{dm.GetValueString("opportunityFolder")}</b> ha terminado con éxito.</p>

          <p>Podrá encontrar los documentos en la carpeta <b>2.ING&#8203;/{dm.SBidRevision}/OUTPUT</b> de la oportunidad (los documentos podrán tardar unos minutos en actualizarse en SharePoint).</p>

          <table style='font-family:Courier New, monospace; border-collapse:collapse;'>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Created by:</td>
                  <td>{TC.ID.Value!.User}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>KAM:</td>
                  <td>{dm.DM.SelectSingleNode("//projectData/kam")!.InnerText}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Call:</td>
                  <td>{Path.GetFileName(filePath)}</td>
              </tr>
              <tr style='border-top: 1px solid black; padding-top: 12px'>
                  <td style='vertical-align: top; font-weight: bold;'>Docs Generated:</td>
                  <td colspan='2'>{string.Join("<br/>", targets.Select(x => " " + x.Code))}
                  </td>
              </tr>
          </table>
          ";

        _ = H.MailTo( recipients:emailRecipients,
                      subject:$@"SmartBid: Opportunity:{dm.GetValueString("opportunityFolder")} revision: {dm.SBidRevision} DONE",
                      body: bodyHtml,
                      isHtml: true);

      }
      catch (Exception ex)
      {
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌--");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌ Error al procesar {dm.GetValueString("opportunityFolder")}❌❌");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"\uD83E\uDDE8 Excepción: {ex.GetType().Name}");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"\uD83D\uDCC4 Mensaje: {ex.Message}");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"\uD83E\uDDED StackTrace:\n{ex.StackTrace}");

        string bodyHtml = $@"
          <p>Enviado desde SmartBid</p>

          <p>El proceso de generación de la opp: <b>{dm.GetValueString("opportunityFolder")}</b> ha terminado con ERROR.</p>

          <table style='font-family:Courier New, monospace; border-collapse:collapse;'>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Error details:</td>
                  <td></td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Type:</td>
                  <td>{ex.GetType().Name}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Message:</td>
                  <td>{ex.Message}</td>
              </tr>
              <tr style='border-top: 1px solid black; padding-top: 12px'>
                  <td style='vertical-align: top; font-weight: bold;'>StackTrace:</td>
                  <td colspan='2'>{ex.StackTrace}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>User:</td>
                  <td>{TC.ID.Value!.User}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>Call:</td>
                  <td>{Path.GetFileName(filePath)}</td>
              </tr>
              <tr>
                  <td style='vertical-align: top; font-weight: bold;'>DataMaster:</td>
                  <td>{dm.FileName}</td>
              </tr>
          </table>
          ";

        _ = H.MailTo(H.GetSProperty("EngineeringEmail").Split(';').ToList(),
            subject: $"SmartBid Error processing {dm.GetValueString("opportunityFolder")}",
            body: bodyHtml, 
            isHtml: true);
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌❌--");
      }
    }

    private static DataMaster CreateDataMaster(XmlDocument xmlCall, List<ToolData> targets) //Creamos el datamaster
    {
      //Instantiating the DataMaster class with the XML string
      DataMaster dm = new(xmlCall, targets);
      //Creating the opportunityFolder in the storage directory
      string projectFolder = Path.Combine(H.GetSProperty("processPath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");
      if (!Directory.Exists(projectFolder))
        _ = Directory.CreateDirectory(projectFolder);

      dm.SaveDataMaster();
      return dm;
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
          H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "StoreCallFile", $"Call File '{callFile}' moved to '{targetDir}'.");
        }
        catch (Exception ex)
        {
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌ - StoreCallFile", $"❌Error❌ al mover '{callFile}': {ex.Message}");
        }
      }
      else
        File.Delete(callFile);
    }

    private static void ReturnRemoveFiles(DataMaster dm)
    {
      //string revisionDateStamp = dm.GetInnerText(@"dm/utils/rev_01/dateTime")[..6];
      string opportunityFolder = dm.GetInnerText(@"dm/utils/utilsData/opportunityFolder");

      string processedPath = Path.Combine(H.GetSProperty("processPath"),
                                          opportunityFolder,
                                          dm.SBidRevision);

      string processedToolsPath =   Path.Combine(processedPath, "TOOLS");
      string processedOutputsPath = Path.Combine(processedPath, "OUTPUT");
      string processedDocsPath =    Path.Combine(processedPath, "DOCS");


      string returnPath = Path.Combine(H.GetSProperty("oppsPath"),
                                            $"OFERTAS {dm.GetInnerText("dm/projectData/opportunityID").Substring(0, 4)}",
                                            opportunityFolder,
                                            @$"2.ING",
                                            dm.SBidRevision);

      string returnToolsPath =   Path.Combine(returnPath, "TOOLS");
      string returnOutputsPath = Path.Combine(returnPath, "OUTPUT");
      string returnDocsPath =    Path.Combine(returnPath, "DOCS");

      if (H.GetBProperty("returnTools"))
        foreach (string file in Directory.GetFiles(processedToolsPath))
        {
          Directory.CreateDirectory(returnToolsPath); // Create if it doesn't exists
          File.Copy(file, Path.Combine(returnToolsPath, Path.GetFileName(file)), overwrite: true);
        }

      if (!H.GetBProperty("storeTools"))
        foreach (string file in Directory.GetFiles(processedToolsPath))
          File.Delete(file);

      if (H.GetBProperty("returnDeliveries"))
        foreach (string file in Directory.GetFiles(processedOutputsPath))
        {
          Directory.CreateDirectory(returnOutputsPath); // Create if it doesn't exists
          File.Copy(file, Path.Combine(returnOutputsPath, Path.GetFileName(file)), overwrite: true);
        }

      if (!H.GetBProperty("storeDeliveries"))
        foreach (string file in Directory.GetFiles(processedOutputsPath))
          File.Delete(file);

      if (H.GetBProperty("createInputDocsShortcut"))
      {
        XmlNode inputDocsXML;
        try
        {
          inputDocsXML = dm.DM.SelectSingleNode($"/dm/utils/{dm.SBidRevision}/inputDocs");
          XmlDocument aaa = new XmlDocument();
          aaa.LoadXml("<root></root>");
          XmlNode importedNode = aaa.ImportNode(inputDocsXML, true);
          _ = aaa.DocumentElement.AppendChild(importedNode);
          H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ReturnRemoveFiles", $"Creating inputDocs shortcut in {returnPath}", aaa);
        }
        catch (Exception)
        {
          throw new Exception($"No inputDocs group found in DataMaster revision {dm.SBidRevision}");
        }

        _ = Directory.CreateDirectory(processedDocsPath);

        if (inputDocsXML != null)
        {
          foreach (XmlNode node in inputDocsXML.ChildNodes)
          {
            string filePath = node.InnerText.Trim();
            string shortcutPath = Path.Combine(processedDocsPath, $"{node.Name} - {Path.GetFileName(filePath)}.url");

            if (Path.Exists(filePath))
            {
              using (StreamWriter writer = new StreamWriter(shortcutPath))
              {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine($"URL={filePath}");
                writer.WriteLine("IconIndex=0");
                writer.WriteLine("IconFile=explorer.exe");
              }

            }
          }
        }
      }

      if (H.GetBProperty("returnDocsShortcuts"))
        foreach (string file in Directory.GetFiles(processedDocsPath))
        {
          _ = Directory.CreateDirectory(returnDocsPath); // Crea si no existe
          File.Copy(file, Path.Combine(returnDocsPath, Path.GetFileName(file)), overwrite: true);
        }

      if (H.GetBProperty("returnDataMaster"))
        File.Copy(dm.FileName, Path.Combine(returnPath, Path.GetFileName(dm.FileName)), overwrite: true);
    }

    private static void createOppsFoldersStructure(string opportunityID, string oppFolder)
    {
      string templatePath = H.GetSProperty("oppsFoldersTemplate");
      string baseDestinationPath = H.GetSProperty("oppsPath");
      string destinationPath = Path.Combine(baseDestinationPath,
                                            $"OFERTAS {opportunityID.Substring(0, 4)}",
                                            oppFolder);

      if (!Directory.Exists(templatePath))
      {
        Console.WriteLine($"Template path does not exist: {templatePath}");
        return;
      }

      // Create the destination folder
      _ = Directory.CreateDirectory(destinationPath);

      // Copy all subfolders and files
      CopyDirectory(templatePath, destinationPath);
    }

    private static void CopyDirectory(string sourceDir, string destinationDir)
    {
      // Create all subdirectories
      foreach (string dirPath in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
      {
        string newDirPath = dirPath.Replace(sourceDir, destinationDir);
        _ = Directory.CreateDirectory(newDirPath);
      }

      // Copy all files
      foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
      {
        string newFilePath = filePath.Replace(sourceDir, destinationDir);
        File.Copy(filePath, newFilePath, overwrite: true);
      }
    }


    // ============================
    // >>> NEW HELPERS & STUBS
    // ============================

    private static Task StartDirectoryListener(
        string path,
        Action<string> onNewFile,
        CancellationToken token,
        string name = "Listener")
    {
      return Task.Run(() =>
      {
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", name, $"⚠️ Path does not exist or is empty: '{path}'. Listener not started.");
          return;
        }

        using var fsw = new FileSystemWatcher
        {
          Path = path,
          Filter = "*.*",
          NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.Size,
          IncludeSubdirectories = false,
          EnableRaisingEvents = true
        };

        fsw.Created += (s, e) =>
        {
          // Fire-and-forget per file (still cancellable)
          _ = Task.Run(() =>
          {
            try
            {
              H.PrintLog(5, "00:00.000", "SYSTEM", name, $"Evento detectado: {e.FullPath}");
              if (WaitForFileReady(e.FullPath, attempts: 15, delay: TimeSpan.FromMilliseconds(300)))
              {
                onNewFile(e.FullPath); // trigger the specific handler
              }
              else
              {
                H.PrintLog(5, "00:00.000", "SYSTEM", name, $"⚠️ File never stabilized: {e.FullPath}");
              }
            }
            catch (Exception ex)
            {
              H.PrintLog(5, "00:00.000", "SYSTEM", name, $"❌ Error handling '{e.FullPath}': {ex.Message}");
            }
          }, token);
        };

        H.PrintLog(4, "00:00.000", "SYSTEM", name, $"Started watching: {path}");

        // Keep this task alive until cancellation
        using var done = new ManualResetEventSlim(false);
        using var reg = token.Register(() => done.Set());
        done.Wait();

        H.PrintLog(4, "00:00.000", "SYSTEM", name, "Stopping...");
      }, token);
    }

    private static bool WaitForFileReady(string path, int attempts, TimeSpan delay)
    {
      long lastSize = -1;

      for (int i = 0; i < attempts; i++)
      {
        try
        {
          var info = new FileInfo(path);
          if (!info.Exists)
          {
            Thread.Sleep(delay);
            continue;
          }

          long size = info.Length;

          // same size twice in a row and >0 -> likely stable
          if (size > 0 && size == lastSize)
          {
            // try to open for read to confirm
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return true;
          }

          lastSize = size;
        }
        catch
        {
          // swallow and retry
        }
        Thread.Sleep(delay);
      }
      return false;
    }


    private static void DoStuff1(string filePath)
    {
      // Mover directo a callsPath sin procesar DWG
      void MoveDirectToCallsPath()
      {
        string callsPath = H.GetSProperty("callsPath");
        if (string.IsNullOrWhiteSpace(callsPath))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1", "⚠️ 'callsPath' vacío en properties.xml");
          return;
        }
        _ = Directory.CreateDirectory(callsPath);

        string dest = Path.Combine(callsPath, Path.GetFileName(filePath));
        try
        {
          File.Move(filePath, dest);
        }
        catch (IOException)
        {
          // Cruce de volumen: Copy + Delete
          File.Copy(filePath, dest, overwrite: false);
          File.Delete(filePath);
        }
        H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff1", $"Archivo {Path.GetFileName(filePath)} movido a calls: {dest}");
      }

      // Si no se desea procesar DWG => directo a callsPath
      if (!H.GetBProperty("processDWG"))
      {
        MoveDirectToCallsPath();
        return;
      }

      // ===== Procesamiento DWG =====
      try
      {
        // 0) Leer PRIMERO el XML en su ubicación original para detectar LYOT
        string lyotPath = string.Empty;
        try
        {
          var xmlProbe = new XmlDocument();
          using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
          {
            xmlProbe.Load(stream);
          }
          var lyotNodeProbe = xmlProbe.SelectSingleNode("//request/requestInfo/inputDocs/doc[@type='LYOT']") as XmlElement;
          lyotPath = lyotNodeProbe?.InnerText?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1", $"⚠️ No se pudo leer el XML antes de mover: {ex.Message}");
          // Si ni siquiera podemos leer, comportarse como "no LYOT": mover directo a callsPath
          MoveDirectToCallsPath();
          return;
        }


        // Si no hay LYOT o el DWG no existe => directo a callsPath
        if (string.IsNullOrWhiteSpace(lyotPath) || !File.Exists(lyotPath))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1",
              $"⚠️ LYOT ausente o DWG inexistente. XML: {filePath} | LYOT: '{lyotPath}'. Se mueve directo a callsPath.");
          MoveDirectToCallsPath();
          return;
        }

        // ===== A partir de aquí, sí hay LYOT y se procesa el flujo DWG =====

        // 1) Mover el XML a la carpeta temporal
        string tempDir = H.GetSProperty("storageTemp");
        if (string.IsNullOrWhiteSpace(tempDir))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1", "⚠️ 'storageTemp' vacío en properties.xml");
          return;
        }
        _ = Directory.CreateDirectory(tempDir);
        string movedPath = Path.Combine(tempDir, Path.GetFileName(filePath));

        try { File.Move(filePath, movedPath); }
        catch (IOException) { File.Copy(filePath, movedPath, overwrite: false); File.Delete(filePath); }

        H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff1", $"Archivo movido a temp: {movedPath}");

        // 2) (Opcional) reabrir desde temp para cualquier otro uso del XML.
        //    Ya tenemos lyotPath de la lectura previa.
        string dwgFileName = Path.GetFileName(lyotPath);
        string dwgBaseName = Path.GetFileNameWithoutExtension(lyotPath);

        // 2.bis) Renombrar el XML a call_xxx_<dwgBaseName>.xml
        string currentXmlNameNoExt = Path.GetFileNameWithoutExtension(movedPath);
        if (!currentXmlNameNoExt.EndsWith($"_{dwgBaseName}", StringComparison.OrdinalIgnoreCase))
        {
          string renamedXml = Path.Combine(
              Path.GetDirectoryName(movedPath)!,
              $"{currentXmlNameNoExt}_{dwgBaseName}.xml"
          );
          File.Move(movedPath, renamedXml);
          movedPath = renamedXml;
          H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff1", $"XML renombrado a: {movedPath}");
        }

        // 3) Enviar el DWG por correo (H.MailTo ya soporta adjunto)
        string recipient = H.GetSProperty("DWG_recipiant");
        if (string.IsNullOrWhiteSpace(recipient))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1",
              "⚠️ 'DWG_recipiant' vacío en properties.xml; no se puede enviar correo.");
          return;
        }

        var recipients = new List<string> { recipient };
        string subject = "New DWG to prepare for SmartBid";
        string body = $"here is a new dwg file to identify the layers. " +
                      $"whenever is done, place a copy at {H.GetSProperty("callsPathDWG")}";

        bool sent = H.MailTo(recipients, subject, body, lyotPath);

        if (sent)
          H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff1", $"DWG enviado a {recipient}: {dwgFileName}");
        else
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1", $"❌ Error al enviar DWG a {recipient}: {dwgFileName}");
      }
      catch (Exception ex)
      {
        H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff1", $"❌ Excepción: {ex.GetType().Name} - {ex.Message}");
        // Como fallback último, intenta no bloquear el flujo
        try { MoveDirectToCallsPath(); } catch { /* ignorar */ }
      }
    }


    private static void DoStuff2(string returnedDwgPath)
    {
      // Utilidades locales
      static string EnsureUniqueFileName(string targetFullPath)
      {
        if (!File.Exists(targetFullPath)) return targetFullPath;
        string dir = Path.GetDirectoryName(targetFullPath)!;
        string name = Path.GetFileNameWithoutExtension(targetFullPath);
        string ext = Path.GetExtension(targetFullPath);
        string candidate;
        int i = 1;
        do
        {
          candidate = Path.Combine(dir, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}_{i}{ext}");
          i++;
        } while (File.Exists(candidate));
        return candidate;
      }

      static void MoveFileCrossVolume(string src, string dst, bool overwrite)
      {
        if (overwrite && File.Exists(dst)) File.Delete(dst);
        try { File.Move(src, dst); }
        catch (IOException)
        {
          // En caso de distinto volumen
          File.Copy(src, dst, overwrite: true);
          File.Delete(src);
        }
      }

      try
      {
        if (string.IsNullOrWhiteSpace(returnedDwgPath) || !File.Exists(returnedDwgPath))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2", $"⚠️ DWG no existe: {returnedDwgPath}");
          return;
        }

        string dwgBaseName = Path.GetFileNameWithoutExtension(returnedDwgPath);
        string tempDir = H.GetSProperty("storageTemp");
        string callsDir = H.GetSProperty("callsPath");

        if (string.IsNullOrWhiteSpace(tempDir) || string.IsNullOrWhiteSpace(callsDir))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2", "⚠️ 'storageTemp' o 'callsPath' vacíos en properties.xml");
          return;
        }

        // 1) Localizar el XML temporal que corresponde a este DWG: call_xxx_<dwgBaseName>.xml
        var candidates = Directory.GetFiles(tempDir, $"*_{dwgBaseName}.xml", SearchOption.TopDirectoryOnly);

        if (candidates.Length == 0)
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2",
              $"⚠️ No se encontró XML en TEMP que termine en _{dwgBaseName}.xml");
          return;
        }

        string? matchedXml = null;
        string? originalDwgPath = null;

        // 2) Confirmar el match leyendo LYOT y comparando el nombre del archivo
        foreach (var candidateXml in candidates)
        {
          try
          {
            var xml = new XmlDocument();
            using (var stream = new FileStream(candidateXml, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
              xml.Load(stream);
            }

            var lyotNode = xml.SelectSingleNode("//request/requestInfo/inputDocs/doc[@type='LYOT']") as XmlElement;
            string lyotPath = lyotNode?.InnerText?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(lyotPath)) continue;

            string lyotBaseName = Path.GetFileNameWithoutExtension(lyotPath);

            if (string.Equals(lyotBaseName, dwgBaseName, StringComparison.OrdinalIgnoreCase))
            {
              matchedXml = candidateXml;
              originalDwgPath = lyotPath;
              break;
            }
          }
          catch { /* probar el siguiente */ }
        }

        if (matchedXml is null || string.IsNullOrWhiteSpace(originalDwgPath))
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2",
              $"⚠️ No se encontró XML cuyo LYOT coincida con '{dwgBaseName}'.");
          //Enviar un correo electrónico al dwg_recipiant indicando el error y poniendo en el texto que revise el nombre del fichero para que 
          // coincida con el nombre del fichero original (enviar el nombre como recordatorio)
          _ = H.MailTo(H.GetSProperty("DWG_recipiant").Split(';').ToList(),
              subject: "Error: DWG name mismatch",
              body: $"The returned DWG file '{Path.GetFileName(returnedDwgPath)}' does not match any LYOT in the pending XML calls.\n" +
                    $"Please ensure the DWG file name matches the original LYOT name exactly.");
          return;
        }

        // 3) Respaldar el DWG original como *_original.dwg (si existe)
        if (File.Exists(originalDwgPath))
        {
          string dir = Path.GetDirectoryName(originalDwgPath)!;
          string name = Path.GetFileNameWithoutExtension(originalDwgPath);
          string ext = Path.GetExtension(originalDwgPath);
          string backup = Path.Combine(dir, $"{name}_original{ext}");
          backup = EnsureUniqueFileName(backup);

          MoveFileCrossVolume(originalDwgPath, backup, overwrite: false);
          H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff2", $"Original respaldado como: {backup}");
          Thread.Sleep(1000); // Pequeña espera para asegurar que el archivo se libera
        }
        else
        {
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2",
              $"⚠️ El DWG original indicado en LYOT no existe: {originalDwgPath}");
        }

        // 4) Mover el DWG devuelto a la ubicación original, manteniendo el nombre original
        string originalDir = Path.GetDirectoryName(originalDwgPath)!;
        string originalName = Path.GetFileName(originalDwgPath);
        string destinationDwg = Path.Combine(originalDir, originalName);

        MoveFileCrossVolume(returnedDwgPath, destinationDwg, overwrite: true);
        H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff2", $"DWG actualizado: {destinationDwg}");

        // 5) Mover el XML a callsPath con su nombre original "call_xxx.xml"
        // matchedXml se renombró en DoStuff1 a "call_xxx_<dwgBaseName>.xml"
        string xmlNameNoExt = Path.GetFileNameWithoutExtension(matchedXml);
        string suffix = "_" + dwgBaseName;

        string originalCallBase = xmlNameNoExt.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? xmlNameNoExt.Substring(0, xmlNameNoExt.Length - suffix.Length)
            : xmlNameNoExt; // fallback si no encontramos el sufijo (no debería ocurrir)

        string finalCallName = originalCallBase + ".xml";
        string finalCallPath = Path.Combine(callsDir, finalCallName);

        // Si ya existiera, hacemos único para no colisionar (pero conservamos el nombre original si es posible)
        if (File.Exists(finalCallPath))
        {
          string unique = Path.Combine(
              callsDir,
              $"{Path.GetFileNameWithoutExtension(finalCallName)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(finalCallName)}"
          );
          H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2",
              $"⚠️ Ya existe {finalCallName} en callsPath. Usando: {Path.GetFileName(unique)}");
          finalCallPath = unique;
        }

        MoveFileCrossVolume(matchedXml, finalCallPath, overwrite: false);

        H.PrintLog(4, "00:00.000", "SYSTEM", "DoStuff2",
            $"XML movido a callsPath para disparar SmartBid: {finalCallPath}");
      }
      catch (Exception ex)
      {
        H.PrintLog(5, "00:00.000", "SYSTEM", "DoStuff2", $"❌ Excepción: {ex.GetType().Name} - {ex.Message}");
      }
    }
  }

  static class TC // Thread Context
  {
    public class ThreadInfo
    {
      public Stopwatch chrono;

      public int ThreadId { get; }
      public string User { get; }


      public ThreadInfo(string user)
      {
        ThreadId = Environment.CurrentManagedThreadId;
        User = user;
        chrono = new Stopwatch();

        chrono.Start();
      }

      public string Time()
      {
        TimeSpan elapsed = chrono.Elapsed;
        return $"{(int)elapsed.TotalMinutes:D2}:{elapsed.Seconds:D2}.{elapsed.Milliseconds:D3}";
      }

    }
    // ✅ Ahora usamos AsyncLocal en lugar de ThreadLocal
    public static AsyncLocal<ThreadInfo> ID = new();
  }
}




