using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml;

namespace SmartBid
{
  class SB_Main
  {
    // ============================
    // Campos existentes
    // ============================
    private static ConcurrentQueue<string> _fileQueue = new();
    private static AutoResetEvent _eventSignal = new(false);
    private static FileSystemWatcher? watcher;
    private static bool _stopRequested = false;
    private static readonly ConcurrentDictionary<string, DateTime> _recentlyProcessed = new();

    // Listeners adicionales
    private static CancellationTokenSource _cts = new();
    private static Task? _listener1Task, _listener2Task;

    // CAMPO para el seguimiento de tareas
    private static int _activeFileProcessors = 0;

    // ============================
    // Instancia única (por sesión de usuario)
    // ============================
    private const string SingleInstanceMutexName = @"Local\SmartBid.SB_Main.SingleInstance";
    private static Mutex? _singleInstanceMutex;

    private static bool EnsureSingleInstance()
    {
      bool createdNew;
      _singleInstanceMutex = new Mutex(initiallyOwned: true, name: SingleInstanceMutexName, createdNew: out createdNew);
      if (!createdNew)
      {
        var msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ya hay una instancia de SmartBid en ejecución. Saliendo.";
        Console.Error.WriteLine(msg);
        try { Console.Beep(); } catch { /* sin audio en algunos entornos */ }
        try { H.PrintLog(2, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "Main", msg); } catch { /* H aún no inicializado */ }
        Thread.Sleep(2000);
        Environment.ExitCode = 1;
        return false;
      }
      return true;
    }

    static void Main()
    {
      // === INSTANCIA ÚNICA ===
      if (!EnsureSingleInstance()) return;

      try
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
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Usando Varmap: {H.GetSProperty("VarMap")}");

        watcher = new FileSystemWatcher
        {
          Path = path,
          Filter = "*.*",
          NotifyFilter = NotifyFilters.FileName
        };

        watcher.Created += (sender, e) =>
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"*** Evento detectado: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Call detected: {e.FullPath}");

          if (_recentlyProcessed.ContainsKey(e.FullPath))
          {
            if (DateTime.Now.Subtract(_recentlyProcessed[e.FullPath]).TotalSeconds < 2)
            {
              H.PrintLog(2, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Evento duplicado para {e.FullPath}. Ignorando.");
              return;
            }
          }
          _recentlyProcessed[e.FullPath] = DateTime.Now;

          if (Regex.IsMatch(Path.GetFileName(e.FullPath), @"^call_.*\.xml$", RegexOptions.IgnoreCase))
          {
            _fileQueue.Enqueue(e.FullPath);
            _eventSignal.Set();
          }
        };

        SB_Word.CloseWord(H.GetBProperty("closeWord"));
        SB_Excel.CloseExcel(H.GetBProperty("closeExcel"));

        watcher.EnableRaisingEvents = true;
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Observando el directorio: {path}");

        string dir1 = H.GetSProperty("callsPathTemp");
        string dir2 = H.GetSProperty("callsPathDWG");

        _listener1Task = StartDirectoryListener(path: dir1, onNewFile: file => { if (Regex.IsMatch(Path.GetFileName(file), @"^call_.*\.xml$", RegexOptions.IgnoreCase)) DoStuff1(file); else H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "Listener1", $"⚠️ Archivo ignorado (no call_*.xml): {file}"); }, token: _cts.Token, name: "Listener1");
        _listener2Task = StartDirectoryListener(path: dir2, onNewFile: file => { string ext = Path.GetExtension(file); if (ext.Equals(".dwg", StringComparison.OrdinalIgnoreCase) || ext.Equals(".dxf", StringComparison.OrdinalIgnoreCase)) DoStuff2(file); else H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "Listener2", $"⚠️ Archivo ignorado (no .dwg/.dxf): {file}"); }, token: _cts.Token, name: "Listener2");

        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "Main", $"Listening: {dir1}");
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "Main", $"Listening: {dir2}");
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", "Presiona 'Q' para salida progresiva o 'Ctrl+Q' para forzar salida...");

        _ = Task.Run(ProcessFiles);

        Thread.Sleep(400);
        if (!string.IsNullOrEmpty(H.GetSProperty("autorun")))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Ejecutando Autorun: {H.GetSProperty("autorun")}\n" + "Para ejectutar normalmente eliminar el valor en la propiedad 'autorun' en properties.xml\n\n");
          _ = Process.Start(H.GetSProperty("autorun"));
        }

        // ============================================
        // Iniciar el listener de teclado dedicado en un hilo separado
        var inputTask = Task.Run(ListenForExitKeys);

        while (!_stopRequested)
        {
          Thread.Sleep(1000);
        }

        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", "Fase de finalización iniciada. Esperando a que las tareas en curso finalicen...");

        try
        {
          // Esperar a que los listeners se detengan
          Task.WhenAll(_listener1Task ?? Task.CompletedTask, _listener2Task ?? Task.CompletedTask).Wait(3000);
        }
        catch (Exception ex) when (ex is OperationCanceledException || ex is AggregateException)
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", "Listeners adicionales detenidos.");
        }

        // Esperar a que los procesadores de archivos terminen
        while (_activeFileProcessors > 0)
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", $"Esperando a que finalicen {_activeFileProcessors} procesos de archivo...");
          Thread.Sleep(2000);
        }

        _eventSignal.Set(); // Señal final para asegurar que ProcessFiles sale

        DBtools.ResetAllCallRegistries();
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", "Todos los archivos han sido procesados. Programa terminado.");
      }
      finally
      {
        try { _singleInstanceMutex?.ReleaseMutex(); } catch { }
        _singleInstanceMutex?.Dispose();
        H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Main", "*** Aplicación terminada ***");
      }
    }

    // ==========================================
    // ========= LISTENER DE TECLADO ============
    // ==========================================
    /// <summary>
    /// Se ejecuta en una tarea en segundo plano para escuchar las pulsaciones de teclas
    /// sin interferir con el hilo principal.
    /// </summary>
    private static void ListenForExitKeys()
    {
      while (!_stopRequested)
      {
        try
        {
          // Console.ReadKey(true) es una llamada bloqueante. Espera aquí hasta que se pulsa una tecla.
          var keyInfo = Console.ReadKey(intercept: true);

          // Comprobar Ctrl+Q para salida forzada
          if (keyInfo.Key == ConsoleKey.Q && (keyInfo.Modifiers & ConsoleModifiers.Control) != 0)
          {
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "=============================================================");
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "Salida forzada solicitada (Ctrl+Q). Terminando inmediatamente.");
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "=============================================================");
            Thread.Sleep(500); // Dar tiempo a que el log aparezca
            Environment.Exit(1);
          }
          // Comprobar Q para salida progresiva
          else if (keyInfo.Key == ConsoleKey.Q)
          {
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "============================================================================================");
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "Salida progresiva solicitada ('Q'). Finalizando cuando acaben todos los procesos en curso...");
            H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "============================================================================================");

            // Iniciar la secuencia de apagado progresivo
            watcher.EnableRaisingEvents = false;
            _cts.Cancel();
            _stopRequested = true;
            _eventSignal.Set(); // Despertar a ProcessFiles si está esperando

            // El bucle terminará porque _stopRequested es ahora true
          }
        }
        catch (InvalidOperationException)
        {
          // Esto ocurre si la aplicación se ejecuta sin una consola (p. ej., como un servicio).
          // El listener de teclado no puede funcionar, así que simplemente salimos de este hilo.
          H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "Main", "Input", "No se puede leer desde la consola. El listener de teclado se detiene.");
          break;
        }
      }
    }


    static void ProcessFiles()
    {
      while (!_stopRequested)
      {
        _eventSignal.WaitOne();

        while (_fileQueue.TryDequeue(out string? filePath))
        {
          if (_stopRequested)
          {
            H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "Main", "ProcessFiles", $"Parada solicitada. Se omite el procesamiento de: {filePath}");
            continue;
          }

          _ = Task.Run(() =>
          {
            try
            {
              Interlocked.Increment(ref _activeFileProcessors);
              ProcessFile(filePath);
            }
            finally
            {
              Interlocked.Decrement(ref _activeFileProcessors);
            }
          });
        }
      }

      H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "Main", "ProcessFiles", "Bucle de procesamiento detenido.");
      if (!_fileQueue.IsEmpty)
      {
        H.PrintLog(2, DateTime.Now.ToString("HH:mm:ss"), "Main", "ProcessFiles", $"AVISO: {_fileQueue.Count} archivos en cola no serán procesados debido a la parada del servicio.");
        _fileQueue.Clear();
      }
    }

    static void ProcessFile(string filePath)
    {
      XmlDocument xmlCall = new();
      using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      {
        xmlCall.Load(stream);
      }

      string userName = xmlCall.SelectSingleNode(@"request/requestInfo/createdBy")?.InnerText ?? "UnknownUser";

      int callID = DBtools.InsertCallStart(xmlCall);
      DataMaster dm = null!;

      TC.ID.Value = new TC.ThreadInfo(userName, callID);
      TC.RegisterCurrent();
      var me = TC.ID.Value!;
      me.ArmTimeoutMinutes((int)H.GetNProperty("defaultProcessTimeout", 30)!, reason: "Default watchdog");

      try
      {
        if (H.GetBProperty("autorun"))
          Thread.Sleep(2000);

        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"Procesando archivo: {filePath}");


        ThreadPool.GetAvailableThreads(out int workerThreads, out int completionPortThreads);
        ThreadPool.GetMaxThreads(out int maxWorkerThreads, out int maxCompletionPortThreads);

        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"Hilos disponibles: {workerThreads} de {maxWorkerThreads}");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"Hilos IO disponibles: {completionPortThreads} de {maxCompletionPortThreads}");


        List<ToolData> targets = Calculator.GetDeliveryDocs(xmlCall);
        dm = CreateDataMaster(xmlCall, targets);
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"Creada DataMaster con ID de instancia: {dm.InstanceId} para '{dm.GetValueString("opportunityFolder")}'");

        Calculator calculator = new(dm, targets);
        string project = dm.GetValueString("opportunityFolder");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"-- **** PROJECT: {project}  **** --");


        StoreCallFile(H.GetBProperty("storeXmlCall"), filePath, Path.GetDirectoryName(dm.FileName)!);

        calculator.RunCalculations();

        // after success ending the calculations will record the project in the DB
        _ = DBtools.InsertNewProjectWithBid(dm);

        if (xmlCall.SelectSingleNode("/request/requestInfo")!.Attributes!["type"]?.Value == "create"
            && H.GetBProperty("createOppsFoldersStructure"))
          createOppsFoldersStructure(dm.GetInnerText("dm/projectData/opportunityID"),
                                     dm.GetInnerText("dm/utils/utilsData/opportunityFolder"));

        ReturnRemoveFiles(dm);
        DBtools.UpdateCallRegistry(callID, "DONE", "OK");

        string a = $"-- **** PROJECT: {project} DONE  **** --";
        string b = "--" + string.Concat(Enumerable.Repeat("*", a.Length - 4)) + "--";
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", b);
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", a);
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", b);

        List<string> emailRecipients = new();
        if (H.GetBProperty("mailKAM")) emailRecipients.Add(dm.GetValueString("kam"));
        if (H.GetBProperty("mailCreatedBy")) emailRecipients.Add(dm.GetValueString("createdBy"));

        string bodyHtml = $@"
          <p>Enviado desde SmartBid</p>
          <p>Le informamos de que la generación de los documentos técnicos relacionados con el proyecto <b>{dm.GetValueString("opportunityFolder")}</b> ha terminado con éxito.</p>
          <p>Podrá encontrar los documentos en la carpeta <b>2.ING&#8203;/{dm.SBidRevision}/OUTPUT</b> de la oportunidad (los documentos podrán tardar unos minutos en actualizarse en SharePoint).</p>
          <table style='font-family:Courier New, monospace; border-collapse:collapse;'>
          <tr><td style='vertical-align: top; font-weight: bold;'>Created by:</td><td>{TC.ID.Value!.User}</td></tr>
          <tr><td style='vertical-align: top; font-weight: bold;'>KAM:</td><td>{dm.DM.SelectSingleNode("//projectData/kam")!.InnerText}</td></tr>
          <tr><td style='vertical-align: top; font-weight: bold;'>Call:</td><td>{Path.GetFileName(filePath)}</td></tr>
          <tr style='border-top: 1px solid black; padding-top: 12px'><td style='vertical-align: top; font-weight: bold;'>Docs Generated:</td>
          <td colspan='2'>{string.Join("<br/>", targets.Select(x => " " + x.Code))}</td></tr>
          </table>";

        _ = H.MailTo(recipients: emailRecipients,
                     subject: $@"SmartBid: Opportunity:{dm.GetValueString("opportunityFolder")} revision: {dm.SBidRevision} DONE",
                     body: bodyHtml,
                     isHtml: true);
      }
      catch (Exception ex)
      {
        string a = $"--❌❌ Error al procesar {dm.GetValueString("opportunityFolder")}❌❌";
        string b = "--" + string.Concat(Enumerable.Repeat("❌", a.Length / 2 - 4)) + "--\n";
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", b);
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", a);
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", b);
        H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌ 🧨  Excepción: {ex.GetType().Name} ");
        H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌ 📄    Mensaje: {ex.Message} ");
        H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", $"--❌❌ 🧭 StackTrace:\n{ex.StackTrace} ");
        H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ProcessFile", a);

        string bodyHtml = $@"
<p>Enviado desde SmartBid</p>
<p>El proceso de generación de la opp: <b>{dm.GetValueString("opportunityFolder")}</b> ha terminado con ERROR.</p>
<table style='font-family:Courier New, monospace; border-collapse:collapse;'>
<tr><td style='vertical-align: top; font-weight: bold;'>Error details:</td><td></td></tr>
<tr><td style='vertical-align: top; font-weight: bold;'>Type:</td><td>{ex.GetType().Name}</td></tr>
<tr><td style='vertical-align: top; font-weight: bold;'>Message:</td><td>{ex.Message}</td></tr>
<tr style='border-top: 1px solid black; padding-top: 12px'><td style='vertical-align: top; font-weight: bold;'>StackTrace:</td><td colspan='2'>{ex.StackTrace}</td></tr>
<tr><td style='vertical-align: top; font-weight: bold;'>User:</td><td>{TC.ID.Value!.User}</td></tr>
<tr><td style='vertical-align: top; font-weight: bold;'>Call:</td><td>{Path.GetFileName(filePath)}</td></tr>
<tr><td style='vertical-align: top; font-weight: bold;'>DataMaster:</td><td>{dm.FileName}</td></tr>
</table>";

        _ = H.MailTo(H.GetSProperty("EngineeringEmail").Split(';').ToList(),
                     subject: $"SmartBid Error processing {dm.GetValueString("opportunityFolder")}",
                     body: bodyHtml,
                     isHtml: true);


        DBtools.UpdateCallRegistry(callID, "FAILED", ex.Message);
      }
      finally
      {
        TC.ID.Value?.DisarmTimeout();
        TC.UnregisterCurrent();
        TC.ID.Value?.Dispose();
      }
    }

    // ... (el resto de la clase no necesita cambios)

    private static DataMaster CreateDataMaster(XmlDocument xmlCall, List<ToolData> targets)
    {
      DataMaster dm = new(xmlCall, targets);

      string projectFolder = Path.Combine(H.GetSProperty("processPath"),
                          dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");
      if (!Directory.Exists(projectFolder)) _ = Directory.CreateDirectory(projectFolder);

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
          string finalTargetDir = Path.Combine(targetDir, oppFolder);
          _ = Directory.CreateDirectory(finalTargetDir);
          File.Move(callFile, Path.Combine(finalTargetDir, fileName));
          H.PrintLog(4, TC.ID.Value!.Time(), TC.ID.Value!.User, "StoreCallFile", $"Call File '{Path.GetFileName(callFile)}' moved to '{finalTargetDir}'.");
        }
        catch (Exception ex)
        {
          H.PrintLog(6, TC.ID.Value!.Time(), TC.ID.Value!.User, $"❌❌ Error ❌❌ - StoreCallFile", $"❌Error❌ al mover '{callFile}': {ex.Message}");
        }
      }
      else
      {
        File.Delete(callFile);
      }
    }

    private static void ReturnRemoveFiles(DataMaster dm)
    {
      string opportunityFolder = dm.GetInnerText(@"dm/utils/utilsData/opportunityFolder");

      string processedPath = Path.Combine(H.GetSProperty("processPath"), opportunityFolder, dm.SBidRevision);
      string processedToolsPath = Path.Combine(processedPath, "TOOLS");
      string processedOutputsPath = Path.Combine(processedPath, "OUTPUT");
      string processedDocsPath = Path.Combine(processedPath, "DOCS");

      string returnPath = Path.Combine(H.GetSProperty("oppsPath"),
                                       $"OFERTAS {dm.GetInnerText("dm/projectData/opportunityID").Substring(0, 4)}",
                                       opportunityFolder,
                                       @"2.ING",
                                       dm.SBidRevision);

      string returnToolsPath = Path.Combine(returnPath, "TOOLS");
      string returnOutputsPath = Path.Combine(returnPath, "OUTPUT");
      string returnDocsPath = Path.Combine(returnPath, "DOCS");

      if (H.GetBProperty("returnTools"))
      {
        if (Directory.Exists(processedToolsPath))
        {
          foreach (string file in Directory.GetFiles(processedToolsPath))
          {
            _ = Directory.CreateDirectory(returnToolsPath);
            File.Copy(file, Path.Combine(returnToolsPath, Path.GetFileName(file)), overwrite: true);
          }
        }
      }

      if (!H.GetBProperty("storeTools"))
      {
        if (Directory.Exists(processedToolsPath))
        {
          foreach (string file in Directory.GetFiles(processedToolsPath))
            File.Delete(file);
        }
      }

      if (H.GetBProperty("returnDeliveries"))
      {
        if (Directory.Exists(processedOutputsPath))
        {
          foreach (string file in Directory.GetFiles(processedOutputsPath))
          {
            _ = Directory.CreateDirectory(returnOutputsPath);
            File.Copy(file, Path.Combine(returnOutputsPath, Path.GetFileName(file)), overwrite: true);
          }
        }
      }

      if (!H.GetBProperty("storeDeliveries"))
      {
        if (Directory.Exists(processedOutputsPath))
        {
          foreach (string file in Directory.GetFiles(processedOutputsPath))
            File.Delete(file);
        }
      }

      if (H.GetBProperty("createInputDocsShortcut"))
      {
        XmlNode? inputDocsXML;
        try
        {
          inputDocsXML = dm.DM.SelectSingleNode($"/dm/utils/{dm.SBidRevision}/inputDocs");
          if (inputDocsXML != null)
          {
            XmlDocument aaa = new();
            aaa.LoadXml("<root></root>");
            XmlNode importedNode = aaa.ImportNode(inputDocsXML, true);
            _ = aaa.DocumentElement!.AppendChild(importedNode);
            H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ReturnRemoveFiles", $"Creating inputDocs shortcut in {returnPath}", aaa);
          }
          else
          {
            H.PrintLog(5, TC.ID.Value!.Time(), TC.ID.Value!.User, "ReturnRemoveFiles", $"No inputDocs group found in DataMaster revision {dm.SBidRevision}");
          }
        }
        catch (Exception ex)
        {
          H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "ReturnRemoveFiles", $"Error creating inputDocs shortcut: {ex.Message}");
          inputDocsXML = null;
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
              using StreamWriter writer = new StreamWriter(shortcutPath);
              writer.WriteLine("[InternetShortcut]");
              writer.WriteLine($"URL={filePath}");
              writer.WriteLine("IconIndex=0");
              writer.WriteLine("IconFile=explorer.exe");
            }
          }
        }
      }

      if (H.GetBProperty("returnDocsShortcuts"))
      {
        if (Directory.Exists(processedDocsPath))
        {
          foreach (string file in Directory.GetFiles(processedDocsPath))
          {
            _ = Directory.CreateDirectory(returnDocsPath);
            File.Copy(file, Path.Combine(returnDocsPath, Path.GetFileName(file)), overwrite: true);
          }
        }
      }

      if (H.GetBProperty("returnDataMaster"))
      {
        _ = Directory.CreateDirectory(returnPath);
        File.Copy(dm.FileName, Path.Combine(returnPath, Path.GetFileName(dm.FileName)), overwrite: true);
      }
    }

    private static void createOppsFoldersStructure(string opportunityID, string oppFolder)
    {
      string templatePath = H.GetSProperty("oppsFoldersTemplate");
      string baseDestinationPath = H.GetSProperty("oppsPath");
      string destinationPath = Path.Combine(baseDestinationPath, $"OFERTAS {opportunityID.Substring(0, 4)}", oppFolder);

      if (!Directory.Exists(templatePath))
      {
        Console.WriteLine($"Template path does not exist: {templatePath}");
        H.PrintLog(2, TC.ID.Value!.Time(), TC.ID.Value!.User, "createOppsFoldersStructure", $"Template path does not exist: {templatePath}");
        return;
      }

      _ = Directory.CreateDirectory(destinationPath);
      CopyDirectory(templatePath, destinationPath);
    }

    private static void CopyDirectory(string sourceDir, string destinationDir)
    {
      foreach (string dirPath in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
      {
        string newDirPath = dirPath.Replace(sourceDir, destinationDir);
        _ = Directory.CreateDirectory(newDirPath);
      }

      foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
      {
        string newFilePath = filePath.Replace(sourceDir, destinationDir);
        File.Copy(filePath, newFilePath, overwrite: true);
      }
    }

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
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, $"⚠️ Path does not exist or is empty: '{path}'. Listener not started.");
          return;
        }

        using var fsw = new FileSystemWatcher
        {
          Path = path,
          Filter = "Call*.*",
          NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.Size,
          IncludeSubdirectories = false,
          EnableRaisingEvents = true
        };

        fsw.Created += (s, e) =>
        {
          _ = Task.Run(() =>
          {
            try
            {
              H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, $"Evento detectado: {e.FullPath}");
              if (WaitForFileReady(e.FullPath, attempts: 15, delay: TimeSpan.FromMilliseconds(500)))
              {
                onNewFile(e.FullPath);
              }
              else
              {
                H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, $"⚠️ File never stabilized: {e.FullPath}");
              }
            }
            catch (Exception ex)
            {
              H.PrintLog(6, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, $"❌ Error handling '{e.FullPath}': {ex.Message}");
            }
          }, token);
        };

        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, $"Started watching: {path}");

        using var done = new ManualResetEventSlim(false);
        using var reg = token.Register(() => done.Set());
        done.Wait();
        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", name, "Stopping...");
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
          if (size > 0 && size == lastSize)
          {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return true;
          }
          lastSize = size;
        }
        catch { }
        Thread.Sleep(delay);
      }
      return false;
    }

    private static void DoStuff1(string filePath)
    {
      void MoveDirectToCallsPath()
      {
        string callsPath = H.GetSProperty("callsPath");
        if (string.IsNullOrWhiteSpace(callsPath))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", "⚠️ 'callsPath' vacío en properties.xml");
          return;
        }
        _ = Directory.CreateDirectory(callsPath);
        string dest = Path.Combine(callsPath, Path.GetFileName(filePath));
        try { File.Move(filePath, dest); }
        catch (IOException) { File.Copy(filePath, dest, overwrite: false); File.Delete(filePath); }
        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"Archivo {Path.GetFileName(filePath)} movido a calls: {dest}");
      }

      if (!H.GetBProperty("processDWG"))
      {
        MoveDirectToCallsPath();
        return;
      }

      try
      {
        string lyotPath = string.Empty;
        try
        {
          var xmlProbe = new XmlDocument();
          using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
          { xmlProbe.Load(stream); }
          var lyotNodeProbe = xmlProbe.SelectSingleNode("//request/requestInfo/inputDocs/doc[@type='LYOT']") as XmlElement;
          lyotPath = lyotNodeProbe?.InnerText?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"⚠️ No se pudo leer el XML antes de mover: {ex.Message}");
          MoveDirectToCallsPath();
          return;
        }

        if (string.IsNullOrWhiteSpace(lyotPath) || !File.Exists(lyotPath))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1",
              $"⚠️ LYOT ausente o DWG inexistente. XML: {filePath} \n LYOT: '{lyotPath}'. Se mueve directo a callsPath.");
          MoveDirectToCallsPath();
          return;
        }

        string tempDir = H.GetSProperty("storageTemp");
        if (string.IsNullOrWhiteSpace(tempDir))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", "⚠️ 'storageTemp' vacío en properties.xml");
          return;
        }

        _ = Directory.CreateDirectory(tempDir);
        string movedPath = Path.Combine(tempDir, Path.GetFileName(filePath));
        try { File.Move(filePath, movedPath); }
        catch (IOException) { File.Copy(filePath, movedPath, overwrite: false); File.Delete(filePath); }
        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"Archivo movido a temp: {movedPath}");

        string dwgBaseName = Path.GetFileNameWithoutExtension(lyotPath);
        string currentXmlNameNoExt = Path.GetFileNameWithoutExtension(movedPath);
        if (!currentXmlNameNoExt.EndsWith($"_{dwgBaseName}", StringComparison.OrdinalIgnoreCase))
        {
          string renamedXml = Path.Combine(Path.GetDirectoryName(movedPath)!, $"{currentXmlNameNoExt}_{dwgBaseName}.xml");
          File.Move(movedPath, renamedXml);
          movedPath = renamedXml;
          H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"XML renombrado a: {movedPath}");
        }

        string recipient = H.GetSProperty("DWG_recipiant");
        if (string.IsNullOrWhiteSpace(recipient))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", "⚠️ 'DWG_recipiant' vacío en properties.xml; no se puede enviar correo.");
          return;
        }

        var recipients = new List<string> { recipient };
        string subject = "New DWG to prepare for SmartBid";
        string body = "here is a new dwg file to identify the layers. " +
                      $"whenever is done, place a copy at {H.GetSProperty("callsPathDWG")}";
        bool sent = H.MailTo(recipients, subject, body, lyotPath);
        if (sent)
          H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"DWG enviado a {recipient}: {Path.GetFileName(lyotPath)}");
        else
          H.PrintLog(6, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"❌ Error al enviar DWG a {recipient}: {Path.GetFileName(lyotPath)}");
      }
      catch (Exception ex)
      {
        H.PrintLog(6, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff1", $"❌ Excepción: {ex.GetType().Name} - {ex.Message}");
        try { MoveDirectToCallsPath(); } catch { }
      }
    }

    private static void DoStuff2(string returnedDwgPath)
    {
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
          File.Copy(src, dst, overwrite: true);
          File.Delete(src);
        }
      }

      try
      {
        if (string.IsNullOrWhiteSpace(returnedDwgPath) || !File.Exists(returnedDwgPath))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"⚠️ DWG no existe: {returnedDwgPath}");
          return;
        }

        string dwgBaseName = Path.GetFileNameWithoutExtension(returnedDwgPath);
        string tempDir = H.GetSProperty("storageTemp");
        string callsDir = H.GetSProperty("callsPath");

        if (string.IsNullOrWhiteSpace(tempDir) || string.IsNullOrWhiteSpace(callsDir))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", "⚠️ 'storageTemp' o 'callsPath' vacíos en properties.xml");
          return;
        }

        var candidates = Directory.GetFiles(tempDir, $"*_{dwgBaseName}.xml", SearchOption.TopDirectoryOnly);
        if (candidates.Length == 0)
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"⚠️ No se encontró XML en TEMP que termine en _{dwgBaseName}.xml");
          return;
        }

        string? matchedXml = null;
        string? originalDwgPath = null;

        foreach (var candidateXml in candidates)
        {
          try
          {
            var xml = new XmlDocument();
            using (var stream = new FileStream(candidateXml, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            { xml.Load(stream); }
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
          catch { }
        }

        if (matchedXml is null || string.IsNullOrWhiteSpace(originalDwgPath))
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"⚠️ No se encontró XML cuyo LYOT coincida con '{dwgBaseName}'.");
          _ = H.MailTo(H.GetSProperty("DWG_recipiant").Split(';').ToList(),
              subject: "Error: DWG name mismatch",
              body: $"The returned DWG file '{Path.GetFileName(returnedDwgPath)}' does not match any LYOT in the pending XML calls.\n" +
                    "Please ensure the DWG file name matches the original LYOT name exactly.");
          return;
        }

        if (File.Exists(originalDwgPath))
        {
          string dir = Path.GetDirectoryName(originalDwgPath)!;
          string name = Path.GetFileNameWithoutExtension(originalDwgPath);
          string ext = Path.GetExtension(originalDwgPath);
          string backup = Path.Combine(dir, $"{name}_original{ext}");
          backup = EnsureUniqueFileName(backup);
          MoveFileCrossVolume(originalDwgPath, backup, overwrite: false);
          H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"Original respaldado como: {backup}");
          Thread.Sleep(1000);
        }
        else
        {
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"⚠️ El DWG original indicado en LYOT no existe: {originalDwgPath}");
        }

        string originalDir = Path.GetDirectoryName(originalDwgPath)!;
        string originalName = Path.GetFileName(originalDwgPath);
        string destinationDwg = Path.Combine(originalDir, originalName);
        MoveFileCrossVolume(returnedDwgPath, destinationDwg, overwrite: true);
        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"DWG actualizado: {destinationDwg}");

        string xmlNameNoExt = Path.GetFileNameWithoutExtension(matchedXml);
        string suffix = "_" + dwgBaseName;
        string originalCallBase = xmlNameNoExt.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? xmlNameNoExt.Substring(0, xmlNameNoExt.Length - suffix.Length)
            : xmlNameNoExt;

        string finalCallName = originalCallBase + ".xml";
        string finalCallPath = Path.Combine(callsDir, finalCallName);

        if (File.Exists(finalCallPath))
        {
          string unique = Path.Combine(
              callsDir,
              $"{Path.GetFileNameWithoutExtension(finalCallName)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(finalCallName)}"
          );
          H.PrintLog(5, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"⚠️ Ya existe {finalCallName} en callsPath. Usando: {Path.GetFileName(unique)}");
          finalCallPath = unique;
        }

        MoveFileCrossVolume(matchedXml, finalCallPath, overwrite: false);
        H.PrintLog(4, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"XML movido a callsPath para disparar SmartBid: {finalCallPath}");
      }
      catch (Exception ex)
      {
        H.PrintLog(6, DateTime.Now.ToString("HH:mm:ss"), "SYSTEM", "DoStuff2", $"❌ Excepción: {ex.GetType().Name} - {ex.Message}");
      }
    }
  }

  static class TC
  {
    public class ThreadInfo : IDisposable
    {
      private readonly object _lock = new();

      public Stopwatch chrono;
      public int ThreadId { get; }
      public string User { get; }
      public int? CallId { get; }

      public CancellationTokenSource Cts { get; } = new();
      public CancellationToken Token => Cts.Token;

      private Timer? _watchdog;
      private TimeSpan _timeout = Timeout.InfiniteTimeSpan;

      private readonly ConcurrentBag<Process> _children = new();

      private sealed class Disposer : IDisposable
      {
        private readonly Action _dispose;
        private bool _done;
        public Disposer(Action dispose) => _dispose = dispose;
        public void Dispose()
        {
          if (_done) return;
          _done = true;
          _dispose();
        }
      }

      public ThreadInfo(string user, int? callId = null)
      {
        ThreadId = Environment.CurrentManagedThreadId;
        User = user;
        CallId = callId;
        chrono = new Stopwatch();
        chrono.Start();
      }

      public string Time()
      {
        TimeSpan elapsed = chrono.Elapsed;
        return $"{(int)elapsed.TotalMinutes:D2}:{elapsed.Seconds:D2}.{elapsed.Milliseconds:D3}";
      }

      public void ArmTimeoutMinutes(int minutes, string? reason = null, bool killChildren = true)
      {
        if (minutes <= 0)
          throw new ArgumentOutOfRangeException(nameof(minutes), "Minutes must be > 0 for ArmTimeoutMinutes.");

        lock (_lock)
        {
          _timeout = TimeSpan.FromMinutes(minutes);
          _watchdog?.Dispose();
          _watchdog = new Timer(_ =>
          {
            RequestCancel(reason ?? $"Watchdog timeout {minutes} min", killChildren);
          }, null, _timeout, Timeout.InfiniteTimeSpan);

          H.PrintLog(4, Time(), User, "Watchdog", $"Armed for {minutes} minute(s).");
        }
      }

      public void PetTimeout()
      {
        lock (_lock)
        {
          if (_watchdog is null || _timeout == Timeout.InfiniteTimeSpan) return;
          _watchdog.Change(_timeout, Timeout.InfiniteTimeSpan);
          H.PrintLog(5, Time(), User, "Watchdog", "Pet (timer reset).");
        }
      }

      public void DisarmTimeout()
      {
        lock (_lock)
        {
          _timeout = Timeout.InfiniteTimeSpan;
          _watchdog?.Dispose();
          _watchdog = null;
          H.PrintLog(4, Time(), User, "Watchdog", "Disarmed.");
        }
      }

      public IDisposable? ArmScopedTimeoutMinutes(int minutes, string? reason = null, bool killChildren = true)
      {
        if (minutes == 0)
        {
          H.PrintLog(3, Time(), User, "Watchdog", "ArmScopedTimeoutMinutes minutes==0 → no-op.");
          return null;
        }

        if (minutes < 0)
          throw new ArgumentOutOfRangeException(nameof(minutes), "Minutes must be >= 0 (0 means no-op).");

        lock (_lock)
        {
          var prevTimeout = _timeout;
          var prevWatchdog = _watchdog;

          var scoped = TimeSpan.FromMinutes(minutes);
          var localTimer = new Timer(_ =>
          {
            RequestCancel(reason ?? $"Scoped watchdog timeout {minutes} min", killChildren);
          }, null, scoped, Timeout.InfiniteTimeSpan);

          _timeout = scoped;
          _watchdog = localTimer;

          H.PrintLog(3, Time(), User, "Watchdog", $"Scoped armed for {minutes} minute(s).");

          return new Disposer(() =>
          {
            lock (_lock)
            {
              try { _watchdog?.Dispose(); } catch { }

              _timeout = prevTimeout;
              _watchdog = prevWatchdog;

              if (_watchdog != null && _timeout != Timeout.InfiniteTimeSpan)
              {
                try { _watchdog.Change(_timeout, Timeout.InfiniteTimeSpan); } catch { }
                H.PrintLog(3, Time(), User, "Watchdog", "Scoped disposed → previous watchdog restored.");
              }
              else
              {
                H.PrintLog(3, Time(), User, "Watchdog", "Scoped disposed → watchdog disarmed.");
              }
            }
          });
        }
      }

      public void RequestCancel(string? reason = null, bool killChildren = true)
      {
        if (!Cts.IsCancellationRequested)
        {
          H.PrintLog(2, Time(), User, "Cancel", $"Cancellation requested. Reason: {reason ?? "(none)"}");
          try { Cts.Cancel(); } catch { }

          if (killChildren)
            KillChildren();
        }
      }

      public void RegisterProcess(Process p)
      {
        if (p is null) return;
        _children.Add(p);

        if (Cts.IsCancellationRequested) SafeKill(p);

        try
        {
          p.EnableRaisingEvents = true;
          p.Exited += (_, __) =>
          {
            try
            {
              H.PrintLog(2, Time(), User, "RegisterProcess", $"Exited: {p.StartInfo?.FileName} {p.StartInfo?.Arguments}");
            }
            catch { }
          };
        }
        catch { }
      }

      private void KillChildren()
      {
        foreach (var p in _children)
          SafeKill(p);
      }

      private static void SafeKill(Process p)
      {
        try
        {
          if (!p.HasExited)
          {
            try { p.Kill(entireProcessTree: true); }
            catch { p.Kill(); }
            try { p.WaitForExit(5000); } catch { }
          }
        }
        catch { }
      }

      public void Dispose()
      {
        lock (_lock)
        {
          try { _watchdog?.Dispose(); } catch { }
          _watchdog = null;
          _timeout = Timeout.InfiniteTimeSpan;
        }
        try { Cts.Dispose(); } catch { }
      }
    }

    public static AsyncLocal<ThreadInfo> ID = new();

    private static readonly ConcurrentDictionary<int, ThreadInfo> _byCallId = new();

    public static void RegisterCurrent()
    {
      var info = ID.Value;
      if (info?.CallId is int callId)
        _byCallId[callId] = info;
    }

    public static void UnregisterCurrent()
    {
      var info = ID.Value;
      if (info?.CallId is int callId)
        _byCallId.TryRemove(callId, out _);
    }

    public static bool CancelCall(int callId, string? reason = null, bool killChildren = true)
    {
      if (_byCallId.TryGetValue(callId, out var info))
      {
        info.RequestCancel(reason ?? "External cancel", killChildren);
        return true;
      }
      return false;
    }

    public static bool ArmTimeoutFor(int callId, int minutes, string? reason = null, bool killChildren = true)
    {
      if (_byCallId.TryGetValue(callId, out var info))
      {
        info.ArmTimeoutMinutes(minutes, reason, killChildren);
        return true;
      }
      return false;
    }

    public static bool PetTimeoutFor(int callId)
    {
      if (_byCallId.TryGetValue(callId, out var info))
      {
        info.PetTimeout();
        return true;
      }
      return false;
    }

    public static bool DisarmTimeoutFor(int callId)
    {
      if (_byCallId.TryGetValue(callId, out var info))
      {
        info.DisarmTimeout();
        return true;
      }
      return false;
    }
  }
}