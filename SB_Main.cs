using System.Collections.Concurrent;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices.Marshalling;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Windows.Services.Store;
using Windows.UI.ViewManagement;

namespace SmartBid
{
    class SB_Main
    {
        private static ConcurrentQueue<string> _fileQueue = new ConcurrentQueue<string>();
        private static AutoResetEvent _eventSignal = new AutoResetEvent(false);
        private static FileSystemWatcher? watcher;
        private static bool _stopRequested = false;

        static void Main()
        {
            string path = H.GetSProperty("callsPath");

            watcher = new FileSystemWatcher
            {
                Path = path,
                Filter = "*.*",
                NotifyFilter = NotifyFilters.FileName
            };

            watcher.Created += (sender, e) =>
            {
                H.PrintLog(5, "Main", "myEvent", $"Evento detectado: {e.FullPath}");

                if (Regex.IsMatch(Path.GetFileName(e.FullPath), @"^call_\d+\.xml$", RegexOptions.IgnoreCase))
                {
                    _fileQueue.Enqueue(e.FullPath);
                    _ = _eventSignal.Set();
                }
            };

            watcher.EnableRaisingEvents = true;
            H.PrintLog(5, "Main", "myEvent", $"Observando el directorio: {path}");
            H.PrintLog(5, "Main", "myEvent", "Presiona 'Q' para salir...");

            // Procesamiento en un hilo separado
            _ = Task.Run(ProcessFiles);

            // Monitor de entrada para salir con 'Q'
            while (true)
            {
                if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
                {
                    H.PrintLog(5, "Main", "myEvent", "Salida solicitada... deteniendo el watcher.");
                    watcher.EnableRaisingEvents = false; // Detiene la detección de archivos nuevos
                    _stopRequested = true;
                    break;
                }
                Thread.Sleep(1000); // Reduce la carga de la CPU
            }

            H.PrintLog(5, "Main", "myEvent", "Todos los archivos han sido procesados. Programa terminado.");
        }

        static void ProcessFiles()
        {
            while (!_stopRequested || !_fileQueue.IsEmpty) // Sigue procesando hasta vaciar la cola
            {
                _ = _eventSignal.WaitOne();

                while (_fileQueue.TryDequeue(out string filePath))
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

            XmlDocument xmlCall = new XmlDocument();
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xmlCall.Load(stream);
            }

            // ✅ Inicializar el contexto lógico (seguro para ejecución en paralelo)
            string userName = xmlCall.SelectSingleNode(@"request/requestInfo/createdBy")?.InnerText ?? "UnknownUser";
            ThreadContext.CurrentThreadInfo.Value = new ThreadContext.ThreadInfo(userName);

            H.PrintLog(5, userName, "ProcessFile", $"Procesando archivo: {filePath}");

            int callID = DBtools.InsertCallStart(xmlCall); // Report starting process to DB

            DataMaster dm = CreateDataMaster(xmlCall); //Create New DataMaster

            ProcessInputFiles(dm, 1); // checks that all files declare exits and stores the checksum of the file for comparison

            if (H.GetBProperty("storeXmlCall")) //Stores de call file in case configuration says so
                StoreCallFile(filePath);

            Calculator calculator = new Calculator(dm, xmlCall);

            calculator.RunCalculations(xmlCall);

            ReturnRemoveFiles(dm); // Returns or removes files depending on configuration

            DBtools.UpdateCallRegistry(callID, "DONE", "OK");

            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--***************************************--");
            H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--****||PROJECT: {dm.GetValueString("opportunityFolder")} DONE||****--");
            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--***************************************--");

            // H.DeleteBookmarkText("ES_Informe de corrosión_Rev0.0.docx", "Ruta_05", dm, "OUTPUT");


            List<string> emailRecipients = new List<string>();

            // Add KAM email if configured to do so
            if (H.GetBProperty("mailKAM"))
                emailRecipients.Add(dm.GetValueString("kam"));

            // Add CreatedBy email if configured to do so
            if (H.GetBProperty("mailCreatedBy"))
                emailRecipients.Add(dm.GetValueString("createdBy"));

            H.MailTo(emailRecipients, "Mail de Prueba", "Enviado desde SmartBid");



        }

        private static DataMaster CreateDataMaster(XmlDocument xmlCall) //Creamos el datamaster
        {
            H.PrintXML(xmlCall); //Print the XML call for debugging

            //Instantiating the DataMaster class with the XML string 
            DataMaster dm = new DataMaster(xmlCall);

            //Creating the projectFolder in the storage directory
            string projectFolder = Path.Combine(H.GetSProperty("processPath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");

            if (!Directory.Exists(projectFolder))
                _ = Directory.CreateDirectory(projectFolder);

            dm.SaveDataMaster();

            return dm;
        }
        private static void ProcessInputFiles(DataMaster dm, int rev)
        {
            string inputPath = Path.Combine(H.GetSProperty("oppsPath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "");

            foreach (XmlElement doc in dm.DM.SelectNodes(@$"dm/utils/rev_{rev.ToString("D2")}/inputDocs/doc"))
            {
                string fileName = Path.Combine(inputPath, "1.DOC", doc.GetAttribute("type"), doc.InnerText);

                if (!File.Exists(fileName))
                {
                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error - ProcessFile", $"⚠️ El archivo '{fileName}' no existe.");
                    continue; // Saltar este documento y seguir con los demás
                }
                string hash = CalcularMD5(fileName); // Calculate MD5 hash for the file
                string lastModified = File.GetLastWriteTime(fileName).ToString("yyyy-MM-dd HH:mm:ss");

                doc.SetAttribute("hash", hash); // Set the hash attribute in the XML
                doc.SetAttribute("lastModified", lastModified); // Set the hash attribute in the XML

                DBtools.InsertFileHash(fileName, doc.GetAttribute("type"), hash, lastModified); // Store the file hash in the database

                H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"Archivo '{fileName}' registered");

            }
            H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"All input files have been registered'.");


        }
        private static void StoreCallFile(string callFile)
        {
            if (!string.IsNullOrEmpty(callFile))
            {
                try
                {
                    string fileName = $"{DateTime.Now:yyMMdd-HHmmss}_{Path.GetFileName(callFile)}";
                    string targetDir = Path.Combine(H.GetSProperty("oppsPath"), "calls");

                    File.Move(callFile, Path.Combine(targetDir, fileName));

                    H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "StoreCallFile", $"Archivo '{callFile}' movido a '{targetDir}'.");
                }
                catch (Exception ex)
                {
                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error - StoreCallFile", $"Error al mover '{callFile}': {ex.Message}");
                }
            }
        }
        private static string CalcularMD5(string path)
        {
            using var stream = File.OpenRead(path);
            using var md5 = MD5.Create();
            byte[] hash = md5.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
        private static void ReturnRemoveFiles(DataMaster dm)
        {
            string projectFolder = dm.DM.SelectSingleNode(@"dm/utils/utilsData/opportunityFolder")?.InnerText ?? "";
            string processedToolsPath = Path.Combine(H.GetSProperty("processPath"), projectFolder, "TOOLS");
            string processedDeliveriesPath = Path.Combine(H.GetSProperty("processPath"), projectFolder, "OUTPUT");
            string oppsToolsPath = Path.Combine(H.GetSProperty("oppsPath"), projectFolder, @"2.ING\OBS");
            string oppsDeliveriesPath = Path.Combine(H.GetSProperty("oppsPath"), projectFolder, @"2.ING\OBS");


            if (H.GetBProperty("returnTools"))
                foreach (string file in Directory.GetFiles(processedToolsPath))
                {
                    Directory.CreateDirectory(oppsToolsPath); // Crea si no existe
                    File.Copy(file, Path.Combine(oppsToolsPath, Path.GetFileName(file)), overwrite: true);
                }

            if (H.GetBProperty("removeTools"))
                foreach (string file in Directory.GetFiles(processedToolsPath))
                {
                    File.Delete(file);
                }


            if (H.GetBProperty("returnDeliveries"))
                foreach (string file in Directory.GetFiles(processedDeliveriesPath))
                {
                    Directory.CreateDirectory(oppsDeliveriesPath); // Crea si no existe
                    File.Copy(file, Path.Combine(oppsDeliveriesPath, Path.GetFileName(file)), overwrite: true);
                }

            if (H.GetBProperty("removeDeliveries"))
                foreach (string file in Directory.GetFiles(processedDeliveriesPath))
                {
                    File.Delete(file);
                }

            
            if (H.GetBProperty("returnDataMaster"))
                File.Copy(dm.FileName, Path.Combine(H.GetSProperty("oppsPath"), projectFolder), overwrite: true);


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
                ThreadId = Thread.CurrentThread.ManagedThreadId;
                User = user;
            }
        }

        // ✅ Ahora usamos AsyncLocal en lugar de ThreadLocal
        public static AsyncLocal<ThreadInfo> CurrentThreadInfo = new AsyncLocal<ThreadInfo>();
    }


}
