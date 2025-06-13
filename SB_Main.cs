using System.Collections.Concurrent;
using System.Media;
using System.Text.RegularExpressions;
using System.Xml;
using Windows.Services.Store;
using System.Threading;

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


            Console.Beep(1000, 400);


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
                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "myEvent", "Salida solicitada... deteniendo el watcher.");
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



            string inputFilesFolder = xmlCall.SelectSingleNode(@"request/requestInfo/inputFolder")?.InnerText ?? "";
            inputFilesFolder = Path.Combine(H.GetSProperty("callsPath"), inputFilesFolder);

            int callID = DBtools.InsertCallStart(xmlCall); // Report starting process to DB

            DataMaster dm = CreateDataMaster(xmlCall); //Create New DataMaster

            StoreInputFiles(dm, 1, inputFilesFolder, filePath); // Moves call and input files to final storage folder

            Calculator calculator = new Calculator(dm, xmlCall);

            calculator.RunCalculations(xmlCall);

            DBtools.UpdateCallRegistry(callID, "DONE", "OK");

            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--***************************************--");
            H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--****||PROJECT: {dm.GetInnerText(@"dm/utils/utilsData/projectFolder")} DONE||****--");
            H.PrintLog(2, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"--***************************************--");
            
            H.DeleteBookmarkText("ES_Informe de corrosión_Rev0.0.docx", "Ruta_05", dm,"OUTPUT");
            H.EnviarMail(dm);

            //por el momento borramos el fichero de entrada.... luego lo guardaremos en función del nivel de log que tengamos.
        }


        static DataMaster CreateDataMaster(XmlDocument xmlCall) //Creamos el datamaster
        {
            H.PrintXML(xmlCall); //Print the XML call for debugging

            //Instantiating the DataMaster class with the XML string 
            DataMaster dm = new DataMaster(xmlCall);

            //Creating the projectFolder in the storage directory

            string projectFolder = Path.Combine(H.GetSProperty("storagePath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/projectFolder")?.InnerText ?? "");

            if (!Directory.Exists(projectFolder))
                _ = Directory.CreateDirectory(projectFolder);

            dm.SaveDataMaster();

            return dm;
        }

        static void StoreInputFiles (DataMaster dm, int rev, string inputFilesFolder, string callPath = "")

        {
            string targetDir = Path.Combine(H.GetSProperty("storagePath"), dm.DM.SelectSingleNode(@"dm/utils/utilsData/projectFolder")?.InnerText ?? "", "DOC");


            if (!Directory.Exists(targetDir))
                _ = Directory.CreateDirectory(targetDir);

            //AÑADIR LA FUNCIONALIDAD DE COMPROBAR QUE TODOS LOS FICHEROS ANUNCIADOS EN CALL SON CORRECTAMENTE MOVIDOS AL DIRECTORIO DE DESTINO



            foreach (string filePath in Directory.GetFiles(inputFilesFolder))
            {
                string fileName = Path.GetFileName(filePath);
                string destinationPath = Path.Combine(targetDir, fileName);

                try
                {
                    File.Move(filePath, destinationPath);
                    H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"Archivo '{fileName}' movido a '{targetDir}'.");
                }
                catch (Exception ex)
                {
                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error", $"Error al mover '{fileName}': {ex.Message}");
                }
            }
            H.PrintLog(3, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"All input files moved to '{targetDir}'.");

            if (!string.IsNullOrEmpty(callPath)) 
            {
                targetDir = Path.GetDirectoryName(targetDir) ?? targetDir; // Ensure targetDir is not null

                try
                {
                    File.Move(callPath, Path.Combine(targetDir, Path.GetFileName(callPath)));
                    H.PrintLog(4, ThreadContext.CurrentThreadInfo.Value.User, "ProcessFile", $"Archivo '{callPath}' movido a '{targetDir}'.");
                }
                catch (Exception ex)
                {
                    H.PrintLog(5, ThreadContext.CurrentThreadInfo.Value.User, "Error", $"Error al mover '{callPath}': {ex.Message}");
                }
            }

            Directory.Delete(inputFilesFolder, true);

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
