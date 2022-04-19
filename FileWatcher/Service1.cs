using System;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Configuration;
using System.IO.Compression;
using ClosedXML.Excel;
using System.Text;

namespace FileWatcher
{
    public partial class Service1 : ServiceBase
    {
        Logger logger;
        public Service1()
        {
            InitializeComponent();
            this.CanStop = true;
            this.CanPauseAndContinue = true;
            this.AutoLog = true;
        }

        protected override void OnStart(string[] args)
        {
            logger = new Logger();
            Thread loggerThread = new Thread(new ThreadStart(logger.Start));
            loggerThread.Start();
        }

        protected override void OnStop()
        {
            logger.Stop();
            Thread.Sleep(1000);
        }
    }

    class Logger
    {
        FileSystemWatcher watcher;
        object obj = new object();
        bool enabled = true;
        readonly string logfileloc = ConfigurationManager.AppSettings["logfilelocation"];
        readonly string namedirectoryxlsx = ConfigurationManager.AppSettings["namedirectoryxlsx"];
        public Logger()
        {
            watcher = new FileSystemWatcher(ConfigurationManager.AppSettings["pathfilelocation"]);
            watcher.Created += Watcher_Created;
        }

        public void Start()
        {
            RecordEntry("Служба запущена", "");
            watcher.EnableRaisingEvents = true;
            while (enabled)
            {
                Thread.Sleep(1000);
            }
        }
        public void Stop()
        {
            RecordEntry("Служба Остановлена", "");
            watcher.EnableRaisingEvents = false;
            enabled = false;
        }
    
        // создание файлов
        private void Watcher_Created(object sender, FileSystemEventArgs e)
        {
            try
            {
                if (e.Name.Split('.')[1] == "zip")
                {
                    string fileEvent = "Принят";
                    string filePath = e.FullPath;
                    string workfilePathxlsx = filePath.Split('\\')[0] + "\\" + namedirectoryxlsx + "\\" + e.Name + "\\";
                    while (true)
                    {
                        try
                        {
                            ZipArchive zipArchive = ZipFile.OpenRead(filePath);
                            foreach (ZipArchiveEntry entry in zipArchive.Entries)
                            {
                                if (entry.Name.Split('.')[1] == "xlsx")
                                {
                                    DirectoryInfo dirInfo = new DirectoryInfo(workfilePathxlsx);
                                    if (!dirInfo.Exists)
                                    {
                                        dirInfo.Create();
                                    }
                                    entry.ExtractToFile(workfilePathxlsx + entry.Name, true);

                                    var workbook = new XLWorkbook(workfilePathxlsx + entry.Name);
                                    var ws1 = workbook.Worksheet(1);
                                    var file = File.Create(workfilePathxlsx + entry.Name.Split('.')[0]+".txt");
                                    foreach (var Row in ws1.Rows())
                                    {
                                        foreach (var Cell in Row.Cells())
                                        {
                                            byte[] buffer = Encoding.Default.GetBytes(Cell.Address.ColumnLetter + Row.RowNumber().ToString() + " = " + Cell.Value.ToString() + "\n");
                                            file.Write(buffer, 0, buffer.Length);
                                        }
                                    }
                                    file.Close();

                                }
                            }
                            zipArchive.Dispose();
                            break;
                        }
                        catch (IOException)
                        {

                            Thread.Sleep(500);
                        }
                    }

                    RecordEntry(fileEvent, filePath);
                    
                    RecordEntry("Обработан", filePath);
                }
            }
            catch
            {
                RecordEntry("Ошибка", "");
            }
        }
     

        private void RecordEntry(string fileEvent, string filePath)
        {
            lock (obj)
            {
                using (StreamWriter writer = new StreamWriter(logfileloc, true))
                {
                    if (filePath == "")
                    {
                        writer.WriteLine(String.Format("{0} {1}  {2}",
                                                   DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), filePath, fileEvent));
                    }
                    else
                    {
                        writer.WriteLine(String.Format("{0} файл {1} был {2}",
                            DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), filePath, fileEvent));
                    }
                    writer.Flush();
                }
               
            }

        }
    }
}



