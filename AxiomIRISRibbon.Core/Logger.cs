using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace AxiomIRISRibbon.Core
{
    public class Logger
    {

        private static readonly string STARTUP_MESSAGE = "************************** Application Starting Up **************************";
        private static readonly string SHUTDWN_MESSAGE = "************************* Application Shutting Down *************************";

        private static readonly string LOG_FILE_NAME_PFX = Constants.LOG_PATH + "\\AxiomIRISRibbon";
        private static readonly string LOG_FILE_NAME = LOG_FILE_NAME_PFX + ".log";
        private static readonly int MAX_CHARS = 2 * 1024 * 1024; //[2M characters]



        private static object _lock = new object();

        private static StreamWriter _writer;

        private static int _chars = 0;




        private static Queue<LogEntry> _queue = new Queue<LogEntry>();
        private static Thread _worker;



        public static void Init()
        {
            if (!Directory.Exists(Constants.LOG_PATH))
            {
                Directory.CreateDirectory(Constants.LOG_PATH);
            }
            FileInfo info = new FileInfo(LOG_FILE_NAME);
            if (info.Exists) _chars = (int)info.Length;
            else _chars = 0;
            _writer = new StreamWriter(LOG_FILE_NAME, true);
            _worker = new Thread(write);
            _worker.Name = "LogWriter";
            _worker.Start();
            Log(STARTUP_MESSAGE);
        }



        public static void Log(string msg)
        {
            lock (_lock)
            {
                if (_writer == null) throw new Exception("Attempt to use Logging before initialization");

                _queue.Enqueue(new LogEntry(msg));
                Monitor.PulseAll(_lock);
            }
        }


        public static void Log(string msg, Exception e)
        {
            lock (_lock)
            {
                if (_writer == null) throw new Exception("Attempt to use Logging before initialization");

                _queue.Enqueue(new LogEntry(msg, e));
                Monitor.PulseAll(_lock);
            }
        }

        public static void Log(Exception e, string msg)
        {
            lock (_lock)
            {
                if (_writer == null) throw new Exception("Attempt to use Logging before initialization");

                _queue.Enqueue(new LogEntry(msg, e));
                Monitor.PulseAll(_lock);
            }
        }

        public static void Close()
        {
            lock (_lock)
            {
                if (_writer != null)
                {
                    Log(SHUTDWN_MESSAGE);
                }
            }
        }



        private static void write()
        {
            lock (_lock)
            {
                while (true)
                {
                    while (_queue.Count == 0) Monitor.Wait(_lock);
                    LogEntry entry = _queue.Dequeue();
                    if (entry.Exception == null)
                    {
                        switchLog();
                        string line = "[" + System.DateTime.Now + "] " + entry.Message;
                        _writer.WriteLine(line);
                        _writer.Flush();
                        _chars += line.Length;
                        if (entry.Message == SHUTDWN_MESSAGE)
                        {
                            _writer.Close();
                            return;
                        }
                    }
                    else
                    {
                        switchLog();
                        string line = "[" + System.DateTime.Now + "] " + entry.Message + "\r\n";
                        _writer.WriteLine(line);
                        _chars += line.Length;
                        Exception e = entry.Exception;
                        while (e != null)
                        {
                            line = "[" + System.DateTime.Now + "] " + e.Message;
                            _writer.WriteLine(line);
                            _chars += line.Length;
                            string lines = e.StackTrace;
                            if (lines != null)
                            {
                                _writer.WriteLine(lines);
                                _chars += lines.Length;
                            }
                            e = e.InnerException;
                        }
                        _writer.Flush();
                    }
                }
            }
        }


        // Expected to be called after acquiring _lock
        private static void switchLog()
        {
            Exception ex = null;
            if (_chars < MAX_CHARS) return;

            DateTime d = DateTime.Now;
            try
            {
                _writer.Flush();
                if (_writer != null) _writer.Close();

                string destFileName = LOG_FILE_NAME_PFX + "_" + d.Year + "_" + pad2(d.Month) + "_" + pad2(d.Day) + "_" + pad2(d.Hour) +
                                      "_" + pad2(d.Minute) + "_" + pad2(d.Second) + "_" + pad3(d.Millisecond) + ".log";
                if (!File.Exists(destFileName)) File.Move(LOG_FILE_NAME, destFileName);
                _writer = new StreamWriter(LOG_FILE_NAME, true);
                _chars = 0;
            }
            catch (Exception e)
            {
                ex = e;
            }
            if (ex != null)
            {
                Log("SwitchLog Error", ex);
            }
        }


        private static string pad2(int i)
        {
            if (i < 10) return "0" + i;
            return "" + i;
        }


        private static string pad3(int i)
        {
            if (i < 10) return "00" + i;
            if (i < 100) return "0" + i;
            return "" + i;
        }


        private class LogEntry
        {
            public string Message { get; private set; }
            public Exception Exception { get; private set; }


            public LogEntry(string m)
            {
                Message = m;
            }

            public LogEntry(string m, Exception e)
            {
                Message = m;
                Exception = e;
            }
        }

    }
}
