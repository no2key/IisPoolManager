using System;
using System.Diagnostics;
using System.DirectoryServices;
using System.Runtime.InteropServices;

namespace IisPoolManager
{
    internal sealed class IisPoolWorker
    {
        private readonly string Host;
        private readonly int Interval;
        private readonly string Password;
        private readonly string PoolName;
        private readonly string User;
        private readonly DirectoryEntry _directoryEntry;
        private string applicationPoolUrl;

        public IisPoolWorker()
        {
            Host = IISPoolManager.Default.host;
            User = IISPoolManager.Default.user;
            Password = IISPoolManager.Default.password;
            PoolName = IISPoolManager.Default.pool;
            Interval = IISPoolManager.Default.interval;
            IISPoolManager.Default.Save();
            applicationPoolUrl = String.Format("IIS://{0}/W3SVC/AppPools/{1}", Host, PoolName);
            _directoryEntry = new DirectoryEntry(applicationPoolUrl, User, Password);
        }

        public int GetPoolState()
        {
            int state = 0;
            try
            {
                state = (int) _directoryEntry.InvokeGet("AppPoolState");
                switch (state)
                {
                    case 2:
                        LogWriter.WriteLog(DateTime.Now,
                            "IIS Pool (" + PoolName + ") is " + PoolStatesEnum.Running + " on " + Host);
                        Console.WriteLine("IIS Pool (" + PoolName + ") is " + PoolStatesEnum.Running + " on " + Host);
                        break;
                    case 4:
                        LogWriter.WriteLog(DateTime.Now,
                            "IIS Pool (" + PoolName + ") is " + PoolStatesEnum.Stopped + " on " + Host);
                        Console.WriteLine("IIS Pool (" + PoolName + ") is " + PoolStatesEnum.Stopped + " on " + Host);
                        break;
                }
            }
            catch (COMException exception)
            {
                ComExceptionWorker(exception);
            }
            return state;
        }

        public bool StartPool()
        {
            try
            {
                _directoryEntry.Invoke("Start", null);
                _directoryEntry.Close();
                LogWriter.WriteLog(DateTime.Now, "Starting IIS Pool (" + PoolName + ") is done" + " on " + Host);
                Console.WriteLine("Starting IIS Pool (" + PoolName + ") is done" + " on " + Host);
                return true;
            }
            catch (COMException exception)
            {
                ComExceptionWorker(exception);
                return false;
            }
        }

        public bool StopPool()
        {
            try
            {
                _directoryEntry.Invoke("Stop", null);
                _directoryEntry.Close();
                LogWriter.WriteLog(DateTime.Now, "Stopping IIS Pool (" + PoolName + ") is done" + " on " + Host);
                Console.WriteLine("Stopping IIS Pool (" + PoolName + ") is done" + " on " + Host);
                return true;
            }
            catch (COMException exception)
            {
                ComExceptionWorker(exception);
                return false;
            }
        }

        private static void ComExceptionWorker(COMException exception)
        {
            switch (exception.ErrorCode)
            {
                case -2147463168:
                    LogWriter.WriteLog(DateTime.Now,
                        "Please enable component IIS Metabase and IIS 6 configuration compatibility");
                    Console.WriteLine("Please enable component IIS Metabase and IIS 6 configuration compatibility");
                    Process.Start(Environment.SystemDirectory + @"\OptionalFeatures.exe");
                    break;
                case -2147024893:
                    LogWriter.WriteLog(DateTime.Now,
                        "The system can't find the path specified. Check the configuration file");
                    Console.WriteLine("The system can't find the path specified. Check the configuration file");
                    break;
                case -2147023174:
                    LogWriter.WriteLog(DateTime.Now, "Host unreachable. Check the configuration file");
                    Console.WriteLine("Host unreachable. Check the configuration file");
                    break;
                case -2147024891:
                    LogWriter.WriteLog(DateTime.Now, "Access denied. Need to be an administrator");
                    Console.WriteLine("Access denied. Need to be an administrator");
                    break;
                default:
                    LogWriter.WriteLog(DateTime.Now, exception.Message + ". Undocumented error " + exception.ErrorCode);
                    Console.WriteLine(exception.Message);
                    break;
            }
        }
    }
}