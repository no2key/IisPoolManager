using System;
using System.Threading;

namespace IisPoolManager
{
    internal class DoJob
    {
        private static bool _success;
        private IisPoolWorker _ipw;
        private int _poolState;

        public bool DoWork()
        {
            while (_success != true)
            {
                _ipw = new IisPoolWorker();
                _poolState = _ipw.GetPoolState();
                switch (_poolState)
                {
                    case 2:
                        _success = _ipw.StopPool();
                        break;
                    case 4:
                        _success = _ipw.StartPool();
                        break;
                    case 0:
                        Thread.Sleep(IISPoolManager.Default.interval*100);
                        _success = false;
                        continue;
                }
                _ipw.GetPoolState();
                break;
            }
            return _success;
        }
    }
}