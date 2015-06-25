using System;

namespace xl2vso
{
    class Program
    {
        static void Main(string[] args)
        {
            VSOConfig.Initialize();
            VSOWorker _worker = new VSOWorker();
            _worker._office_CreateWorkItem();
        }
    }
}
