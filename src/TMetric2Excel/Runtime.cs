namespace TMetric2Excel
{
    internal class Runtime
    {
        internal static void Log(string msg)
        {
            //if (Logger != null)
            //    Logger.LogInfo(msg);
            //else
            //Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} {msg}");
            Printf(msg);
        }
        internal static void LogError(string msg)
        {
            //if (Logger != null)
            //    Logger.LogError(msg);
            //else
            //Console.WriteLine($"{DateTime.Now.ToString("HH:mm:ss")} {msg}");
            Printf($"!!{msg}");
        }

        internal static void Printf(string msg)
        {
            Console.WriteLine(msg);
        }

        public static void Wait(string msg = "Press any key to continue")
        {
            Console.WriteLine(msg);
            while (!Console.KeyAvailable) { System.Threading.Thread.Sleep(64); }
            var key = Console.ReadKey();
            System.Threading.Thread.Sleep(64);
        }


    }
}