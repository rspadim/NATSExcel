using ExcelDna.Integration;
using NLog;

namespace NATSExcel
{
    /// <summary>
    /// User defined functions exposed to Excel. This first milestone carries
    /// the config/diagnostic surface; the NATS read/write functions land as the
    /// runtime is wired in.
    /// </summary>
    public static class NatsUdf
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>Version string of the loaded add-in (smoke test / support).</summary>
        [ExcelFunction(Description = "Returns the NATSExcel add-in version", IsVolatile = true)]
        public static string NatsUDFVersion()
        {
            return AppConfig.Version;
        }

        /// <summary>Path of the configuration file actually in use, or "(defaults)".</summary>
        [ExcelFunction(Description = "Returns the NATSExcel.json file the add-in is using", IsVolatile = true)]
        public static string NatsUDFConfigPath()
        {
            return AppConfig.ConfigPath ?? "(defaults)";
        }

        /// <summary>Number of live NATS connections the add-in holds. Guarded so a
        /// teardown returns 0 instead of throwing into the cell.</summary>
        [ExcelFunction(Description = "Returns the number of active NATS connections", IsVolatile = true)]
        public static object NatsUDFConnectionCount()
        {
            try
            {
                return NatsRuntime.Connections.LiveConnectionCount();
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>The identity of the running add-in, for support and the
        /// end-to-end smoke test.</summary>
        [ExcelFunction(Description = "Returns the log file the add-in writes", IsVolatile = true)]
        public static string NatsUDFLogPath()
        {
            return System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(typeof(NatsUdf).Assembly.Location) ?? ".",
                "NATSExcel.log");
        }
    }
}
