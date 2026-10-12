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

        /// <summary>The log file the add-in writes (for support and the E2E
        /// smoke test). Reads the path from the Excel-DNA runtime when the
        /// assembly has no file location (the XLL loads from bytes), so it never
        /// throws.</summary>
        [ExcelFunction(Description = "Returns the log file the add-in writes", IsVolatile = true)]
        public static string NatsUDFLogPath()
        {
            try
            {
                string dir = null;
                var location = typeof(NatsUdf).Assembly.Location;
                if (!string.IsNullOrEmpty(location))
                    dir = System.IO.Path.GetDirectoryName(location);
                if (string.IsNullOrEmpty(dir))
                {
                    try
                    {
                        var xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                        if (!string.IsNullOrEmpty(xll))
                            dir = System.IO.Path.GetDirectoryName(xll);
                    }
                    catch { /* no Excel context */ }
                }
                return System.IO.Path.Combine(dir ?? ".", "NATSExcel.log");
            }
            catch
            {
                return "NATSExcel.log";
            }
        }
    }
}
