using System;
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

        /// <summary>
        /// Points the NLog file target at <paramref name="logPath"/> - the path
        /// <see cref="NatsUDFLogPath"/> reports - so the log really lands next to the
        /// add-in. NLog resolves a RELATIVE <c>fileName</c> against
        /// <c>AppDomain.BaseDirectory</c>, which in Excel is the OFFICE folder: the
        /// file went to Program Files while every document, the installer's
        /// uninstaller and this UDF all said "next to the add-in". Returns TRUE when
        /// the target was repointed; FALSE leaves the configured default in place,
        /// because an unwritable add-in folder must never cost the user the log.
        /// </summary>
        internal static bool TryPointLogAt(string logPath)
        {
            try
            {
                if (string.IsNullOrEmpty(logPath))
                    return false;

                string dir = System.IO.Path.GetDirectoryName(logPath);
                if (string.IsNullOrEmpty(dir) || !System.IO.Directory.Exists(dir))
                    return false;

                // Probe first: a folder we cannot write to must keep the default.
                string probe = System.IO.Path.Combine(dir, ".natsexcel-write-probe");
                using (var fs = new System.IO.FileStream(
                    probe, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
                {
                    fs.WriteByte(0);
                }
                System.IO.File.Delete(probe);

                NLog.Targets.FileTarget target =
                    NLog.LogManager.Configuration?.FindTargetByName<NLog.Targets.FileTarget>("file");
                if (target == null)
                    return false;

                target.FileName = logPath;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Serializes a range to JSON, as the standalone twin of the `...JSON`
        /// write functions. Non-volatile: a conversion of its inputs has no
        /// reason to re-run on every recalculation. Pure - it touches no server
        /// and no config.
        /// </summary>
        [ExcelFunction(Description = "Converts a range to a JSON string (non-volatile)")]
        public static object NatsUDFMatrixToJSON(
            [ExcelArgument(Description = "The range to convert")] object[,] range)
        {
            try
            {
                string validation = NatsUdfCommon.RequireRange(range, "range");
                if (validation != null)
                    return validation;
                return NatsResultFormatter.BoundCellText(NatsJson.MatrixToJson(range));
            }
            catch (ArgumentException ex)
            {
                return NatsUdfCommon.Fail(ex.Message);
            }
            catch (Exception ex)
            {
                return NatsUdfCommon.Fail(NatsUdfAsync.SafeMessage(ex));
            }
        }

        /// <summary>
        /// Parses a JSON string into a range, as the standalone twin of the
        /// `...JSON` read functions. Non-volatile and pure, like
        /// <see cref="NatsUDFMatrixToJSON"/>. A JSON `null` element becomes the
        /// optional <paramref name="nullValue"/> (blank by default).
        /// </summary>
        [ExcelFunction(Description = "Converts a JSON string to a range (non-volatile)")]
        public static object NatsUDFJSONToMatrix(
            [ExcelArgument(Description = "The JSON text to convert")] object json,
            [ExcelArgument(Description = "The value to use for a JSON null (optional)")] object nullValue = null)
        {
            try
            {
                string text = NatsUdfCommon.ToNatsString(json);
                if (text == null)
                    return NatsUdfCommon.Fail("a JSON string is required");
                object replacement = nullValue is ExcelDna.Integration.ExcelMissing
                    || nullValue is ExcelDna.Integration.ExcelEmpty
                    ? null
                    : nullValue;
                return NatsJson.JsonToMatrix(text, replacement);
            }
            catch (ArgumentException ex)
            {
                return NatsUdfCommon.Fail(ex.Message);
            }
            catch (Exception ex)
            {
                return NatsUdfCommon.Fail(NatsUdfAsync.SafeMessage(ex));
            }
        }
    }
}
