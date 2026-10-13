using System;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using NLog;

namespace NATSExcel
{
    /// <summary>
    /// The add-in entry point. Excel calls AutoOpen when the XLL loads and
    /// AutoClose when it unloads; the COM ProgId that =RTD("NatsRtd",...) needs
    /// is registered here.
    /// </summary>
    public class AddIn : IExcelAddIn
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public void AutoOpen()
        {
            // Register the credential-masking exception renderer BEFORE the first
            // log line. NLog resolves a custom layout renderer by scanning the
            // assembly, but do it explicitly so the order is guaranteed and an
            // unknown renderer never silently logs an unmasked exception.
            try
            {
                // NLog 6: register through the setup builder.
                NLog.LogManager.Setup().SetupExtensions(ext =>
                    ext.RegisterLayoutRenderer<MaskedExceptionRenderer>("maskedException"));
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "AutoOpen: could not register the exception layout renderer");
            }

            // A same-process reload (Excel re-opening the add-in without a new
            // AppDomain) must not reuse the shut-down runtime.
            try
            {
                NatsRuntime.ResetAfterAddInReload();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "AutoOpen: runtime reset failed");
            }
            try
            {
                // The UDF-level state (cached latest values, the async queues) is
                // owned by static maps that a same-process reload would otherwise
                // keep serving from the dead session.
                NatsUdfAsync.ResetAfterAddInReload();
                NatsUdfCore.ResetAfterAddInReload();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "AutoOpen: UDF reset failed");
            }
            try
            {
                // An RTD server Excel never released (no ServerTerminate) keeps
                // running timers alive across the reload; tear every instance down.
                NatsRtd.ResetAfterAddInReload();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "AutoOpen: RTD reset failed");
            }
            try
            {
                ComServer.DllRegisterServer();
            }
            catch (Exception ex)
            {
                // A COM registration failure (no writable hive, policy...) must not
                // abort the whole AutoOpen: UDFs keep working and =RTD shows #N/D
                // until the ProgId is registered.
                logger.Error(ex, "AutoOpen: ComServer.DllRegisterServer failed");
            }
            logger.Info("NATSExcel loaded");
        }

        public void AutoClose()
        {
            try
            {
                NatsRuntime.Shutdown();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "AutoClose: runtime shutdown failed");
            }
            try
            {
                ComServer.DllUnregisterServer();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "AutoClose: ComServer.DllUnregisterServer failed");
            }
            logger.Info("NATSExcel unloaded");
        }
    }
}
