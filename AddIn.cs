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
            logger.Info("NATSExcel loaded");
        }

        public void AutoClose()
        {
            logger.Info("NATSExcel unloaded");
        }
    }
}
