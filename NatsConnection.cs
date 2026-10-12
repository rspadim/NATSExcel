using NLog;

namespace NATSExcel
{
    /// <summary>
    /// Raw NATS protocol client used for the add-in's first milestone: it
    /// connects, reads the INFO banner, authenticates (user/password or token)
    /// and answers PING. The full NATS.Net-based runtime replaces this; keeping
    /// it dependency-light lets the XLL be packed and loaded before the async
    /// client is wired in.
    /// </summary>
    public sealed class NatsConnection
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
    }
}
