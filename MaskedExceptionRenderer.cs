using System;
using System.Text;
using NLog;
using NLog.Config;
using NLog.LayoutRenderers;

namespace NATSExcel
{
    /// <summary>
    /// NLog layout renderer that masks credentials inside a logged exception.
    ///
    /// A NATS endpoint may be a full connection string
    /// (<c>nats://user:pass@host</c>, <c>host,password=secret</c>), and a client
    /// exception message embeds that endpoint verbatim. Logging the exception
    /// object (the usual <c>logger.Error(ex, ...)</c>) renders
    /// <c>${exception:format=toString}</c> and would write the secret to the log
    /// file, violating the project rule that credentials are masked in EVERY log
    /// line.
    ///
    /// Register it in <c>NLog.config</c> and use it for the exception field:
    /// <code>layout="... ${message} ${maskedException:format=toString}"</code>
    ///
    /// It is registered by the add-in at load; a renderer unknown to NLog renders
    /// as an empty string, so an old config never throws.
    /// </summary>
    [LayoutRenderer("maskedException")]
    public sealed class MaskedExceptionRenderer : LayoutRenderer
    {
        /// <summary>The inner format, mirroring NLog's own
        /// <c>${exception:format=...}</c> (toString is the default).</summary>
        [DefaultParameter]
        public string Format { get; set; } = "toString";

        /// <summary>The separator between inner exceptions.</summary>
        public string Separator { get; set; } = " " + Environment.NewLine;

        protected override void Append(StringBuilder builder, LogEventInfo logEvent)
        {
            if (logEvent?.Exception == null)
                return;

            var inner = new ExceptionLayoutRenderer { Format = Format, Separator = Separator };
            var raw = inner.Render(logEvent);
            // Mask last: the renderer has already flattened the exception chain,
            // so one pass covers the outer message and every inner one.
            builder.Append(AppConfig.MaskNatsEndpoint(raw));
        }
    }
}
