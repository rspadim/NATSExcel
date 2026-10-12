using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;

namespace NATSExcel
{
    /// <summary>
    /// Shared UDF primitives: the frozen markers and error vocabulary
    /// (DESIGN.md §7, PLAN.md §1.3), argument validation, invariant value
    /// conversion, pair-range flattening, the write-mode rules and the admin /
    /// JetStream gates (PLAN S3.3, S3.6).
    ///
    /// Client-agnostic on purpose: it never references a NATS client type and
    /// never touches the network, so the unit suite exercises it with no server
    /// and no Excel instance.
    /// </summary>
    public static class NatsUdfCommon
    {
        // ------------------------------------------------------------------
        // Frozen markers (PLAN.md §1.3). These exact strings are the contract:
        // spreadsheets and tests compare them literally, so never reword them.
        // ------------------------------------------------------------------

        /// <summary>Shown by an RTD cell before the first value arrives.</summary>
        public const string ConnectData = "(ConnectData)";

        /// <summary>Shown while a subscription topic is (re)subscribing.</summary>
        public const string Subscribed = "(subscribed)";

        /// <summary>A KV/message read that found nothing.</summary>
        public const string NoValue = "(no value)";

        /// <summary>A stored null payload (distinct from a missing value).</summary>
        public const string NullValue = "(null)";

        /// <summary>A write skipped because the payload was unchanged.</summary>
        public const string NoChange = "No change";

        /// <summary>A completed write/operation.</summary>
        public const string Ok = "OK";

        /// <summary>A result-agnostic write sent fire-and-forget.</summary>
        public const string FireAndForget = "OK (fire and forget)";

        /// <summary>A reply-dependent write forced fire-and-forget by the mode.</summary>
        public const string FireAndForgetAll = "OK (fire and forget: all)";

        /// <summary>Request/reply gave up waiting.</summary>
        public const string Timeout = "TIMEOUT";

        /// <summary>Request/reply reached the server but nobody answered.</summary>
        public const string NoResponders = "No Responders";

        /// <summary>Configuration problem (never a silent localhost fallback).</summary>
        public const string ConfigError = "#CONFIG";

        // ------------------------------------------------------------------
        // Frozen bounds (PLAN.md §1.3).
        // ------------------------------------------------------------------

        /// <summary>Hard Excel cell-text limit; a longer result is unusable.</summary>
        public const int MaxCellTextLength = 32767;

        /// <summary>Object Store text bound; a larger object is not a text cell.</summary>
        public const int MaxObjectCellBytes = 1_048_576;

        // ------------------------------------------------------------------
        // Error vocabulary. Every helper returns null when the argument is fine,
        // or the text to hand to a cell (the caller wraps it with Fail).
        // ------------------------------------------------------------------

        /// <summary>Prefix of every terse, stable cell error.</summary>
        public const string ErrorPrefix = "Error: ";

        /// <summary>A required range argument was omitted.</summary>
        public const string ErrorRangeRequired = "a range is required";

        /// <summary>An Excel error cell reached a value position.</summary>
        public const string ErrorExcelErrorCell = "Excel error cells are not valid arguments";

        /// <summary>A whole range was used where a single value belongs.</summary>
        public const string ErrorMultiCellScalar = "A multi-cell range is not a valid scalar argument";

        /// <summary>A pair range is neither a 2-column nor a 2-row layout.</summary>
        public const string ErrorRangeShape = "expected a range with 2 columns or 2 rows";

        /// <summary>Keys and values ranges hold a different number of cells.</summary>
        public const string ErrorCellCountMismatch = "keys and values must have the same number of cells";

        /// <summary>The write had no usable entry after blank cells were skipped.</summary>
        public const string ErrorNoEntries = "no entries to write";

        /// <summary>The range cannot fit an Excel worksheet.</summary>
        public const string ErrorRangeTooLarge = "range is too large for an Excel sheet";

        /// <summary>Named cause: the server runs without JetStream (DESIGN.md §6).</summary>
        public const string ErrorJetStreamDisabled = "JetStream is not enabled on this server";

        /// <summary>Named cause: an admin/destructive call with the gate closed (§7.3).</summary>
        public const string ErrorAdminWritesDisabled = "admin writes are disabled (set \"AllowAdminWrites\": true)";

        // ------------------------------------------------------------------
        // Write-mode names (PLAN.md §1.4 / DESIGN.md §8.1 SyncWrite).
        // ------------------------------------------------------------------

        /// <summary>Wait for the JetStream ack / flush.</summary>
        public const string SyncWriteSync = "sync";

        /// <summary>Fire and forget for result-agnostic writes only (default).</summary>
        public const string SyncWriteFireForget = "fireforget";

        /// <summary>Fire and forget everything, reply-dependent writes included.</summary>
        public const string SyncWriteFireForgetAll = "fireforget-all";

        // Excel worksheet limits, applied before a range is materialized flatter.
        private const int MaxExcelRows = 1048576;
        private const int MaxExcelColumns = 16384;
        // Total-cell budget: the per-dimension checks alone let a wide-and-tall
        // range (billions of cells) reach the allocation before failing.
        private const long MaxExcelCells = 50000000;

        // ------------------------------------------------------------------
        // Credential masking seam.
        // ------------------------------------------------------------------

        /// <summary>
        /// Masks every credential form (URL userinfo, token=/user=/password=,
        /// JWT, NKey seed, creds/TLS paths) before text reaches a cell or a log
        /// line (DESIGN.md §8, PLAN F1.5).
        ///
        /// NOTE: <c>AppConfig.MaskNatsEndpoint</c> is planned (PLAN F1.5) and is
        /// not implemented in <c>Config.cs</c> yet. This is the only call site in
        /// this file, so wiring it is a one-line change in Config.cs; until then
        /// the add-in project does not compile.
        /// </summary>
        private static string Mask(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            return AppConfig.MaskNatsEndpoint(text);
        }

        // ==================================================================
        // Value conversion
        // ==================================================================

        /// <summary>
        /// Converts an Excel cell / programmatic value to the invariant text
        /// written to NATS. Numbers use the invariant culture (so 67000.5 does
        /// not become 67000,5 on a comma-decimal locale), booleans become
        /// true/false, a DateTime becomes ISO 8601 round-trip text and a truly
        /// missing cell (null / ExcelMissing / ExcelEmpty) becomes null.
        ///
        /// Throws <see cref="ArgumentException"/> (the frozen terse message) for a
        /// multi-cell range or an Excel error cell, exactly like the sibling
        /// project; the Require* helpers convert that into an error string.
        /// </summary>
        public static string ToNatsString(object value)
        {
            if (value is Array)
                throw new ArgumentException(ErrorMultiCellScalar);
            if (value is ExcelError)
                throw new ArgumentException(ErrorExcelErrorCell);
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                return null;
            if (value is string text)
                return text;
            if (value is bool flag)
                return flag ? "true" : "false";
            // ISO 8601 round-trip text so JSON wrappers can parse a date/time
            // value (only reachable programmatically, not from a cell).
            if (value is DateTime dateTime)
                return dateTime.ToString("o", CultureInfo.InvariantCulture);
            if (value is double d)
            {
                // G15 keeps common values compact ("0.1" stays "0.1"); fall back
                // to G17 only when G15 does not round-trip (e.g. double.MaxValue).
                string formatted = d.ToString("G15", CultureInfo.InvariantCulture);
                if (!(double.TryParse(formatted, NumberStyles.Float, CultureInfo.InvariantCulture, out var back) && back == d))
                    formatted = d.ToString("G17", CultureInfo.InvariantCulture);
                // -0.0 formats as "0" and the round-trip check above accepts it
                // (0.0 == -0.0), silently dropping the sign bit. Keep the sign
                // (double.IsNegative is .NET Core only, so inspect the bits).
                if (d == 0.0 && BitConverter.DoubleToInt64Bits(d) < 0)
                    return "-0";
                return formatted;
            }
            // float is not an Excel cell type but can arrive from a caller;
            // format it with round-trip precision too (its plain ToString() uses
            // only 7 digits and silently changes the value).
            if (value is float f)
            {
                if (f == 0f && BitConverter.ToInt32(BitConverter.GetBytes(f), 0) < 0)
                    return "-0";
                return f.ToString("R", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Invariant integral conversion used by numeric arguments (sequences,
        /// counts, revisions, timeouts). Returns false, never throws, for a
        /// missing/blank cell, an Excel error cell, a multi-cell range, a
        /// boolean, a fraction, a non-numeric text or an out-of-range value.
        /// </summary>
        public static bool TryToLong(object value, out long result)
        {
            result = 0;
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                return false;
            if (value is Array || value is ExcelError)
                return false;
            // Booleans are not integers here (matches the sibling's numeric
            // guard): a numeric argument must be a real number, not TRUE/FALSE.
            if (value is bool)
                return false;

            if (value is double d)
            {
                if (double.IsNaN(d) || double.IsInfinity(d) || d != Math.Truncate(d))
                    return false;
                // 2^63 as a double: the cast is only defined strictly below it.
                if (d < -9223372036854775808.0 || d >= 9223372036854775808.0)
                    return false;
                result = (long)d;
                return true;
            }
            if (value is float f)
            {
                if (float.IsNaN(f) || float.IsInfinity(f) || f != Math.Truncate(f))
                    return false;
                if (f < -9223372036854775808.0f || f >= 9223372036854775808.0f)
                    return false;
                result = (long)f;
                return true;
            }
            if (value is decimal m)
            {
                if (m != Math.Truncate(m) || m < long.MinValue || m > long.MaxValue)
                    return false;
                result = (long)m;
                return true;
            }
            if (value is string text)
                return long.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result);

            try
            {
                result = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                return true;
            }
            catch (OverflowException) { result = 0; return false; }
            catch (FormatException) { result = 0; return false; }
            catch (InvalidCastException) { result = 0; return false; }
        }

        /// <summary>
        /// Integer counterpart of <see cref="TryToLong"/>; false when the value
        /// is not integral or does not fit an <see cref="int"/> (a sequence or a
        /// count never silently wraps).
        /// </summary>
        public static bool TryToInt(object value, out int result)
        {
            long parsed;
            if (!TryToLong(value, out parsed))
            {
                result = 0;
                return false;
            }
            if (parsed < int.MinValue || parsed > int.MaxValue)
            {
                result = 0;
                return false;
            }
            result = (int)parsed;
            return true;
        }

        // ==================================================================
        // Argument validation: null = fine, otherwise the cell error text
        // ==================================================================

        /// <summary>
        /// Validates a required scalar text argument (bucket, key, subject,
        /// object name, stream, consumer...). Returns null when usable, or
        /// "Error: a &lt;what&gt; is required" for a missing/blank cell, and the
        /// frozen terse error for an Excel error cell or a multi-cell range.
        /// Required identifiers reject blank text: a blank bucket or subject
        /// matches nothing and must fail before any I/O.
        /// </summary>
        public static string RequireText(object value, string what)
        {
            string label = string.IsNullOrWhiteSpace(what) ? "value" : what;
            string text;
            try
            {
                text = ToNatsString(value);
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message);
            }
            if (string.IsNullOrWhiteSpace(text))
                return Fail("a " + label + " is required");
            return null;
        }

        /// <summary>
        /// Validates a NATS subject. Wildcards (`*` = one token, `&gt;` = one or
        /// more) are subscribe-only: pass <paramref name="allowWildcard"/> false
        /// for a publish, which needs a concrete subject.
        /// </summary>
        public static string RequireSubject(object value, bool allowWildcard)
        {
            string subject;
            try
            {
                subject = ToNatsString(value);
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message);
            }
            if (string.IsNullOrWhiteSpace(subject))
                return Fail("a non-empty subject is required");
            if (!allowWildcard && ContainsWildcardToken(subject))
                return Fail("a non-empty subject is required (a publish needs a concrete subject; wildcards are subscribe-only)");
            return null;
        }

        /// <summary>Validates a required KV bucket name.</summary>
        public static string RequireBucket(object value)
        {
            return RequireText(value, "bucket");
        }

        /// <summary>Validates a required JetStream stream name.</summary>
        public static string RequireStream(object value)
        {
            return RequireText(value, "stream");
        }

        /// <summary>Validates a required JetStream consumer name.</summary>
        public static string RequireConsumer(object value)
        {
            return RequireText(value, "consumer");
        }

        /// <summary>Validates a required Object Store object name.</summary>
        public static string RequireObjectName(object value)
        {
            return RequireText(value, "object name");
        }

        /// <summary>Validates a required KV key.</summary>
        public static string RequireKey(object value)
        {
            return RequireText(value, "key");
        }

        /// <summary>
        /// Validates a required range argument: null / empty, a range that
        /// cannot fit a worksheet, or a range holding an Excel error cell is
        /// rejected before it is flattened. Null means the range is usable.
        /// </summary>
        public static string RequireRange(object[,] range, string what = "range")
        {
            string label = string.IsNullOrWhiteSpace(what) ? "range" : what;
            if (range == null)
                return Fail("a " + label + " is required");
            if (range.GetLength(0) == 0 || range.GetLength(1) == 0)
                return Fail("a " + label + " is required");
            try
            {
                EnsureRangeFits(range);
            }
            catch (ArgumentException ex)
            {
                return Fail(ex.Message);
            }
            foreach (object cell in range)
            {
                if (cell is ExcelError)
                    return Fail(ErrorExcelErrorCell);
            }
            return null;
        }

        /// <summary>
        /// Admin gate (PLAN S3.6, DESIGN.md §7.3): admin and destructive
        /// functions (stream/consumer create-update-delete, purge, delete
        /// message, KV/Object delete) return the named error unless the config
        /// opts in. Null means the call may proceed.
        /// </summary>
        public static string RequireAdminWrites(bool allowAdminWrites)
        {
            return allowAdminWrites ? null : Fail(ErrorAdminWritesDisabled);
        }

        /// <summary>
        /// JetStream gate (PLAN S3.6): a JS/KV/OBJ call on a server started
        /// without -js returns the named cause, never a hang and never an empty
        /// result that looks healthy. Null means JetStream is available.
        /// </summary>
        public static string RequireJetStream(bool jetStreamEnabled)
        {
            return jetStreamEnabled ? null : Fail(ErrorJetStreamDisabled);
        }

        // ==================================================================
        // Error building
        // ==================================================================

        /// <summary>
        /// Wraps a message in the frozen "Error: " prefix. Argument-validation
        /// text is intentionally NOT masked: the terse strings are a pinned
        /// contract and never carry credentials. Runtime causes go through
        /// <see cref="VerboseError"/>, which masks every part.
        /// </summary>
        public static string Fail(string message)
        {
            return ErrorPrefix + (message ?? string.Empty);
        }

        /// <summary>
        /// Cell text for a runtime failure: "Error: &lt;context&gt;: &lt;cause&gt;
        /// | &lt;hint&gt;". Every part is masked (golden rule: no credential in a
        /// cell) and bounded, so a server error naming a whole cluster cannot
        /// produce an unusable multi-KB error cell. The hint is omitted when
        /// null/blank.
        /// </summary>
        public static string VerboseError(string what, string cause, string hint)
        {
            string context = BoundContext(Mask(what));
            string causeText = BoundContext(Mask(cause));
            if (string.IsNullOrWhiteSpace(context))
                context = "(no context)";
            if (string.IsNullOrWhiteSpace(causeText))
                causeText = "unknown error";

            string text = context + ": " + causeText;
            if (!string.IsNullOrWhiteSpace(hint))
                text += " | " + BoundContext(Mask(hint));
            return Fail(text);
        }

        /// <summary>
        /// Short, actionable hint for the recurrent failure classes named in
        /// PLAN §1.3; null when the cause speaks for itself (the caller then
        /// omits the " | hint" part).
        /// </summary>
        public static string ErrorHint(string cause)
        {
            if (string.IsNullOrEmpty(cause))
                return null;
            if (Contains(cause, "jetstream not enabled") || Contains(cause, "jetstream is not enabled"))
                return "start nats-server with -js (or add a jetstream {} block) to enable JetStream";
            if (Contains(cause, "no suitable peers for placement"))
                return "the cluster is too small for the stream's replicas; add a node or lower num_replicas";
            if (Contains(cause, "no responders"))
                return "no subscriber answered the request; check the subject and that a responder is running";
            if (Contains(cause, "wrong last revision"))
                return "the key changed since it was read; re-read the revision and retry";
            if (Contains(cause, "consumer") && Contains(cause, "not found"))
                return "create the consumer (NatsUDFConsumerCreate) or check the consumer name";
            if (Contains(cause, "stream") && Contains(cause, "not found"))
                return "create the stream (NatsUDFStreamCreate) or check the stream name";
            if (Contains(cause, "timeout") || Contains(cause, "timed out"))
                return "the server did not answer in time; raise UDF.timeout (NATSExcel.json) or check the connection";
            return null;
        }

        // Named-cause builders (PLAN §1.3). They produce the cause text that
        // VerboseError/ErrorHint then shape into a cell error.
        /// <summary>Named cause for a missing stream.</summary>
        public static string StreamNotFound(string stream)
        {
            return "stream '" + (stream ?? string.Empty) + "' not found";
        }

        /// <summary>Named cause for a missing consumer on a stream.</summary>
        public static string ConsumerNotFound(string consumer, string stream)
        {
            return "consumer '" + (consumer ?? string.Empty) + "' not found on stream '" + (stream ?? string.Empty) + "'";
        }

        /// <summary>Named cause for a KV compare-and-set against a stale revision.</summary>
        public static string WrongLastRevision(string key, long expectedRevision)
        {
            return "wrong last revision for key '" + (key ?? string.Empty) + "' (expected "
                + expectedRevision.ToString(CultureInfo.InvariantCulture) + ")";
        }

        // ==================================================================
        // Log safety
        // ==================================================================

        /// <summary>
        /// Caps a value embedded in a log line: a payload can be arbitrarily
        /// large, so only the head is logged (a 10 MB value must not produce a
        /// 10 MB log line). Size only - masking is the caller's job via
        /// <see cref="VerboseError"/> or the log site's own masking.
        /// </summary>
        public static string ShortForLog(string value, int max = 64)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            if (max < 1)
                max = 1;
            return value.Length <= max ? value : value.Substring(0, max) + "...(truncated)";
        }

        /// <summary>
        /// Caps a diagnostic context (call-site context, server cause) so an
        /// error cell/log line stays readable. Same bound as the sibling
        /// project's context bound.
        /// </summary>
        public static string BoundContext(string value, int max = 200)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            if (max < 1)
                max = 1;
            return value.Length <= max ? value : value.Substring(0, max) + "...";
        }

        // ==================================================================
        // Write modes (DESIGN.md §8.1, PLAN S3.3)
        // ==================================================================

        /// <summary>
        /// Frozen fire-and-forget rule: "sync" never fires and forgets;
        /// "fireforget-all" always does, reply-dependent writes included; the
        /// default "fireforget" does it only for result-agnostic writes. An
        /// unknown/blank mode falls back to the default.
        /// </summary>
        public static bool ShouldFireAndForget(string syncWriteMode, bool replyDependent)
        {
            string mode = NormalizeMode(syncWriteMode);
            if (mode == SyncWriteSync)
                return false;
            if (mode == SyncWriteFireForgetAll)
                return true;
            return !replyDependent;
        }

        /// <summary>
        /// Result text shown when a write was sent fire-and-forget: there is no
        /// reply to report, so an explicit marker takes its place. The default
        /// mode only fires result-agnostic writes, so its marker is
        /// "OK (fire and forget)"; "fireforget-all" also fires reply-dependent
        /// writes, and the ": all" marker explains why a call that normally
        /// returns data returned none.
        /// </summary>
        public static string FireAndForgetMarker(string syncWriteMode)
        {
            return NormalizeMode(syncWriteMode) == SyncWriteFireForgetAll ? FireAndForgetAll : FireAndForget;
        }

        // ==================================================================
        // Range flattening
        // ==================================================================

        /// <summary>
        /// Reads a key/value pair range: exactly 2 columns is read row by row
        /// (vertical pairs), otherwise exactly 2 rows is read column by column
        /// (horizontal pairs) - so a 2x2 range yields two vertical pairs. Any
        /// other shape throws the frozen
        /// "expected a range with 2 columns or 2 rows"; a truly missing key
        /// (null / blank cell) skips its row, and an empty result throws
        /// "no entries to write". A missing value becomes "" so row positions
        /// survive.
        /// </summary>
        public static List<KeyValuePair<string, string>> FlattenPairRange(object[,] range)
        {
            if (range == null)
                throw new ArgumentException(ErrorRangeRequired);
            EnsureRangeFits(range);

            int rows = range.GetLength(0);
            int cols = range.GetLength(1);
            if (rows == 0 || cols == 0)
                throw new ArgumentException(ErrorNoEntries);

            var pairs = new List<KeyValuePair<string, string>>();
            if (cols == 2)
            {
                for (int r = 0; r < rows; r++)
                    AddPair(pairs, range[r, 0], range[r, 1]);
            }
            else if (rows == 2)
            {
                for (int c = 0; c < cols; c++)
                    AddPair(pairs, range[0, c], range[1, c]);
            }
            else
            {
                throw new ArgumentException(ErrorRangeShape);
            }

            if (pairs.Count == 0)
                throw new ArgumentException(ErrorNoEntries);
            return pairs;
        }

        /// <summary>
        /// Reads two parallel ranges (keys and values) row-major into pairs -
        /// the shape of a batch KV/header write, where the two ranges are
        /// separate arguments. A differing cell count throws the frozen
        /// "keys and values must have the same number of cells"; a truly
        /// missing key skips its entry; an empty result throws
        /// "no entries to write".
        /// </summary>
        public static List<KeyValuePair<string, string>> FlattenPairRange(object[,] keys, object[,] values)
        {
            if (keys == null || values == null)
                throw new ArgumentException(ErrorRangeRequired);
            EnsureRangeFits(keys);
            EnsureRangeFits(values);

            var keyCells = FlattenRowMajor(keys);
            var valueCells = FlattenRowMajor(values);
            if (keyCells.Count == 0 || keyCells.Count != valueCells.Count)
                throw new ArgumentException(ErrorCellCountMismatch);

            var pairs = new List<KeyValuePair<string, string>>(keyCells.Count);
            for (int i = 0; i < keyCells.Count; i++)
                AddPair(pairs, keyCells[i], valueCells[i]);

            if (pairs.Count == 0)
                throw new ArgumentException(ErrorNoEntries);
            return pairs;
        }

        /// <summary>
        /// Rejects a range that cannot fit a worksheet before it is flattened or
        /// echoed, so the cell shows a stable size error instead of allocating a
        /// multi-GB list on the Excel thread. A null range is a no-op (the
        /// presence check belongs to RequireRange).
        /// </summary>
        public static void EnsureRangeFits(object[,] range)
        {
            if (range == null)
                return;
            EnsureRangeFits(range.GetLength(0), range.GetLength(1));
        }

        /// <summary>Dimension-only counterpart, exposed internally so the guard
        /// is testable without allocating the huge array (sibling pattern).</summary>
        internal static void EnsureRangeFits(int rows, int cols)
        {
            if (rows < 0 || cols < 0 || rows > MaxExcelRows || cols > MaxExcelColumns
                || (long)rows * cols > MaxExcelCells)
                throw new ArgumentException(ErrorRangeTooLarge);
        }

        // ==================================================================
        // Internals
        // ==================================================================

        /// <summary>Appends one pair; a truly missing key skips the entry and a
        /// missing value becomes "" (an ExcelError cell throws, per ToNatsString).</summary>
        private static void AddPair(List<KeyValuePair<string, string>> pairs, object key, object value)
        {
            string keyText = ToNatsString(key);
            if (keyText == null)
                return;
            pairs.Add(new KeyValuePair<string, string>(keyText, ToNatsString(value) ?? string.Empty));
        }

        /// <summary>Row-major cell list (the order Excel uses), guarded so the
        /// backing list is bounded before it is allocated.</summary>
        private static List<object> FlattenRowMajor(object[,] range)
        {
            int rows = range.GetLength(0);
            int cols = range.GetLength(1);
            var cells = new List<object>(rows * cols);
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    cells.Add(range[r, c]);
            return cells;
        }

        /// <summary>True when a subject token is exactly "*" or "&gt;".
        /// Token-based, matching NATS: "a*b" is not a wildcard.</summary>
        private static bool ContainsWildcardToken(string subject)
        {
            int start = 0;
            while (start <= subject.Length)
            {
                int dot = subject.IndexOf('.', start);
                int end = dot < 0 ? subject.Length : dot;
                if (end - start == 1 && (subject[start] == '*' || subject[start] == '>'))
                    return true;
                if (dot < 0)
                    return false;
                start = dot + 1;
            }
            return false;
        }

        /// <summary>Trims and lowercases a write mode (invariant); an unknown or
        /// blank mode resolves to the "fireforget" default.</summary>
        private static string NormalizeMode(string syncWriteMode)
        {
            if (string.IsNullOrWhiteSpace(syncWriteMode))
                return SyncWriteFireForget;
            string mode = syncWriteMode.Trim().ToLowerInvariant();
            if (mode == SyncWriteSync || mode == SyncWriteFireForgetAll)
                return mode;
            return SyncWriteFireForget;
        }

        /// <summary>Ordinal-ignore-case substring test used by the hint map.</summary>
        private static bool Contains(string haystack, string needle)
        {
            return haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
