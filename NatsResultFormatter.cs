using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NATSExcel
{
    /// <summary>
    /// Formatting helpers for values sent to Excel (DESIGN.md §7, PLAN.md S3.4).
    /// <para>
    /// This file is intentionally dependency-free: every method takes plain CLR
    /// values (string, long, int, bool, DateTime, byte[]) and never a NATS.Net type,
    /// so it can be linked into any test project without pulling a client.
    /// </para>
    /// <para>
    /// <b>Determinism contract.</b> Every JSON producer emits compact output with a
    /// fixed key order (the order the fields are passed below). DateTimes are
    /// normalized to UTC with <see cref="DateTimeZoneHandling.Utc"/> and serialized
    /// ISO-8601, so two machines in different time zones produce the same text; a
    /// null parameter is emitted as JSON <c>null</c>, never an empty string (except
    /// collections, which become <c>[]</c>). Tests assert exact text, so do not
    /// reorder the field lists.
    /// </para>
    /// </summary>
    public static class NatsResultFormatter
    {
        /// <summary>Excel's hard limit on characters in a text cell (PLAN.md §1.3 bounds).</summary>
        public const int MaxCellTextLength = 32767;

        /// <summary>
        /// The single cell a zero-row matrix formatter returns. A matrix function
        /// must never return an empty range: Excel renders an empty range as a
        /// blank cell, which looks like a broken formula rather than "no rows".
        /// </summary>
        public const string EmptyCell = "(empty)";

        /// <summary>
        /// Compact, invariant, UTC-normalized serializer settings shared by every
        /// JSON producer. No indentation (a pretty-printed string would waste the
        /// 32,767-char cell budget and break exact-text tests).
        /// </summary>
        private static readonly JsonSerializerSettings CompactSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.None,
            NullValueHandling = NullValueHandling.Include,
            DateFormatHandling = DateFormatHandling.IsoDateFormat,
            DateTimeZoneHandling = DateTimeZoneHandling.Utc,
            Culture = CultureInfo.InvariantCulture,
            MaxDepth = 64
        };

        /// <summary>
        /// UTF-8 with replacement fallback: invalid sequences decode to U+FFFD
        /// instead of throwing, and no BOM is prepended on encode-equivalent paths.
        /// </summary>
        private static readonly UTF8Encoding Utf8Lenient = new UTF8Encoding(false, false);

        // ------------------------------------------------------------------
        // Cell bounding / payload decoding
        // ------------------------------------------------------------------

        /// <summary>
        /// Clamps a string to <see cref="MaxCellTextLength"/>. A string at or under
        /// the limit is returned unchanged; a longer one keeps its head and gets a
        /// <c>...(truncated N chars)</c> suffix, where N is the number of characters
        /// dropped. The returned string is always within the Excel limit (the suffix
        /// length is solved as a fixed point, since N changes the suffix width).
        /// null/empty pass through unchanged; this never throws.
        /// </summary>
        public static string BoundCellText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            if (text.Length <= MaxCellTextLength)
                return text;

            // The head/suffix split is a fixed point: the suffix width depends on
            // the number of dropped characters, which depends on the head width.
            // Starting at "keep everything" is within a couple of characters of the
            // answer, so this converges in two or three passes; the cap is a guard.
            int keep = MaxCellTextLength;
            for (int i = 0; i < 8; i++)
            {
                int dropped = text.Length - keep;
                int nextKeep = MaxCellTextLength - TruncationSuffix(dropped).Length;
                if (nextKeep == keep)
                    break;
                keep = nextKeep;
            }

            if (keep < 0)
                keep = 0;
            // Never split a surrogate pair: a lone high surrogate is invalid UTF-16
            // and Excel renders it as garbage. If the cut lands between a high and
            // a low surrogate, drop the high one too (the pair is atomic).
            if (keep > 0 && keep < text.Length
                && char.IsHighSurrogate(text[keep - 1]) && char.IsLowSurrogate(text[keep]))
                keep--;
            return text.Substring(0, keep) + TruncationSuffix(text.Length - keep);
        }

        /// <summary>The truncation suffix for a given number of dropped characters.</summary>
        private static string TruncationSuffix(int dropped)
        {
            return "...(truncated " + dropped.ToString(CultureInfo.InvariantCulture) + " chars)";
        }

        /// <summary>
        /// TRUE when two payloads are byte-for-byte identical, used by the
        /// <c>SkipRepeatedMessages</c> dedup and by the RTD change detectors.
        /// <para>
        /// It compares the RAW BYTES on purpose. A string/numeric comparison would
        /// be wrong in two ways that both leave a stale cell on screen:
        /// </para>
        /// <list type="bullet">
        /// <item><description>
        /// A RedisValue-style equality compares numeric-looking strings numerically,
        /// so <c>"007"</c> would equal <c>"7"</c> and a genuine change would be
        /// swallowed. NATS subjects and payloads are opaque text, not numbers.
        /// </description></item>
        /// <item><description>
        /// Decoding to text first maps every invalid UTF-8 sequence to U+FFFD, so
        /// two DIFFERENT invalid payloads (<c>\x80</c> vs <c>\x81</c>) would compare
        /// equal after decoding.
        /// </description></item>
        /// </list>
        /// <para>
        /// Null-safe: null equals null, and null never equals a non-null (even a
        /// zero-length) payload.
        /// </para>
        /// </summary>
        public static bool SameMessage(byte[] a, byte[] b)
        {
            if (ReferenceEquals(a, b))
                return true;
            if (a == null || b == null)
                return false;
            if (a.Length != b.Length)
                return false;

            // A manual ordinal byte loop: no allocation, no culture, no locale and
            // no BCL version dependency (the sibling project's span compare needs
            // System.Memory; this file stays dependency-free).
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Decodes a NATS payload as UTF-8. Invalid byte sequences become U+FFFD
        /// (replacement character) and this never throws: a malformed payload must
        /// still produce a readable cell, not a failed formula. A null payload
        /// returns null (the caller decides whether that is the <c>(null)</c>
        /// marker); a zero-length payload returns the empty string. A leading BOM is
        /// NOT stripped: it is real payload content (unlike a JSON document, where
        /// it is a transport artifact).
        /// </summary>
        public static string DecodePayload(byte[] data)
        {
            if (data == null)
                return null;
            if (data.Length == 0)
                return string.Empty;

            try
            {
                return Utf8Lenient.GetString(data);
            }
            catch (Exception)
            {
                // Utf8Lenient does not throw for bad bytes, but a defensive fallback
                // keeps the "never throw" guarantee absolute.
                return string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // KV / Object / JetStream / Server / Stats JSON (deterministic key order)
        // ------------------------------------------------------------------

        /// <summary>
        /// KV status as compact JSON, fields in this order:
        /// <c>{"revision":&lt;long&gt;,"created":&lt;iso|null&gt;,"updated":&lt;iso|null&gt;,"operation":&lt;string|null&gt;}</c>.
        /// Example: <c>{"revision":3,"created":"2026-01-02T03:04:05Z","updated":null,"operation":"PUT"}</c>.
        /// DateTimes are normalized to UTC (Kind is ignored) so the text is stable
        /// across time zones.
        /// </summary>
        public static string FormatKvStatusJson(long revision, DateTime? created, DateTime? updated, string operation)
        {
            return OrderedJson(
                "revision", revision,
                "created", created,
                "updated", updated,
                "operation", operation);
        }

        /// <summary>
        /// Object Store metadata as compact JSON, fields in this order:
        /// <c>{"name":&lt;string|null&gt;,"digest":&lt;string|null&gt;,"size":&lt;long&gt;,"chunks":&lt;int&gt;,"description":&lt;string|null&gt;}</c>.
        /// Example: <c>{"name":"report.pdf","digest":"SHA-256=...","size":1048576,"chunks":1,"description":null}</c>.
        /// The digest is the server-side object digest; this formatter never reads
        /// or embeds the object body (DESIGN.md §7.5).
        /// </summary>
        public static string FormatObjectInfoJson(string name, string digest, long size, int chunks, string description)
        {
            return OrderedJson(
                "name", name,
                "digest", digest,
                "size", size,
                "chunks", chunks,
                "description", description);
        }

        /// <summary>
        /// Stream info as compact JSON, fields in this order:
        /// <c>{"name":&lt;string|null&gt;,"subjects":[&lt;string&gt;...],"messages":&lt;long&gt;,"bytes":&lt;long&gt;,"first_sequence":&lt;long&gt;,"last_sequence":&lt;long&gt;,"consumers":&lt;long&gt;,"storage":&lt;string|null&gt;}</c>.
        /// A null <paramref name="subjects"/> becomes <c>[]</c>. Sequence numbers are
        /// passed as <see cref="long"/> values (callers convert the client's ulong;
        /// real sequences are far below 2^63).
        /// </summary>
        public static string FormatStreamInfoJson(
            string name,
            IEnumerable<string> subjects,
            long messages,
            long bytes,
            long firstSequence,
            long lastSequence,
            long consumerCount,
            string storage)
        {
            return OrderedJson(
                "name", name,
                "subjects", subjects,
                "messages", messages,
                "bytes", bytes,
                "first_sequence", firstSequence,
                "last_sequence", lastSequence,
                "consumers", consumerCount,
                "storage", storage);
        }

        /// <summary>
        /// Consumer info as compact JSON, fields in this order:
        /// <c>{"stream":&lt;string|null&gt;,"name":&lt;string|null&gt;,"durable":&lt;bool&gt;,"pending":&lt;long&gt;,"redelivered":&lt;long&gt;,"ack_pending":&lt;long&gt;,"waiting":&lt;long&gt;,"delivered_sequence":&lt;long&gt;,"ack_floor_sequence":&lt;long&gt;}</c>.
        /// <paramref name="pending"/> is the number of messages the consumer still
        /// has to deliver; <paramref name="waiting"/> is the pull-request queue depth.
        /// </summary>
        public static string FormatConsumerInfoJson(
            string stream,
            string name,
            bool durable,
            long pending,
            long redelivered,
            long ackPending,
            long waiting,
            long deliveredSequence,
            long ackFloorSequence)
        {
            return OrderedJson(
                "stream", stream,
                "name", name,
                "durable", durable,
                "pending", pending,
                "redelivered", redelivered,
                "ack_pending", ackPending,
                "waiting", waiting,
                "delivered_sequence", deliveredSequence,
                "ack_floor_sequence", ackFloorSequence);
        }

        /// <summary>
        /// Server info as compact JSON, fields in this order:
        /// <c>{"server_id":...,"server_name":...,"version":...,"go":...,"host":...,"port":&lt;int&gt;,"cluster":...,"connect_urls":[&lt;string&gt;...],"max_payload":&lt;long&gt;,"proto":&lt;int&gt;,"jetstream":&lt;bool&gt;}</c>.
        /// <paramref name="cluster"/> and <paramref name="connectUrls"/> are the
        /// Tier-1 topology fields of DESIGN.md §7.8 (a standalone server reports an
        /// empty cluster and a single connect URL); a null <paramref name="connectUrls"/>
        /// becomes <c>[]</c>. Endpoints here must already be masked by the caller.
        /// </summary>
        public static string FormatServerInfoJson(
            string serverId,
            string serverName,
            string version,
            string go,
            string host,
            int port,
            string cluster,
            IEnumerable<string> connectUrls,
            long maxPayload,
            int proto,
            bool jetStream)
        {
            return OrderedJson(
                "server_id", serverId,
                "server_name", serverName,
                "version", version,
                "go", go,
                "host", host,
                "port", port,
                "cluster", cluster,
                "connect_urls", connectUrls,
                "max_payload", maxPayload,
                "proto", proto,
                "jetstream", jetStream);
        }

        /// <summary>
        /// Connection stats as compact JSON, fields in this order:
        /// <c>{"in_msgs":&lt;long&gt;,"out_msgs":&lt;long&gt;,"in_bytes":&lt;long&gt;,"out_bytes":&lt;long&gt;}</c>.
        /// Example: <c>{"in_msgs":12,"out_msgs":3,"in_bytes":480,"out_bytes":96}</c>.
        /// </summary>
        public static string FormatStatsJson(long inMsgs, long outMsgs, long inBytes, long outBytes)
        {
            return OrderedJson(
                "in_msgs", inMsgs,
                "out_msgs", outMsgs,
                "in_bytes", inBytes,
                "out_bytes", outBytes);
        }

        /// <summary>
        /// A list of keys as a compact JSON string array, in the order given (no
        /// sorting; the caller decides ordering). null or empty becomes <c>[]</c>.
        /// Example: <c>["alpha","beta"]</c>.
        /// </summary>
        public static string FormatKeysJson(IEnumerable<string> keys)
        {
            var list = keys == null ? new List<string>() : new List<string>(keys);
            return JsonConvert.SerializeObject(list, CompactSettings);
        }

        // ------------------------------------------------------------------
        // Excel matrices
        // ------------------------------------------------------------------

        /// <summary>
        /// A single-column N×1 <see cref="object[,]"/> for Excel: cell <c>[i,0]</c> is
        /// row <c>i</c> (bounded by <see cref="BoundCellText"/>). Backs KV keys,
        /// bucket lists and any other single-column list. A null or empty input
        /// returns a 1×1 matrix whose only cell is <see cref="EmptyCell"/>
        /// (<c>"(empty)"</c>), so the range is never blank.
        /// </summary>
        public static object[,] FormatStringMatrix(IEnumerable<string> rows)
        {
            var list = rows == null ? new List<string>() : new List<string>(rows);
            if (list.Count == 0)
                return SingleCell(EmptyCell);

            var matrix = new object[list.Count, 1];
            for (int i = 0; i < list.Count; i++)
                matrix[i, 0] = BoundCellText(list[i]);
            return matrix;
        }

        /// <summary>
        /// An N×C <see cref="object[,]"/> for Excel from ragged rows: each input row
        /// becomes a matrix row, shorter rows are padded with null (a blank cell) and
        /// the column count is the widest row (at least 1). Every string cell is
        /// bounded by <see cref="BoundCellText"/>. Backs stream, consumer and object
        /// listings. A null or empty input returns a 1×1 matrix containing
        /// <see cref="EmptyCell"/>.
        /// </summary>
        public static object[,] FormatRowsMatrix(IEnumerable<string[]> rows)
        {
            var list = rows == null ? new List<string[]>() : new List<string[]>(rows);
            if (list.Count == 0)
                return SingleCell(EmptyCell);

            int columns = 1;
            for (int i = 0; i < list.Count; i++)
            {
                var row = list[i];
                if (row != null && row.Length > columns)
                    columns = row.Length;
            }

            var matrix = new object[list.Count, columns];
            for (int i = 0; i < list.Count; i++)
            {
                var row = list[i];
                for (int j = 0; j < columns; j++)
                    matrix[i, j] = (row != null && j < row.Length) ? BoundCellText(row[j]) : null;
            }

            return matrix;
        }

        /// <summary>
        /// The KV history as a fixed N×4 <see cref="object[,]"/> (DESIGN.md §7.4,
        /// PLAN.md S5.19 history is N×4). Column order per row is
        /// <c>[0] revision</c>, <c>[1] operation</c>, <c>[2] timestamp</c>,
        /// <c>[3] value</c>. Each input row is an <see cref="object"/>[] of plain
        /// values: revision (long), operation (string), timestamp (DateTime or null)
        /// and value (string). Rows shorter than 4 are padded with null, rows longer
        /// than 4 are truncated to the first 4; string cells are bounded by
        /// <see cref="BoundCellText"/> and non-string cells (a boxed DateTime or
        /// number) pass through so Excel keeps its native date/number formatting.
        /// A null or empty input returns a 1×1 matrix containing
        /// <see cref="EmptyCell"/>.
        /// </summary>
        public static object[,] FormatKvHistoryMatrix(IEnumerable<object[]> rows)
        {
            const int columns = 4;
            var list = rows == null ? new List<object[]>() : new List<object[]>(rows);
            if (list.Count == 0)
                return SingleCell(EmptyCell);

            var matrix = new object[list.Count, columns];
            for (int i = 0; i < list.Count; i++)
            {
                var row = list[i];
                for (int j = 0; j < columns; j++)
                {
                    object cell = (row != null && j < row.Length) ? row[j] : null;
                    var text = cell as string;
                    matrix[i, j] = text != null ? BoundCellText(text) : cell;
                }
            }

            return matrix;
        }

        // ------------------------------------------------------------------
        // Internals
        // ------------------------------------------------------------------

        /// <summary>A 1×1 matrix holding a single value.</summary>
        private static object[,] SingleCell(object value)
        {
            var matrix = new object[1, 1];
            matrix[0, 0] = value;
            return matrix;
        }

        /// <summary>
        /// Serializes alternating name/value pairs as one compact JSON object with
        /// the fields in the order given. JObject preserves insertion order, which
        /// is what makes the output deterministic (a Dictionary would not guarantee
        /// it, and reflection over an anonymous type is not a contract).
        /// </summary>
        private static string OrderedJson(params object[] nameValuePairs)
        {
            if (nameValuePairs == null || (nameValuePairs.Length & 1) != 0)
                throw new ArgumentException("An even number of name/value arguments is required.", nameof(nameValuePairs));

            var obj = new JObject();
            for (int i = 0; i < nameValuePairs.Length; i += 2)
                obj.Add((string)nameValuePairs[i], ToToken(nameValuePairs[i + 1]));
            return JsonConvert.SerializeObject(obj, CompactSettings);
        }

        /// <summary>
        /// Converts a plain CLR value to a JSON token without ever throwing:
        /// null → null, string/bool/number/DateTime → scalar,
        /// <see cref="IEnumerable{T}"/> of string → array (null/empty → []), and any
        /// other value falls back to its invariant text so a formatter can never
        /// fail a formula on an unexpected type.
        /// </summary>
        private static JToken ToToken(object value)
        {
            if (value == null)
                return JValue.CreateNull();
            if (value is string s)
                return new JValue(s);
            if (value is bool b)
                return new JValue(b);
            if (value is long l)
                return new JValue(l);
            if (value is int i)
                return new JValue(i);
            if (value is DateTime dt)
            {
                // Normalize the Kind explicitly: the serializer's
                // DateTimeZoneHandling setting is not forwarded to the writer for a
                // JValue, so a Local DateTime could otherwise serialize with a
                // machine-dependent offset. Unspecified is treated as UTC (not as
                // local time) so the same logical instant is stable everywhere.
                var utc = dt.Kind == DateTimeKind.Unspecified
                    ? DateTime.SpecifyKind(dt, DateTimeKind.Utc)
                    : dt.ToUniversalTime();
                return new JValue(utc);
            }
            if (value is ulong ul)
                return new JValue(ul);
            if (value is IEnumerable<string> strings)
            {
                var array = new JArray();
                foreach (var item in strings)
                    array.Add(item == null ? (JToken)JValue.CreateNull() : new JValue(item));
                return array;
            }

            try
            {
                return new JValue(Convert.ToString(value, CultureInfo.InvariantCulture));
            }
            catch (Exception)
            {
                return JValue.CreateNull();
            }
        }
    }
}
