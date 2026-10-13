using ExcelDna.Integration;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.JetStream.Models;
using NATS.Client.KeyValueStore;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// Key/Value Store worksheet functions (DESIGN.md §7.4) with the additive
    /// <c>...NonVolatile</c> twin of every write (DESIGN.md §7.7).
    ///
    /// <para><b>No implicit creation.</b> A KV bucket is a JetStream stream named
    /// <c>KV_&lt;bucket&gt;</c>. Every bucket-scoped function opens the store with
    /// <c>GetStoreAsync</c> and never with <c>CreateStoreAsync</c>, so a typo in a
    /// bucket name is a named <c>bucket 'x' not found</c> error instead of a
    /// silently created bucket: a worksheet must not mutate the server's
    /// topology.</para>
    ///
    /// <para><b>Gates.</b> Every call resolves the optional <c>[host]</c>
    /// argument first, then applies the admin gate (<c>KVDelete</c>/<c>KVPurge</c>
    /// only, after argument validation and endpoint resolution), then the
    /// JetStream gate: the config flag (<see cref="AppConfig.JetStreamEnabled"/>)
    /// and the live per-endpoint probe
    /// (<see cref="NatsConnectionManager.IsJetStreamEnabled"/>). A server started
    /// without <c>-js</c> gets the exact
    /// "JetStream is not enabled on this server" text - never a hang and never an
    /// empty result that looks healthy.</para>
    ///
    /// <para><b>Reads and writes.</b> Reads are volatile and each one runs its
    /// whole body - store open plus data drain - inside a single task whose wait
    /// is bounded by <see cref="NatsConnectionManager.WaitBoundedResult{T}"/>;
    /// bounding the call once (instead of per round trip) keeps the Excel thread
    /// hold independent of the number of keys, and the drain is cancelled when the
    /// bound elapses. Every write goes through <see cref="NatsUdfAsync.Run"/>
    /// (the AsyncWrites dispatch). Every KV write reports a revision, a create/CAS
    /// outcome or a tombstone, so all of them are reply-dependent and the default
    /// <c>fireforget</c> mode still awaits the reply; only <c>fireforget-all</c>
    /// turns them into a fire-and-forget marker (DESIGN.md §8.1).</para>
    ///
    /// <para><b>Package note.</b> This file needs
    /// <c>NATS.Client.KeyValueStore</c> on <c>NATSExcel.csproj</c> (plus a
    /// <c>&lt;Reference Pack="true"&gt;</c> in <c>NATSExcel.dna</c>); the KV types
    /// are not part of Core/JetStream.</para>
    /// </summary>
    public static class NatsUdfKv
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // A KV call is one (or a few) JetStream request/reply round trip(s). The
        // blocking bound is never shorter than the configured response budget and
        // never longer than this, so a formula cannot hold the Excel thread for
        // minutes.
        private const int MaxKvTimeoutMs = 60000;

        // Listing keys, a history or the bucket names drains a server stream that
        // may legitimately carry more data than one round trip; this only widens
        // the wait, it never shortens the configured budget.
        private const int MinDrainTimeoutMs = 10000;

        // Defensive caps on what a drain accumulates on the Excel thread's behalf.
        private const int MaxHistoryRows = 10000;
        private const int MaxKeyRows = 100000;
        private const int MaxBucketRows = 10000;

        // ==================================================================
        // §7.4 - reads (volatile, twin-less)
        // ==================================================================

        /// <summary>
        /// Returns the value stored under <c>key</c>, <c>(no value)</c> for a
        /// missing (or deleted) key, and <c>(null)</c> for a stored empty payload:
        /// "key absent" and "key present with an empty value" are different facts
        /// and must not collapse into the same cell.
        /// </summary>
        [ExcelFunction(Description = "Returns a Key/Value entry, or (no value) when the key is missing", IsVolatile = true)]
        public static object NatsUDFKVGet(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to read")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadGet(endpoint, TextOf(bucket), TextOf(key));
        }

        /// <summary>
        /// Reads several keys: an N×1 matrix of values, or an N×2 matrix
        /// (<c>[0] key</c>, <c>[1] value</c>) when <c>[columns]</c> is TRUE. Row
        /// <c>i</c> corresponds to the request's <c>i</c>-th key (row-major, so a
        /// multi-column <c>keys</c> range keeps the spreadsheet's reading order).
        /// A missing key is <c>(no value)</c> and a stored empty payload is
        /// <c>(null)</c>.
        /// </summary>
        [ExcelFunction(Description = "Reads several Key/Value entries into a matrix", IsVolatile = true)]
        public static object NatsUDFKVGetMultiple(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "keys", Description = "Keys range, or a single key")] object keys,
            [ExcelArgument(Name = "columns", Description = "TRUE to include the key column")] object columns = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = TryReadKeys(keys, out List<string> keyList);
            if (error != null)
                return error;
            if (!TryReadFlag(columns, "columns", out bool withColumns, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadGetMultiple(endpoint, TextOf(bucket), keyList, withColumns);
        }

        /// <summary>
        /// Lists the bucket's keys as an N×1 matrix, ordinal-sorted (a stable cell
        /// beats a server-dependent order). <c>[pattern]</c> is a subject-style key
        /// filter (<c>a.*</c> = one token, <c>a.&gt;</c> = one or more); omitted,
        /// the whole bucket is listed. Tombstones are never listed.
        /// </summary>
        [ExcelFunction(Description = "Lists the keys in a Key/Value bucket", IsVolatile = true)]
        public static object NatsUDFKVKeys(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "pattern", Description = "Optional key filter (wildcards allowed)")] object pattern = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;

            string patternText = null;
            if (!IsBlank(pattern))
            {
                // A key filter is a subject filter over the keys, so wildcard
                // tokens are valid; a blank or non-text filter is not.
                error = NatsUdfCommon.RequireText(pattern, "pattern");
                if (error != null)
                    return error;
                patternText = TextOf(pattern);
            }

            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadKeyList(endpoint, TextOf(bucket), patternText);
        }

        /// <summary>
        /// One entry's status as compact JSON:
        /// <c>{"revision":N,"created":&lt;iso|null&gt;,"updated":&lt;iso|null&gt;,"operation":"PUT"}</c>.
        /// A deleted key still answers, naming the <c>DEL</c> tombstone and the
        /// revision that carries it.
        /// </summary>
        /// <remarks>
        /// A KV read surfaces a single timestamp per entry - the time of the
        /// revision it returned - so <c>created</c> carries that instant and
        /// <c>updated</c> is always null: the store keeps no separately retrievable
        /// "first created" fact, and inventing one would be a lie in a cell.
        /// </remarks>
        [ExcelFunction(Description = "Returns a Key/Value entry's status as JSON", IsVolatile = true)]
        public static object NatsUDFKVStatus(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to inspect")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadStatus(endpoint, TextOf(bucket), TextOf(key));
        }

        /// <summary>
        /// A key's revision history as an N×4 matrix, oldest first:
        /// <c>[0] revision</c>, <c>[1] operation</c> (<c>PUT</c>/<c>DEL</c>/
        /// <c>PURGE</c>), <c>[2] timestamp</c>, <c>[3] value</c> (<c>(null)</c> for
        /// a delete/purge marker). A key with no history returns the <c>(empty)</c>
        /// single cell, never a blank range.
        /// </summary>
        [ExcelFunction(Description = "Returns a Key/Value key's revision history as an Nx4 matrix", IsVolatile = true)]
        public static object NatsUDFKVHistory(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to inspect")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadHistory(endpoint, TextOf(bucket), TextOf(key));
        }

        /// <summary>
        /// Lists the KV buckets on the endpoint (the <c>KV_</c>-prefixed JetStream
        /// streams) as an N×1 matrix, ordinal-sorted. Read-only: it never creates
        /// a bucket.
        /// </summary>
        [ExcelFunction(Description = "Lists the Key/Value buckets", IsVolatile = true)]
        public static object NatsUDFKVBuckets(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            if (!TryPrepare(host, admin: false, out string endpoint, out string error))
                return error;

            return ReadBuckets(endpoint);
        }

        // ==================================================================
        // §7.4 - writes (+ ...NonVolatile twins)
        // ==================================================================

        /// <summary>
        /// Upserts a value and returns the stored revision as a number. A missing
        /// value cell is a legitimate empty payload.
        /// </summary>
        [ExcelFunction(Description = "Upserts a Key/Value entry and returns its revision", IsVolatile = true)]
        public static object NatsUDFKVPut(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to write")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPayload(value, out string payload, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVPut", host, new object[] { bucket, key, value, host },
                () => WritePut(endpoint, bucketText, keyText, payload));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVPut"/>.</summary>
        [ExcelFunction(Description = "Upserts a Key/Value entry and returns its revision", IsVolatile = false)]
        public static object NatsUDFKVPutNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to write")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVPut(bucket, key, value, host);
        }

        /// <summary>
        /// Upserts a range serialized as JSON (matrix → JSON, the
        /// <see cref="NatsJson.MatrixToJson"/> shape) and returns the revision.
        /// </summary>
        [ExcelFunction(Description = "Upserts a range as JSON into a Key/Value entry and returns its revision", IsVolatile = true)]
        public static object NatsUDFKVPutJSON(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to write")] object key,
            [ExcelArgument(Name = "range", Description = "Range to serialize as JSON")] object[,] range,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireRange(range, "range");
            if (error != null)
                return error;

            string json = NatsJson.MatrixToJson(range);
            if (json != null && json.StartsWith(NatsUdfCommon.ErrorPrefix, StringComparison.Ordinal))
                return json;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVPutJSON", host, new object[] { bucket, key, range, host },
                () => WritePut(endpoint, bucketText, keyText, json ?? string.Empty));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVPutJSON"/>.</summary>
        [ExcelFunction(Description = "Upserts a range as JSON into a Key/Value entry and returns its revision", IsVolatile = false)]
        public static object NatsUDFKVPutJSONNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to write")] object key,
            [ExcelArgument(Name = "range", Description = "Range to serialize as JSON")] object[,] range,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVPutJSON(bucket, key, range, host);
        }

        /// <summary>
        /// Creates the entry only when the key is absent (a compare-and-set
        /// against revision 0) and returns the new revision. An existing key is a
        /// named error, never an overwrite.
        /// </summary>
        [ExcelFunction(Description = "Creates a Key/Value entry only if the key is absent", IsVolatile = true)]
        public static object NatsUDFKVCreate(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to create")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPayload(value, out string payload, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVCreate", host, new object[] { bucket, key, value, host },
                () => WriteCreate(endpoint, bucketText, keyText, payload));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVCreate"/>.</summary>
        [ExcelFunction(Description = "Creates a Key/Value entry only if the key is absent", IsVolatile = false)]
        public static object NatsUDFKVCreateNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to create")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVCreate(bucket, key, value, host);
        }

        /// <summary>
        /// Compare-and-set: writes only when the key's current revision equals
        /// <c>revision</c>, and returns the new revision. A stale revision is the
        /// named <c>wrong last revision for key 'k' (expected N)</c> error, so two
        /// spreadsheets cannot silently overwrite each other.
        /// </summary>
        [ExcelFunction(Description = "Updates a Key/Value entry only if the revision matches", IsVolatile = true)]
        public static object NatsUDFKVUpdate(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to update")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "revision", Description = "Expected current revision")] object revision,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPayload(value, out string payload, out error))
                return error;
            if (!TryRevision(revision, out long expected, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVUpdate", host, new object[] { bucket, key, value, revision, host },
                () => WriteUpdate(endpoint, bucketText, keyText, payload, expected));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVUpdate"/>.</summary>
        [ExcelFunction(Description = "Updates a Key/Value entry only if the revision matches", IsVolatile = false)]
        public static object NatsUDFKVUpdateNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to update")] object key,
            [ExcelArgument(Name = "value", Description = "Value text")] object value,
            [ExcelArgument(Name = "revision", Description = "Expected current revision")] object revision,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVUpdate(bucket, key, value, revision, host);
        }

        /// <summary>
        /// Deletes a key by writing a <c>DEL</c> tombstone: the revision history
        /// keeps the marker and the key stops being listed or readable.
        /// Destructive; gated by <see cref="AppConfig.AllowAdminWrites"/>.
        /// </summary>
        [ExcelFunction(Description = "Deletes a Key/Value entry (tombstone)", IsVolatile = true)]
        public static object NatsUDFKVDelete(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to delete")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVDelete", host, new object[] { bucket, key, host },
                () => WriteDelete(endpoint, bucketText, keyText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVDelete"/>.</summary>
        [ExcelFunction(Description = "Deletes a Key/Value entry (tombstone)", IsVolatile = false)]
        public static object NatsUDFKVDeleteNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to delete")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVDelete(bucket, key, host);
        }

        /// <summary>
        /// Purges a key: a <c>PURGE</c> marker is written and every stored revision
        /// of that subject is removed. Destructive; gated by
        /// <see cref="AppConfig.AllowAdminWrites"/>.
        /// </summary>
        [ExcelFunction(Description = "Purges a Key/Value entry and its history", IsVolatile = true)]
        public static object NatsUDFKVPurge(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to purge")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireKey(key);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string keyText = TextOf(key);
            return NatsUdfAsync.Run("NatsUDFKVPurge", host, new object[] { bucket, key, host },
                () => WritePurge(endpoint, bucketText, keyText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFKVPurge"/>.</summary>
        [ExcelFunction(Description = "Purges a Key/Value entry and its history", IsVolatile = false)]
        public static object NatsUDFKVPurgeNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Key/Value bucket name")] object bucket,
            [ExcelArgument(Name = "key", Description = "Key to purge")] object key,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFKVPurge(bucket, key, host);
        }

        // ==================================================================
        // Read implementations (one bounded wait per call)
        // ==================================================================

        private static object ReadGet(string endpoint, string bucket, string key)
        {
            return ReadKv(endpoint, "KV get", bucket, async (kv, token) =>
            {
                INatsKVStore store = await OpenStoreAsync(kv, bucket, token).ConfigureAwait(false);
                // TryGetEntryAsync reports a missing key as a result, not as a
                // throw, so the marker is produced without an exception; a
                // deleted key comes back as NatsKVKeyDeletedException and is
                // mapped the same way by the shared error mapping.
                NatsResult<NatsKVEntry<byte[]>> result =
                    await store.TryGetEntryAsync<byte[]>(key, cancellationToken: token).ConfigureAwait(false);
                if (!result.Success)
                    throw result.Error;
                return (object)ValueText(result.Value.Value);
            });
        }

        private static object ReadGetMultiple(string endpoint, string bucket, List<string> keys, bool withColumns)
        {
            return ReadKv(endpoint, "KV get multiple", bucket, async (kv, token) =>
            {
                INatsKVStore store = await OpenStoreAsync(kv, bucket, token).ConfigureAwait(false);
                int columns = withColumns ? 2 : 1;
                var matrix = new object[keys.Count, columns];
                for (int i = 0; i < keys.Count; i++)
                {
                    string key = keys[i];
                    string cell;
                    NatsResult<NatsKVEntry<byte[]>> result =
                        await store.TryGetEntryAsync<byte[]>(key, cancellationToken: token).ConfigureAwait(false);
                    if (result.Success)
                    {
                        cell = ValueText(result.Value.Value);
                    }
                    else if (result.Error is NatsKVKeyNotFoundException || result.Error is NatsKVKeyDeletedException)
                    {
                        cell = NatsUdfCommon.NoValue;
                    }
                    else
                    {
                        // Anything else (connection, authorization, timeout) is
                        // systemic: fail the batch instead of hiding it in one
                        // cell of an otherwise healthy-looking matrix.
                        throw result.Error;
                    }

                    if (withColumns)
                        matrix[i, 0] = NatsResultFormatter.BoundCellText(key);
                    matrix[i, columns - 1] = cell;
                }

                return (object)matrix;
            });
        }

        private static object ReadKeyList(string endpoint, string bucket, string pattern)
        {
            return ReadKv(endpoint, "KV keys", bucket, async (kv, token) =>
            {
                INatsKVStore store = await OpenStoreAsync(kv, bucket, token).ConfigureAwait(false);
                var keys = new List<string>();
                IAsyncEnumerable<string> source = string.IsNullOrWhiteSpace(pattern)
                    ? store.GetKeysAsync(cancellationToken: token)
                    : store.GetKeysAsync(new[] { pattern }, cancellationToken: token);

                await foreach (string key in source.ConfigureAwait(false))
                {
                    keys.Add(key);
                    if (keys.Count >= MaxKeyRows)
                        break;
                }

                // Ordinal sort: a key listing has no meaningful revision order and
                // a stable cell is far more useful in a sheet.
                keys.Sort(StringComparer.Ordinal);
                return (object)NatsResultFormatter.FormatStringMatrix(keys);
            });
        }

        private static object ReadStatus(string endpoint, string bucket, string key)
        {
            return ReadKv(endpoint, "KV status", bucket, async (kv, token) =>
            {
                INatsKVStore store = await OpenStoreAsync(kv, bucket, token).ConfigureAwait(false);
                NatsResult<NatsKVEntry<byte[]>> result =
                    await store.TryGetEntryAsync<byte[]>(key, cancellationToken: token).ConfigureAwait(false);

                if (result.Success)
                {
                    NatsKVEntry<byte[]> entry = result.Value;
                    return (object)NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatKvStatusJson(
                        ToLong(entry.Revision), CreatedUtc(entry), null, OperationName(entry.Operation)));
                }

                // A deleted key still has a status: report the tombstone and the
                // revision that carries it instead of pretending the key never
                // existed.
                if (result.Error is NatsKVKeyDeletedException deleted)
                {
                    return (object)NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatKvStatusJson(
                        ToLong(deleted.Revision), null, null, "DEL"));
                }

                throw result.Error;
            });
        }

        private static object ReadHistory(string endpoint, string bucket, string key)
        {
            return ReadKv(endpoint, "KV history", bucket, async (kv, token) =>
            {
                INatsKVStore store = await OpenStoreAsync(kv, bucket, token).ConfigureAwait(false);
                var rows = new List<object[]>();
                // HistoryAsync enables IncludeHistory internally and ends when the
                // entry's Delta reaches 0 (its newest revision).
                await foreach (NatsKVEntry<byte[]> entry in
                    store.HistoryAsync<byte[]>(key, cancellationToken: token).ConfigureAwait(false))
                {
                    rows.Add(new object[]
                    {
                        ToLong(entry.Revision),
                        OperationName(entry.Operation),
                        CreatedUtc(entry),
                        ValueText(entry.Value)
                    });
                    if (rows.Count >= MaxHistoryRows)
                        break;
                }

                return (object)NatsResultFormatter.FormatKvHistoryMatrix(rows);
            });
        }

        private static object ReadBuckets(string endpoint)
        {
            // The bucket listing has no bucket: it drains the bucket names off the
            // context itself, so it must not open a store first.
            return ReadKv(endpoint, "KV buckets", null, async (kv, token) =>
            {
                var names = new List<string>();
                await foreach (string name in kv.GetBucketNamesAsync(token).ConfigureAwait(false))
                {
                    names.Add(name);
                    if (names.Count >= MaxBucketRows)
                        break;
                }

                names.Sort(StringComparer.Ordinal);
                return (object)NatsResultFormatter.FormatStringMatrix(names);
            }, listing: true);
        }

        /// <summary>
        /// Shared read body. The whole read - store open included - runs inside one
        /// task waited on with a single bounded wait, so the Excel thread is held
        /// for at most <paramref name="what"/>'s own budget regardless of how much
        /// data the drain touches; the drain is cancelled when that budget
        /// elapses. A missing bucket therefore also fails inside the bound, as the
        /// named bucket error.
        /// </summary>
        private static object ReadKv(
            string endpoint,
            string what,
            string bucket,
            Func<INatsKVContext, CancellationToken, Task<object>> body,
            bool listing = false)
        {
            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            int bound = listing ? DrainTimeoutMs() : KvTimeoutMs();
            var cts = new CancellationTokenSource();
            try
            {
                Task<object> task = Task.Run(() => body(kv, cts.Token));
                return NatsConnectionManager.WaitBoundedResult(task, bound);
            }
            catch (TimeoutException)
            {
                CancelQuietly(cts);
                return MapKvError(what, new KvTimeoutException(what, bound), bucket, null);
            }
            catch (Exception ex)
            {
                CancelQuietly(cts);
                return MapKvError(what, ex, bucket, null);
            }
        }

        /// <summary>
        /// Opens an existing bucket. <c>GetStoreAsync</c> is the "must exist"
        /// open: a missing bucket fails here and is mapped to the named bucket
        /// error. <c>CreateStoreAsync</c> is deliberately never called, so a
        /// formula can never create a bucket by accident (DESIGN.md §7.4).
        /// </summary>
        private static Task<INatsKVStore> OpenStoreAsync(INatsKVContext kv, string bucket, CancellationToken token)
        {
            return kv.GetStoreAsync(bucket, token).AsTask();
        }

        // ==================================================================
        // Write implementations (every caller runs these through NatsUdfAsync)
        // ==================================================================

        private static object WritePut(string endpoint, string bucket, string key, string payload)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFKVPut", endpoint, bucket,
                    store => store.PutAsync<string>(key, payload).AsTask());
            }

            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            try
            {
                INatsKVStore store = WaitResult(OpenStoreAsync(kv, bucket, CancellationToken.None), "KV put");
                ulong revision = WaitResult(store.PutAsync<string>(key, payload).AsTask(), "KV put");
                return (double)revision;
            }
            catch (Exception ex)
            {
                return MapKvError("KV put", ex, bucket, key);
            }
        }

        private static object WriteCreate(string endpoint, string bucket, string key, string payload)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFKVCreate", endpoint, bucket,
                    store => store.CreateAsync<string>(key, payload).AsTask());
            }

            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            try
            {
                INatsKVStore store = WaitResult(OpenStoreAsync(kv, bucket, CancellationToken.None), "KV create");
                NatsResult<ulong> created = WaitResult(
                    store.TryCreateAsync<string>(key, payload).AsTask(), "KV create");
                if (!created.Success)
                {
                    if (created.Error is NatsKVCreateException)
                    {
                        return NatsUdfCommon.VerboseError(
                            "KV create", KeyExists(key, bucket),
                            "the create is a compare-and-set against revision 0; use NatsUDFKVPut to overwrite or NatsUDFKVUpdate to CAS");
                    }

                    throw created.Error;
                }

                return (double)created.Value;
            }
            catch (Exception ex)
            {
                return MapKvError("KV create", ex, bucket, key);
            }
        }

        private static object WriteUpdate(string endpoint, string bucket, string key, string payload, long expectedRevision)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFKVUpdate", endpoint, bucket,
                    store => store.UpdateAsync<string>(key, payload, (ulong)expectedRevision).AsTask());
            }

            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            try
            {
                INatsKVStore store = WaitResult(OpenStoreAsync(kv, bucket, CancellationToken.None), "KV update");
                NatsResult<ulong> updated = WaitResult(
                    store.TryUpdateAsync<string>(key, payload, (ulong)expectedRevision).AsTask(), "KV update");
                if (!updated.Success)
                {
                    // The CAS precondition is the one failure worth naming
                    // precisely: the caller passed the revision it expected.
                    if (updated.Error is NatsKVWrongLastRevisionException)
                    {
                        string cause = NatsUdfCommon.WrongLastRevision(key, expectedRevision);
                        return NatsUdfCommon.VerboseError("KV update", cause, NatsUdfCommon.ErrorHint("wrong last revision"));
                    }

                    throw updated.Error;
                }

                return (double)updated.Value;
            }
            catch (Exception ex)
            {
                return MapKvError("KV update", ex, bucket, key);
            }
        }

        private static object WriteDelete(string endpoint, string bucket, string key)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                // DeleteAsync is the throwing form; under fireforget-all there is
                // no cell to report a refusal to, so the background task logs it.
                return FireAndForget("NatsUDFKVDelete", endpoint, bucket,
                    store => store.DeleteAsync(key).AsTask());
            }

            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            try
            {
                INatsKVStore store = WaitResult(OpenStoreAsync(kv, bucket, CancellationToken.None), "KV delete");
                NatsResult deleted = WaitResult(store.TryDeleteAsync(key).AsTask(), "KV delete");
                if (!deleted.Success)
                    throw deleted.Error;
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapKvError("KV delete", ex, bucket, key);
            }
        }

        private static object WritePurge(string endpoint, string bucket, string key)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFKVPurge", endpoint, bucket,
                    store => store.PurgeAsync(key).AsTask());
            }

            if (!TryGetKvContext(endpoint, out INatsKVContext kv, out string error))
                return error;

            try
            {
                INatsKVStore store = WaitResult(OpenStoreAsync(kv, bucket, CancellationToken.None), "KV purge");
                NatsResult purged = WaitResult(store.TryPurgeAsync(key).AsTask(), "KV purge");
                if (!purged.Success)
                    throw purged.Error;
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapKvError("KV purge", ex, bucket, key);
            }
        }

        // ==================================================================
        // Gates, bounded waits and error shaping
        // ==================================================================

        /// <summary>
        /// Shared prologue: resolve the optional <c>[host]</c> to a concrete
        /// endpoint, apply the admin gate (after argument validation and endpoint
        /// resolution, DESIGN.md §7.3), apply the JetStream gate and probe the
        /// server. False returns the exact cell text in <paramref name="error"/>.
        /// </summary>
        private static bool TryPrepare(object host, bool admin, out string endpoint, out string error)
        {
            endpoint = null;
            error = null;

            try
            {
                endpoint = NatsRuntime.ResolveEndpointForDispatch(host);
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("endpoint", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (endpoint == null)
            {
                // A non-text host is a bad argument; an unresolvable text host is
                // a dangling reference, reported as #CONFIG instead of silently
                // connecting to localhost (DESIGN.md §8.1).
                error = IsBlank(host) || host is string
                    ? NatsUdfCommon.ConfigError
                    : NatsUdfCommon.Fail(NatsUdfCommon.ErrorHostNotText);
                return false;
            }

            if (admin)
            {
                string gate = NatsUdfCommon.RequireAdminWrites(AppConfig.AllowAdminWrites);
                if (gate != null)
                {
                    error = gate;
                    return false;
                }
            }

            string jetStream = NatsUdfCommon.RequireJetStream(AppConfig.Current.JetStreamEnabled);
            if (jetStream != null)
            {
                error = jetStream;
                return false;
            }

            // Materialize the client (cheap; the socket opens behind us) so the
            // probe can read the server INFO, then refuse explicitly instead of
            // issuing a call that would hang until its request timeout. On a cold
            // first call the INFO has not arrived yet, so one bounded round trip
            // forces the connect and makes the probe truthful instead of a false
            // "not enabled" on a healthy server.
            try
            {
                NatsConnection connection = NatsRuntime.Connections.GetConnection(endpoint);
                if (connection.ServerInfo == null)
                    NatsConnectionManager.WaitBounded(connection.PingAsync().AsTask(), KvTimeoutMs());
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("connection", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (!NatsRuntime.Connections.IsJetStreamEnabled(endpoint))
            {
                // Two different causes must not share one message: a server that
                // answered but has no JetStream is a config problem, while an
                // endpoint that never answered is a connectivity problem. Saying
                // "JetStream is not enabled" for an unreachable host sends the
                // user looking in the wrong place.
                if (!NatsRuntime.Connections.HasConnection(endpoint))
                    error = NatsUdfCommon.VerboseError(
                        "connection", "no server answered at the endpoint",
                        "check that the server is running and that the URL/token are right");
                else
                    error = NatsUdfCommon.RequireJetStream(false); // exact frozen text
                return false;
            }

            return true;
        }

        /// <summary>
        /// The KV context for an endpoint. It is a thin wrapper over the cached
        /// JetStream context (no extra socket), so it is rebuilt per call instead
        /// of being cached separately from the connection it points at.
        /// </summary>
        private static bool TryGetKvContext(string endpoint, out INatsKVContext kv, out string error)
        {
            kv = null;
            error = null;
            try
            {
                INatsJSContext jetStream = NatsRuntime.Connections.GetJetStream(endpoint);
                if (jetStream == null)
                {
                    error = NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
                    return false;
                }

                kv = new NatsKVContext(jetStream);
                return true;
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("Key/Value", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }
        }

        /// <summary>
        /// Bounded wait for one KV round trip. A timeout becomes a named,
        /// hint-carrying cell error; any other failure goes through
        /// <see cref="MapKvError"/> at the call site.
        /// </summary>
        private static T WaitResult<T>(Task<T> task, string what)
        {
            try
            {
                return NatsConnectionManager.WaitBoundedResult(task, KvTimeoutMs());
            }
            catch (TimeoutException)
            {
                throw new KvTimeoutException(what, KvTimeoutMs());
            }
        }

        /// <summary>The blocking bound for one KV round trip.</summary>
        private static int KvTimeoutMs()
        {
            int configured;
            try
            {
                configured = AppConfig.Current.JetStreamTimeoutMs;
            }
            catch
            {
                configured = 0;
            }

            int floor = NatsConnectionManager.ResponseTimeoutMs(NatsPool.Udf);
            if (configured < floor)
                configured = floor;
            if (configured > MaxKvTimeoutMs)
                configured = MaxKvTimeoutMs;
            return configured;
        }

        /// <summary>
        /// The blocking bound for a listing drain (keys, history, buckets): a
        /// stream that legitimately carries more data than one round trip gets a
        /// wider window, never a shorter one than the configured budget.
        /// </summary>
        private static int DrainTimeoutMs()
        {
            int configured = KvTimeoutMs();
            return configured < MinDrainTimeoutMs ? MinDrainTimeoutMs : configured;
        }

        /// <summary>
        /// Fire-and-forget for a reply-dependent write forced by
        /// <c>fireforget-all</c>: the bucket is opened and the operation runs
        /// entirely in the background, and any failure is logged (with every
        /// credential masked by the log-site masking discipline) because there is
        /// no cell left to report it to.
        /// </summary>
        private static object FireAndForget(string what, string endpoint, string bucket, Func<INatsKVStore, Task> operation)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    INatsJSContext jetStream = NatsRuntime.Connections.GetJetStream(endpoint);
                    if (jetStream == null)
                        return;
                    INatsKVStore store = await new NatsKVContext(jetStream).GetStoreAsync(bucket).ConfigureAwait(false);
                    await operation(store).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    try
                    {
                        logger.Warn(what + ": fire-and-forget write failed: " + NatsUdfAsync.SafeMessage(ex));
                    }
                    catch
                    {
                        // Never let a logging failure escape a task continuation.
                    }
                }
            });

            return NatsUdfCommon.FireAndForgetMarker(AppConfig.SyncWrite);
        }

        /// <summary>
        /// Maps a KV failure to the frozen cell vocabulary. A missing bucket is the
        /// named <c>bucket 'x' not found</c> (never an implicit create), a missing
        /// or deleted key is <c>(no value)</c>, a stale CAS names the key and the
        /// revision the caller expected, and a JetStream-disabled server keeps the
        /// exact "JetStream is not enabled on this server" text. Every cause is
        /// masked by <see cref="NatsUdfCommon.VerboseError"/>.
        /// </summary>
        private static string MapKvError(string what, Exception ex, string bucket, string key)
        {
            if (ex is KvTimeoutException timeout)
            {
                string cause = "timed out after " + timeout.TimeoutMs.ToString(CultureInfo.InvariantCulture) + "ms";
                return NatsUdfCommon.VerboseError(timeout.What, cause, NatsUdfCommon.ErrorHint("timed out"));
            }

            if (ex is NatsKVCreateException)
                return NatsUdfCommon.VerboseError(what, KeyExists(key, bucket), null);

            if (ex is NatsKVWrongLastRevisionException)
            {
                // Only the call site knows the compared revision; without it the
                // library's own cause text is the honest one.
                return NatsUdfCommon.VerboseError(
                    what, "wrong last revision", NatsUdfCommon.ErrorHint("wrong last revision"));
            }

            if (ex is NatsKVKeyNotFoundException || ex is NatsKVKeyDeletedException)
                return NatsUdfCommon.NoValue;

            if (TryApiError(ex, out ApiError api))
            {
                string cause = ApiCause(api);
                string lower = cause.ToLowerInvariant();

                if (lower.Contains("jetstream") && lower.Contains("not enabled"))
                    return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);

                // A KV bucket lives in the "KV_<bucket>" stream, so the server's
                // 404/10059 names the stream, not the user's bucket label.
                if (api.Code == 404 || lower.Contains("stream not found"))
                {
                    if (bucket != null)
                        return BucketNotFoundError(what, bucket);
                    if (key != null)
                        return NatsUdfCommon.NoValue;
                }

                return NatsUdfCommon.VerboseError(what, cause, NatsUdfCommon.ErrorHint(cause));
            }

            string message = NatsUdfAsync.SafeMessage(ex);
            string text = (message ?? string.Empty).ToLowerInvariant();
            if (text.Contains("jetstream") && text.Contains("not enabled"))
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
            if (key != null && (text.Contains("key not found") || text.Contains("no message found")))
                return NatsUdfCommon.NoValue;
            if (bucket != null && (text.Contains("stream not found") || text.Contains("bucket not found")))
                return BucketNotFoundError(what, bucket);

            return NatsUdfCommon.VerboseError(what, message, NatsUdfCommon.ErrorHint(message));
        }

        /// <summary>The named missing-bucket error (never an implicit create).</summary>
        private static string BucketNotFoundError(string what, string bucket)
        {
            return NatsUdfCommon.VerboseError(
                what,
                "bucket '" + (bucket ?? string.Empty) + "' not found",
                "check the bucket name (a KV bucket is a JetStream stream named KV_<bucket>); create it on the server, not from a cell");
        }

        /// <summary>Finds the server's structured error in an (aggregated) exception chain.</summary>
        private static bool TryApiError(Exception ex, out ApiError api)
        {
            api = null;
            if (ex == null)
                return false;
            if (ex is NatsJSApiException apiException && apiException.Error != null)
            {
                api = apiException.Error;
                return true;
            }
            if (ex is AggregateException aggregate)
            {
                foreach (Exception inner in aggregate.Flatten().InnerExceptions)
                {
                    if (TryApiError(inner, out api))
                        return true;
                }
            }
            return TryApiError(ex.InnerException, out api);
        }

        private static string ApiCause(ApiError error)
        {
            string description = error.Description;
            string code = "code " + error.Code.ToString(CultureInfo.InvariantCulture)
                + ", err_code " + error.ErrCode.ToString(CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(description) ? code : description + " (" + code + ")";
        }

        // ==================================================================
        // Argument parsing
        // ==================================================================

        /// <summary>
        /// The payload text of a write: a truly missing value cell is a legitimate
        /// empty NATS payload, while a multi-cell range or an Excel error cell is
        /// the frozen terse argument error.
        /// </summary>
        private static bool TryPayload(object value, out string payload, out string error)
        {
            error = null;
            try
            {
                payload = NatsUdfCommon.ToNatsString(value) ?? string.Empty;
                return true;
            }
            catch (ArgumentException ex)
            {
                payload = null;
                error = NatsUdfCommon.Fail(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// A CAS revision: present, integral and positive (revision 0 means "must
        /// not exist" and belongs to <c>NatsUDFKVCreate</c>).
        /// </summary>
        private static bool TryRevision(object value, out long revision, out string error)
        {
            error = null;
            if (IsBlank(value))
            {
                revision = 0;
                error = NatsUdfCommon.Fail("a revision is required");
                return false;
            }

            if (!NatsUdfCommon.TryToLong(value, out revision) || revision < 1)
            {
                revision = 0;
                error = NatsUdfCommon.Fail("the revision must be a positive integer");
                return false;
            }

            return true;
        }

        /// <summary>An optional boolean argument (a cell TRUE/FALSE or its text).</summary>
        private static bool TryReadFlag(object value, string what, out bool flag, out string error)
        {
            flag = false;
            error = null;
            if (IsBlank(value))
                return true;
            if (value is bool boolean)
            {
                flag = boolean;
                return true;
            }
            if (value is ExcelError)
            {
                error = NatsUdfCommon.Fail(NatsUdfCommon.ErrorExcelErrorCell);
                return false;
            }

            string text = TextOf(value);
            if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase))
            {
                flag = true;
                return true;
            }
            if (string.Equals(text, "false", StringComparison.OrdinalIgnoreCase))
            {
                flag = false;
                return true;
            }

            error = NatsUdfCommon.Fail("the " + what + " must be TRUE or FALSE");
            return false;
        }

        /// <summary>
        /// Reads the <c>keys</c> argument: a range of any shape (row-major order,
        /// blank cells skipped) or a single key. An empty result is the frozen
        /// "no entries to write" error rather than a silently empty matrix.
        /// </summary>
        private static string TryReadKeys(object value, out List<string> keys)
        {
            keys = new List<string>();
            if (IsBlank(value))
                return NatsUdfCommon.Fail("a keys range is required");
            if (value is ExcelError)
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorExcelErrorCell);

            if (value is object[,] range)
            {
                string rangeError = NatsUdfCommon.RequireRange(range, "keys range");
                if (rangeError != null)
                    return rangeError;

                foreach (object cell in range)
                {
                    // Reject an Excel error cell explicitly: TextOf swallows it
                    // (ToNatsString throws and TextOf returns null), so a #N/A
                    // would be skipped like a blank and shift every later key
                    // onto the wrong row of the result. RequireRange already
                    // checks the range, but it is re-read here, so the per-cell
                    // guard is repeated (NatsUdfCommon.RequireRange / PLAN 1.3).
                    if (cell is ExcelError)
                        return NatsUdfCommon.Fail(NatsUdfCommon.ErrorExcelErrorCell);

                    string text = TextOf(cell);
                    if (!string.IsNullOrWhiteSpace(text))
                        keys.Add(text);
                }
            }
            else if (value is string)
            {
                string keyError = NatsUdfCommon.RequireKey(value);
                if (keyError != null)
                    return keyError;
                keys.Add(TextOf(value));
            }
            else
            {
                return NatsUdfCommon.Fail("a keys range is required");
            }

            if (keys.Count == 0)
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorNoEntries);
            return null;
        }

        // ==================================================================
        // Small helpers
        // ==================================================================

        /// <summary>
        /// The cell text of a stored payload: <c>(null)</c> for an empty payload
        /// (including a DEL/PURGE tombstone's absent value), otherwise the lenient
        /// UTF-8 decoding of the bytes, bounded to the Excel cell limit.
        /// Distinguishing it from <see cref="NatsUdfCommon.NoValue"/> is the point:
        /// "key absent" and "key present with an empty value" are different facts.
        /// </summary>
        private static string ValueText(byte[] data)
        {
            if (data == null || data.Length == 0)
                return NatsUdfCommon.NullValue;
            return NatsResultFormatter.BoundCellText(NatsResultFormatter.DecodePayload(data) ?? string.Empty);
        }

        private static string OperationName(NatsKVOperation operation)
        {
            switch (operation)
            {
                case NatsKVOperation.Put: return "PUT";
                case NatsKVOperation.Del: return "DEL";
                case NatsKVOperation.Purge: return "PURGE";
                default: return operation.ToString().ToUpperInvariant();
            }
        }

        /// <summary>The entry's timestamp as UTC, or null when the server sent none.</summary>
        private static DateTime? CreatedUtc(NatsKVEntry<byte[]> entry)
        {
            return entry.Created == default ? (DateTime?)null : entry.Created.UtcDateTime;
        }

        private static string KeyExists(string key, string bucket)
        {
            return "key '" + (key ?? string.Empty) + "' already exists in bucket '" + (bucket ?? string.Empty) + "'";
        }

        private static long ToLong(ulong value)
        {
            return value > long.MaxValue ? long.MaxValue : (long)value;
        }

        /// <summary>
        /// Cancels a drain that outlived its bound so the background enumerator
        /// stops instead of leaking a watcher. The source is deliberately not
        /// disposed: the cancelled work may still be unwinding and disposing would
        /// race an in-flight registration, while the source itself holds no timer
        /// and no unmanaged resource.
        /// </summary>
        private static void CancelQuietly(CancellationTokenSource cts)
        {
            try
            {
                cts.Cancel();
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "NatsUdfKv: cancelling the drain failed");
            }
        }

        /// <summary>The invariant text of a cell value, trimmed; null when it cannot be read.</summary>
        private static string TextOf(object value)
        {
            try
            {
                return NatsUdfCommon.ToNatsString(value)?.Trim();
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static bool IsBlank(object value)
        {
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                return true;
            return value is string text && string.IsNullOrWhiteSpace(text);
        }

        /// <summary>Internal timeout carrier so the mapping sees a named bound.</summary>
        private sealed class KvTimeoutException : Exception
        {
            public KvTimeoutException(string what, int timeoutMs)
                : base(what + " timed out after " + timeoutMs.ToString(CultureInfo.InvariantCulture) + "ms")
            {
                What = what;
                TimeoutMs = timeoutMs;
            }

            public string What { get; }

            public int TimeoutMs { get; }
        }
    }
}
