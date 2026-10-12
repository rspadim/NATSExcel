using ExcelDna.Integration;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.JetStream.Models;
using NATS.Client.ObjectStore;
using NATS.Client.ObjectStore.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// Object Store worksheet functions (DESIGN.md Â§7.5) with the additive
    /// <c>...NonVolatile</c> twin of every write (DESIGN.md Â§7.7).
    ///
    /// <para><b>No implicit creation.</b> An Object Store bucket is a JetStream
    /// stream named <c>OBJ_&lt;bucket&gt;</c>. Every bucket-scoped function opens
    /// the store with <c>GetObjectStoreAsync</c> and never with
    /// <c>CreateObjectStoreAsync</c>, so a typo in a bucket name is a named
    /// <c>bucket 'x' not found</c> error instead of a silently created bucket: a
    /// worksheet must not mutate the server's topology.</para>
    ///
    /// <para><b>Gates.</b> Every call resolves the optional <c>[host]</c>
    /// argument first, then applies the admin gate
    /// (<c>ObjectDelete</c> only, after argument validation and endpoint
    /// resolution), then the JetStream gate: the config flag
    /// (<see cref="AppConfig.JetStreamEnabled"/>) and the live per-endpoint probe
    /// (<see cref="NatsConnectionManager.IsJetStreamEnabled"/>). A server started
    /// without <c>-js</c> gets the exact
    /// "JetStream is not enabled on this server" text - never a hang and never an
    /// empty result that looks healthy.</para>
    ///
    /// <para><b>Reads and writes.</b> Reads are volatile and each one runs its
    /// whole body - store open plus data drain - inside a single task whose wait
    /// is bounded by <see cref="NatsConnectionManager.WaitBoundedResult{T}"/>.
    /// Every write goes through <see cref="NatsUdfAsync.Run"/> (the AsyncWrites
    /// dispatch). <c>ObjectPut</c> reports the server-side digest, so it is
    /// reply-dependent and the default <c>fireforget</c> mode still awaits it;
    /// only <c>fireforget-all</c> turns it into a fire-and-forget marker
    /// (DESIGN.md Â§8.1).</para>
    ///
    /// <para><b>Object text in a cell (DESIGN.md open decision 3).</b>
    /// <c>ObjectGet</c> refuses an object larger than
    /// <see cref="NatsUdfCommon.MaxObjectCellBytes"/> (1 MiB) and refuses a binary
    /// body (a NUL byte or an invalid UTF-8 sequence) with a named, masked error
    /// that states the size and the limit: a binary object is never rendered as
    /// replacement-character garbage. A body that passes both checks is then
    /// bounded to the Excel cell limit
    /// (<see cref="NatsResultFormatter.BoundCellText"/>), because a 1 MiB string
    /// still cannot live in a 32,767-character cell.</para>
    ///
    /// <para><b>Package note.</b> This file needs
    /// <c>NATS.Client.ObjectStore</c> on <c>NATSExcel.csproj</c> (plus a
    /// <c>&lt;Reference Pack="true"&gt;</c> for <c>NATS.Client.ObjectStore.dll</c>
    /// in <c>NATSExcel.dna</c>); the object-store types are not part of
    /// Core/JetStream/KeyValueStore. Until both are added the add-in project does
    /// not compile.</para>
    ///
    /// <para><b>Live watch.</b> <c>ObjectWatch</c> is a one-shot snapshot. The
    /// LIVE, push-on-change path is the RTD <c>OBJ</c> topic
    /// (<c>=RTD("NatsRtd",,"OBJ",bucket,object)</c>), which never downloads the
    /// body - it watches metadata only.</para>
    /// </summary>
    public static class NatsUdfObjectStore
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // An Object Store call is one (or a few) JetStream API round trip(s); the
        // blocking bound is never shorter than the configured response budget and
        // never longer than this, so a formula cannot hold the Excel thread for
        // minutes.
        private const int MaxObjectTimeoutMs = 60000;

        // A listing/watch drains a server stream that may legitimately carry more
        // metadata than one round trip; this only widens the wait, never shortens
        // the configured budget.
        private const int MinDrainTimeoutMs = 10000;

        // Defensive cap on what a drain accumulates on the Excel thread's behalf.
        private const int MaxObjectRows = 10000;

        // Strict UTF-8: an invalid byte sequence throws instead of becoming U+FFFD,
        // so a binary object is detected instead of rendered as garbage.
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

        // ==================================================================
        // Â§7.5 - reads (volatile, twin-less)
        // ==================================================================

        /// <summary>
        /// Returns the object body as text. Refuses (never truncates into
        /// garbage) an object larger than the 1 MiB cell bound or one whose bytes
        /// are not valid text; a missing object is <c>(no value)</c>.
        /// </summary>
        [ExcelFunction(Description = "Returns an Object Store object as text, or (no value) when it is missing", IsVolatile = true)]
        public static object NatsUDFObjectGet(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireObjectName(name);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadGet(endpoint, TextOf(bucket), TextOf(name));
        }

        /// <summary>
        /// The object's metadata as compact JSON
        /// (<c>name</c>, <c>digest</c>, <c>size</c>, <c>chunks</c>,
        /// <c>description</c>); the body is never read.
        /// </summary>
        [ExcelFunction(Description = "Returns Object Store object metadata as JSON", IsVolatile = true)]
        public static object NatsUDFObjectInfo(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireObjectName(name);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadInfo(endpoint, TextOf(bucket), TextOf(name));
        }

        /// <summary>
        /// The bucket's objects as an NÃ—4 matrix, ordinal-sorted by name:
        /// <c>[0] name</c>, <c>[1] size</c> (bytes),
        /// <c>[2] digest</c> (<c>SHA-256=...</c>), <c>[3] description</c>.
        /// An empty bucket returns the <c>(empty)</c> single cell, never a blank
        /// range.
        /// </summary>
        [ExcelFunction(Description = "Lists an Object Store bucket's objects as an Nx4 matrix", IsVolatile = true)]
        public static object NatsUDFObjectList(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadSnapshot(endpoint, TextOf(bucket), watch: false);
        }

        /// <summary>
        /// A one-shot metadata snapshot as an NÃ—4 matrix with the same columns as
        /// <see cref="NatsUDFObjectList"/>: <c>[0] name</c>, <c>[1] size</c>,
        /// <c>[2] digest</c>, <c>[3] description</c>. The LIVE, push-on-change
        /// path is the RTD <c>OBJ</c> topic (<c>=RTD("NatsRtd",,"OBJ",bucket,object)</c>);
        /// this function only ever takes a single snapshot, so a cell that must
        /// follow changes belongs in an RTD formula, not here.
        /// </summary>
        [ExcelFunction(Description = "Takes a one-shot Object Store metadata snapshot as an Nx4 matrix (live changes: RTD OBJ)", IsVolatile = true)]
        public static object NatsUDFObjectWatch(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadSnapshot(endpoint, TextOf(bucket), watch: true);
        }

        /// <summary>
        /// Lists the Object Store buckets on the endpoint (the <c>OBJ_</c>-prefixed
        /// JetStream streams) as an NÃ—1 matrix, ordinal-sorted. Read-only: it never
        /// creates a bucket.
        /// </summary>
        [ExcelFunction(Description = "Lists the Object Store buckets", IsVolatile = true)]
        public static object NatsUDFObjectBuckets(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            if (!TryPrepare(host, admin: false, out string endpoint, out string error))
                return error;

            return ReadBuckets(endpoint);
        }

        // ==================================================================
        // Â§7.5 - writes (+ ...NonVolatile twins)
        // ==================================================================

        /// <summary>
        /// Stores/overwrites an object and returns the server-computed
        /// <c>SHA-256=...</c> digest. A missing value cell is a legitimate empty
        /// object body.
        /// </summary>
        [ExcelFunction(Description = "Stores an Object Store object and returns its digest", IsVolatile = true)]
        public static object NatsUDFObjectPut(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "value", Description = "Object content text")] object value,
            [ExcelArgument(Name = "description", Description = "Optional description")] object description = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireObjectName(name);
            if (error != null)
                return error;
            if (!TryPayload(value, out string payload, out error))
                return error;
            if (!TryDescription(description, out string descriptionText, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFObjectPut", host,
                new object[] { bucket, name, value, description, host },
                () => WritePut(endpoint, bucketText, nameText, payload, descriptionText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFObjectPut"/>.</summary>
        [ExcelFunction(Description = "Stores an Object Store object and returns its digest", IsVolatile = false)]
        public static object NatsUDFObjectPutNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "value", Description = "Object content text")] object value,
            [ExcelArgument(Name = "description", Description = "Optional description")] object description = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFObjectPut(bucket, name, value, description, host);
        }

        /// <summary>
        /// Deletes an object (writes the deleted metadata marker and purges its
        /// chunks). Destructive; gated by <see cref="AppConfig.AllowAdminWrites"/>.
        /// </summary>
        [ExcelFunction(Description = "Deletes an Object Store object", IsVolatile = true)]
        public static object NatsUDFObjectDelete(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireBucket(bucket);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireObjectName(name);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string bucketText = TextOf(bucket);
            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFObjectDelete", host,
                new object[] { bucket, name, host },
                () => WriteDelete(endpoint, bucketText, nameText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFObjectDelete"/>.</summary>
        [ExcelFunction(Description = "Deletes an Object Store object", IsVolatile = false)]
        public static object NatsUDFObjectDeleteNonVolatile(
            [ExcelArgument(Name = "bucket", Description = "Object Store bucket name")] object bucket,
            [ExcelArgument(Name = "name", Description = "Object name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFObjectDelete(bucket, name, host);
        }

        // ==================================================================
        // Read implementations (one bounded wait per call)
        // ==================================================================

        private static object ReadGet(string endpoint, string bucket, string name)
        {
            return ReadObject(endpoint, "object get", bucket, async (store, token) =>
            {
                // Read the metadata first so an oversized body is refused BEFORE it
                // is downloaded: pulling 500 MB over the wire only to reject it
                // would hold the Excel thread for the whole transfer.
                ObjectMetadata info = await store.GetInfoAsync(name, cancellationToken: token).ConfigureAwait(false);
                if (info.Size > (ulong)NatsUdfCommon.MaxObjectCellBytes)
                    return OversizeError(name, info.Size);

                byte[] data = await store.GetBytesAsync(name, token).ConfigureAwait(false);
                if (data == null || data.Length == 0)
                    return NatsUdfCommon.NullValue;
                if (data.Length > NatsUdfCommon.MaxObjectCellBytes)
                    return OversizeError(name, (ulong)data.Length);

                string text;
                try
                {
                    text = StrictUtf8.GetString(data);
                }
                catch (DecoderFallbackException)
                {
                    return BinaryError(name);
                }

                // A NUL is valid UTF-8 but cannot live in a worksheet cell; it is
                // the same "this is binary, not text" case as an invalid sequence.
                if (text.IndexOf('\0') >= 0)
                    return BinaryError(name);

                return NatsResultFormatter.BoundCellText(text);
            });
        }

        private static object ReadInfo(string endpoint, string bucket, string name)
        {
            return ReadObject(endpoint, "object info", bucket, async (store, token) =>
            {
                ObjectMetadata info = await store.GetInfoAsync(name, cancellationToken: token).ConfigureAwait(false);
                return (object)NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatObjectInfoJson(
                    NameOr(info.Name, name),
                    info.Digest,
                    ToLong(info.Size),
                    ToInt(info.Chunks),
                    info.Description));
            });
        }

        /// <summary>
        /// The NÃ—4 metadata matrix: <c>[0] name</c>, <c>[1] size</c>,
        /// <c>[2] digest</c>, <c>[3] description</c>. <paramref name="watch"/>
        /// drains the watch channel (metadata only); otherwise the plain listing
        /// is used. Both are one-shot snapshots.
        /// </summary>
        private static object ReadSnapshot(string endpoint, string bucket, bool watch)
        {
            return ReadObject(endpoint, watch ? "object watch" : "object list", bucket, async (store, token) =>
            {
                var rows = new List<string[]>();
                IAsyncEnumerable<ObjectMetadata> source = watch
                    ? store.WatchAsync(new NatsObjWatchOpts { InitialSetOnly = true, IgnoreDeletes = true }, token)
                    : store.ListAsync(new NatsObjListOpts(), token);

                await foreach (ObjectMetadata item in source.ConfigureAwait(false))
                {
                    rows.Add(Row(item));
                    if (rows.Count >= MaxObjectRows)
                        break;
                }

                rows.Sort(CompareNames);
                return (object)NatsResultFormatter.FormatRowsMatrix(rows);
            }, listing: true);
        }

        private static object ReadBuckets(string endpoint)
        {
            // There is no bucket listing on the object-store client: a bucket IS a
            // JetStream stream named OBJ_<bucket>, so the stream names are the
            // source of truth and the prefix is stripped for the cell.
            if (!TryGetObjectContext(endpoint, out INatsObjContext context, out string error))
                return error;

            int bound = DrainTimeoutMs();
            var cts = new CancellationTokenSource();
            try
            {
                Task<object> task = Task.Run(async () =>
                {
                    var names = new List<string>();
                    await foreach (string streamName in
                        context.JetStreamContext.ListStreamNamesAsync(cancellationToken: cts.Token).ConfigureAwait(false))
                    {
                        if (streamName != null && streamName.StartsWith(StreamPrefix, StringComparison.Ordinal))
                        {
                            names.Add(streamName.Substring(StreamPrefix.Length));
                            if (names.Count >= MaxObjectRows)
                                break;
                        }
                    }

                    names.Sort(StringComparer.Ordinal);
                    return (object)NatsResultFormatter.FormatStringMatrix(names);
                });
                return NatsConnectionManager.WaitBoundedResult(task, bound);
            }
            catch (TimeoutException)
            {
                CancelQuietly(cts);
                return MapObjError("object buckets", new ObjTimeoutException("object buckets", bound), null, null);
            }
            catch (Exception ex)
            {
                CancelQuietly(cts);
                return MapObjError("object buckets", ex, null, null);
            }
        }

        /// <summary>
        /// Shared read body. The whole read - store open included - runs inside one
        /// task waited on with a single bounded wait, so the Excel thread is held
        /// for at most the drain budget regardless of how much data the drain
        /// touches; the drain is cancelled when that budget elapses. A missing
        /// bucket therefore also fails inside the bound, as the named bucket error.
        /// </summary>
        private static object ReadObject(
            string endpoint,
            string what,
            string bucket,
            Func<INatsObjStore, CancellationToken, Task<object>> body,
            bool listing = false)
        {
            if (!TryOpenStore(endpoint, bucket, out INatsObjStore store, out string error))
                return error;

            int bound = listing ? DrainTimeoutMs() : ObjectTimeoutMs();
            var cts = new CancellationTokenSource();
            try
            {
                Task<object> task = Task.Run(() => body(store, cts.Token));
                return NatsConnectionManager.WaitBoundedResult(task, bound);
            }
            catch (TimeoutException)
            {
                CancelQuietly(cts);
                return MapObjError(what, new ObjTimeoutException(what, bound), bucket, null);
            }
            catch (Exception ex)
            {
                CancelQuietly(cts);
                return MapObjError(what, ex, bucket, null);
            }
        }

        /// <summary>
        /// Opens an existing bucket. <c>GetObjectStoreAsync</c> is the "must
        /// exist" open: a missing bucket fails here and is mapped to the named
        /// bucket error. <c>CreateObjectStoreAsync</c> is deliberately never
        /// called, so a formula can never create a bucket by accident.
        /// </summary>
        private static bool TryOpenStore(string endpoint, string bucket, out INatsObjStore store, out string error)
        {
            store = null;
            INatsObjContext context;
            if (!TryGetObjectContext(endpoint, out context, out error))
                return false;

            try
            {
                store = WaitValue(context.GetObjectStoreAsync(bucket, CancellationToken.None), "object store open");
                return true;
            }
            catch (Exception ex)
            {
                error = MapObjError("object store", ex, bucket, null);
                return false;
            }
        }

        /// <summary>
        /// The Object Store context for an endpoint. It is a thin wrapper over the
        /// cached JetStream context (no extra socket), so it is rebuilt per call
        /// instead of being cached separately from the connection it points at.
        /// </summary>
        private static bool TryGetObjectContext(string endpoint, out INatsObjContext context, out string error)
        {
            context = null;
            error = null;
            try
            {
                INatsJSContext jetStream = NatsRuntime.Connections.GetJetStream(endpoint);
                if (jetStream == null)
                {
                    error = NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
                    return false;
                }

                context = new NatsObjContext(jetStream);
                return true;
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("Object Store", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }
        }

        // ==================================================================
        // Write implementations (every caller runs these through NatsUdfAsync)
        // ==================================================================

        private static object WritePut(string endpoint, string bucket, string name, string payload, string description)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFObjectPut", endpoint, bucket,
                    store => PutAsync(store, name, payload, description, CancellationToken.None));
            }

            if (!TryOpenStore(endpoint, bucket, out INatsObjStore store, out string error))
                return error;

            try
            {
                ObjectMetadata meta = WaitTask(PutAsync(store, name, payload, description, CancellationToken.None), "object put");
                return NatsResultFormatter.BoundCellText(meta?.Digest ?? string.Empty);
            }
            catch (Exception ex)
            {
                return MapObjError("object put", ex, bucket, name);
            }
        }

        private static object WriteDelete(string endpoint, string bucket, string name)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                return FireAndForget("NatsUDFObjectDelete", endpoint, bucket,
                    store => store.DeleteAsync(name, CancellationToken.None).AsTask());
            }

            if (!TryOpenStore(endpoint, bucket, out INatsObjStore store, out string error))
                return error;

            try
            {
                WaitVoid(store.DeleteAsync(name, CancellationToken.None), "object delete");
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapObjError("object delete", ex, bucket, name);
            }
        }

        /// <summary>
        /// The object-store put. A blank description uses the byte[] overload
        /// (which has no description field); a description goes through the
        /// metadata overload so the server records it alongside the body.
        /// </summary>
        private static Task<ObjectMetadata> PutAsync(
            INatsObjStore store, string name, string payload, string description, CancellationToken token)
        {
            byte[] data = payload == null ? Array.Empty<byte>() : Encoding.UTF8.GetBytes(payload);
            if (string.IsNullOrEmpty(description))
                return store.PutAsync(name, data, token).AsTask();

            var meta = new ObjectMetadata { Name = name, Description = description };
            return store.PutAsync(meta, new MemoryStream(data), leaveOpen: false, cancellationToken: token).AsTask();
        }

        // ==================================================================
        // Gates, bounded waits and error shaping
        // ==================================================================

        /// <summary>
        /// Shared prologue: resolve the optional <c>[host]</c> to a concrete
        /// endpoint, apply the admin gate (after argument validation and endpoint
        /// resolution, DESIGN.md Â§7.3), apply the JetStream gate and probe the
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
                // connecting to localhost (DESIGN.md Â§8.1).
                error = IsBlank(host) || host is string
                    ? NatsUdfCommon.ConfigError
                    : NatsUdfCommon.Fail("the host must be text");
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
                    NatsConnectionManager.WaitBounded(connection.PingAsync(CancellationToken.None).AsTask(), ObjectTimeoutMs());
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("connection", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (!NatsRuntime.Connections.IsJetStreamEnabled(endpoint))
            {
                // Exact frozen text, not a hang and not an empty result.
                error = NatsUdfCommon.RequireJetStream(false);
                return false;
            }

            return true;
        }

        /// <summary>Bounded wait for one Object Store round trip.</summary>
        private static T WaitValue<T>(ValueTask<T> task, string what)
        {
            try
            {
                return NatsConnectionManager.WaitBoundedResult(task.AsTask(), ObjectTimeoutMs());
            }
            catch (TimeoutException)
            {
                throw new ObjTimeoutException(what, ObjectTimeoutMs());
            }
        }

        /// <summary>The same bound for a Task<T> (an API that already returns a Task).</summary>
        private static T WaitTask<T>(Task<T> task, string what)
        {
            try
            {
                return NatsConnectionManager.WaitBoundedResult(task, ObjectTimeoutMs());
            }
            catch (TimeoutException)
            {
                throw new ObjTimeoutException(what, ObjectTimeoutMs());
            }
        }

        /// <summary>The same bound for a ValueTask with no result.</summary>
        private static void WaitVoid(ValueTask task, string what)
        {
            try
            {
                NatsConnectionManager.WaitBounded(task.AsTask(), ObjectTimeoutMs());
            }
            catch (TimeoutException)
            {
                throw new ObjTimeoutException(what, ObjectTimeoutMs());
            }
        }

        /// <summary>The blocking bound for one Object Store round trip.</summary>
        private static int ObjectTimeoutMs()
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
            if (configured > MaxObjectTimeoutMs)
                configured = MaxObjectTimeoutMs;
            return configured;
        }

        /// <summary>
        /// The blocking bound for a metadata drain (list/watch/buckets): a stream
        /// that legitimately carries more metadata than one round trip gets a wider
        /// window, never a shorter one than the configured budget.
        /// </summary>
        private static int DrainTimeoutMs()
        {
            int configured = ObjectTimeoutMs();
            return configured < MinDrainTimeoutMs ? MinDrainTimeoutMs : configured;
        }

        /// <summary>
        /// Fire-and-forget for a reply-dependent write forced by
        /// <c>fireforget-all</c>: the bucket is opened and the operation runs
        /// entirely in the background, and any failure is logged (masked) because
        /// there is no cell left to report it to.
        /// </summary>
        private static object FireAndForget(string what, string endpoint, string bucket, Func<INatsObjStore, Task> operation)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    INatsJSContext jetStream = NatsRuntime.Connections.GetJetStream(endpoint);
                    if (jetStream == null)
                        return;
                    INatsObjStore store = await new NatsObjContext(jetStream).GetObjectStoreAsync(bucket).ConfigureAwait(false);
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
        /// The named "not text" error for a body past the cell bound: it states
        /// the actual size and the limit, so a refuse is diagnosable from the cell.
        /// </summary>
        private static string OversizeError(string name, ulong size)
        {
            return NatsUdfCommon.VerboseError(
                "object get",
                "object '" + (name ?? string.Empty) + "' is " + size.ToString(CultureInfo.InvariantCulture)
                    + " bytes, above the " + NatsUdfCommon.MaxObjectCellBytes.ToString(CultureInfo.InvariantCulture)
                    + "-byte cell limit",
                "read it with a program, or use the RTD OBJ topic for metadata-only monitoring");
        }

        /// <summary>
        /// The named "binary" error for a body that carries a NUL byte or an
        /// invalid UTF-8 sequence: it is never rendered as replacement-character
        /// garbage.
        /// </summary>
        private static string BinaryError(string name)
        {
            return NatsUdfCommon.VerboseError(
                "object get",
                "object '" + (name ?? string.Empty) + "' is binary (a NUL byte or invalid UTF-8), not text",
                "a binary object cannot be a worksheet cell; use the RTD OBJ topic for metadata-only monitoring");
        }

        /// <summary>
        /// Maps an Object Store failure to the frozen cell vocabulary. A missing
        /// bucket is the named <c>bucket 'x' not found</c> (never an implicit
        /// create), a missing object is the named <c>object 'x' not found</c>, a
        /// JetStream-disabled server keeps the exact
        /// "JetStream is not enabled on this server" text, and a timeout names its
        /// bound. Every cause is masked by <see cref="NatsUdfCommon.VerboseError"/>.
        /// </summary>
        private static string MapObjError(string what, Exception ex, string bucket, string name)
        {
            if (ex is ObjTimeoutException timeout)
            {
                string cause = "timed out after " + timeout.TimeoutMs.ToString(CultureInfo.InvariantCulture) + "ms";
                return NatsUdfCommon.VerboseError(timeout.What, cause, NatsUdfCommon.ErrorHint("timed out"));
            }

            if (TryNested(ex, out NatsObjNotFoundException _))
                return ObjectNotFoundError(what, name, bucket);

            if (TryApiError(ex, out ApiError api))
            {
                string cause = ApiCause(api);
                string lower = cause.ToLowerInvariant();

                if (lower.Contains("jetstream") && lower.Contains("not enabled"))
                    return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);

                if (api.Code == 404)
                {
                    // A missing object and a missing bucket both surface as a 404
                    // against the backing stream; the caller's identifiers tell
                    // them apart, and an object read is the more specific case.
                    if (name != null && lower.Contains("message"))
                        return ObjectNotFoundError(what, name, bucket);
                    if (bucket != null && (lower.Contains("stream not found") || lower.Contains("stream") || api.ErrCode == 10059))
                        return BucketNotFoundError(what, bucket);
                }

                return NatsUdfCommon.VerboseError(what, cause, NatsUdfCommon.ErrorHint(cause));
            }

            string message = NatsUdfAsync.SafeMessage(ex);
            string text = (message ?? string.Empty).ToLowerInvariant();
            if (text.Contains("jetstream") && text.Contains("not enabled"))
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
            if (name != null && text.Contains("object not found"))
                return ObjectNotFoundError(what, name, bucket);
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
                "check the bucket name (an Object Store bucket is a JetStream stream named OBJ_<bucket>); create it on the server, not from a cell");
        }

        /// <summary>The named missing-object error.</summary>
        private static string ObjectNotFoundError(string what, string name, string bucket)
        {
            return NatsUdfCommon.VerboseError(
                what,
                "object '" + (name ?? string.Empty) + "' not found"
                    + (bucket == null ? string.Empty : " in bucket '" + bucket + "'"),
                "check the object name (NatsUDFObjectList(bucket) lists them)");
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

        /// <summary>True when the exception chain carries an exception of type <typeparamref name="T"/>.</summary>
        private static bool TryNested<T>(Exception ex, out T found) where T : Exception
        {
            found = null;
            if (ex == null)
                return false;
            if (ex is T match)
            {
                found = match;
                return true;
            }
            if (ex is AggregateException aggregate)
            {
                foreach (Exception inner in aggregate.Flatten().InnerExceptions)
                {
                    if (TryNested(inner, out found))
                        return true;
                }
            }
            return TryNested(ex.InnerException, out found);
        }

        private static string ApiCause(ApiError error)
        {
            string description = error.Description;
            string code = "code " + error.Code.ToString(CultureInfo.InvariantCulture)
                + ", err_code " + error.ErrCode.ToString(CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(description) ? code : description + " (" + code + ")";
        }

        // ==================================================================
        // Argument parsing and small helpers
        // ==================================================================

        /// <summary>The <c>OBJ_</c> stream-name prefix an Object Store bucket maps to.</summary>
        private const string StreamPrefix = "OBJ_";

        /// <summary>
        /// The body text of a write: a truly missing value cell is a legitimate
        /// empty object, while a multi-cell range or an Excel error cell is the
        /// frozen terse argument error.
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

        /// <summary>An optional description: a missing cell is no description.</summary>
        private static bool TryDescription(object value, out string description, out string error)
        {
            error = null;
            if (IsBlank(value))
            {
                description = null;
                return true;
            }

            try
            {
                description = NatsUdfCommon.ToNatsString(value);
                return true;
            }
            catch (ArgumentException ex)
            {
                description = null;
                error = NatsUdfCommon.Fail(ex.Message);
                return false;
            }
        }

        /// <summary>One NÃ—4 metadata row: name, size, digest, description.</summary>
        private static string[] Row(ObjectMetadata item)
        {
            if (item == null)
                return new[] { string.Empty, string.Empty, string.Empty, string.Empty };
            return new[]
            {
                NameOr(item.Name, string.Empty),
                item.Size.ToString(CultureInfo.InvariantCulture),
                item.Digest ?? string.Empty,
                item.Description ?? string.Empty
            };
        }

        private static int CompareNames(string[] left, string[] right)
        {
            string a = left != null && left.Length > 0 ? left[0] : null;
            string b = right != null && right.Length > 0 ? right[0] : null;
            return string.CompareOrdinal(a, b);
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
                logger.Debug(ex, "NatsUdfObjectStore: cancelling the drain failed");
            }
        }

        private static string NameOr(string value, string fallback)
        {
            return string.IsNullOrEmpty(value) ? fallback : value;
        }

        private static long ToLong(ulong value)
        {
            return value > long.MaxValue ? long.MaxValue : (long)value;
        }

        private static int ToInt(uint value)
        {
            return value > int.MaxValue ? int.MaxValue : (int)value;
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
        private sealed class ObjTimeoutException : Exception
        {
            public ObjTimeoutException(string what, int timeoutMs)
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




