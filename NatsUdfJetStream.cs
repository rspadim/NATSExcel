using ExcelDna.Integration;
using NATS.Client.Core;
using NATS.Client.JetStream;
using NATS.Client.JetStream.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace NATSExcel
{
    /// <summary>
    /// JetStream worksheet functions: streams, messages and consumers
    /// (DESIGN.md §7.2) plus the opt-in admin surface (DESIGN.md §7.3), with the
    /// additive <c>...NonVolatile</c> twin of every write (DESIGN.md §7.7).
    ///
    /// <para><b>Gates.</b> Every call first resolves the optional <c>[host]</c>
    /// argument, then applies the admin gate (admin/destructive functions only,
    /// after argument validation and endpoint resolution, so a caller never sees a
    /// gate error for a call that was malformed anyway), then the JetStream gate:
    /// the config flag (<see cref="AppConfig.JetStreamEnabled"/>) and the live
    /// per-endpoint probe (<see cref="NatsConnectionManager.IsJetStreamEnabled"/>).
    /// A server started without <c>-js</c> gets the exact
    /// "JetStream is not enabled on this server" text - never a hang and never an
    /// empty result that looks healthy.</para>
    ///
    /// <para><b>Reads and writes.</b> Reads are volatile and block only through the
    /// bounded <see cref="NatsConnectionManager.WaitBoundedResult{T}"/> /
    /// <see cref="NatsConnectionManager.WaitBounded"/> helpers, so the Excel thread
    /// is never held unbounded. Every write goes through
    /// <see cref="NatsUdfAsync.Run"/> (the AsyncWrites dispatch). All writes in
    /// this file are reply-dependent - they report the server's ack, purge count or
    /// success flag - so the default <c>fireforget</c> mode still awaits the
    /// response and only <c>fireforget-all</c> turns them into a fire-and-forget
    /// mark (PLAN §1.4, DESIGN.md §8.1).</para>
    /// </summary>
    public static class NatsUdfJetStream
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // A JetStream API call is a request/reply with its own client timeout
        // (NatsJSOpts.RequestTimeout, 5s by default). The blocking bound is never
        // shorter than the configured response budget and never longer than this,
        // so a formula cannot hold the Excel thread for minutes.
        private const int MaxJsTimeoutMs = 60000;

        // NatsUDFJSFetch pull window (the [timeoutMs] argument).
        private const int DefaultFetchMs = 1000;
        private const int MaxFetchMs = 60000;

        // NatsUDFJSFetch batch bounds ([count]).
        private const int MaxFetchCount = 10000;

        private const string StorageFile = "file";
        private const string StorageMemory = "memory";

        // ==================================================================
        // §7.2 - streams, messages, consumers (reads)
        // ==================================================================

        /// <summary>Replays one stored message by stream sequence.</summary>
        [ExcelFunction(Description = "Replays a JetStream message by stream sequence", IsVolatile = true)]
        public static object NatsUDFJSGetMessage(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "seq", Description = "Stream sequence number")] object seq,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            if (!TryPositiveSequence(seq, out long sequence, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadGetMessage(endpoint, TextOf(stream), sequence);
        }

        /// <summary>Returns the last stored message matching a subject filter.</summary>
        [ExcelFunction(Description = "Returns the last JetStream message matching a subject filter", IsVolatile = true)]
        public static object NatsUDFJSGetLastMessage(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "filter", Description = "Subject filter (wildcards allowed)")] object filter,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireSubject(filter, allowWildcard: true);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadGetLastMessage(endpoint, TextOf(stream), TextOf(filter));
        }

        /// <summary>
        /// Pulls a batch from an existing consumer into an N×4 matrix:
        /// <c>[0] subject</c>, <c>[1] stream sequence</c>, <c>[2] timestamp</c>,
        /// <c>[3] payload</c>.
        /// </summary>
        /// <remarks>
        /// The batch is not acknowledged, so a pull consumer redelivers those
        /// messages after its <c>ack_wait</c>: <c>JSFetch</c> moves the reader's
        /// view of the stream (DESIGN.md R9), it does not consume for other
        /// subscribers. Use a durable consumer for a stable cursor.
        /// </remarks>
        [ExcelFunction(Description = "Pulls a JetStream batch into an Nx4 matrix", IsVolatile = true)]
        public static object NatsUDFJSFetch(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "consumer", Description = "Consumer name")] object consumer,
            [ExcelArgument(Name = "count", Description = "Maximum messages to pull")] object count,
            [ExcelArgument(Name = "timeoutMs", Description = "Pull window in milliseconds")] object timeoutMs = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireConsumer(consumer);
            if (error != null)
                return error;
            if (!TryToInt(count, "count", 1, MaxFetchCount, 0, out int maxMessages, out error))
                return error;
            if (!TryToInt(timeoutMs, "timeoutMs", 1, MaxFetchMs, DefaultFetchMs, out int expiresMs, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadFetch(endpoint, TextOf(stream), TextOf(consumer), maxMessages, expiresMs);
        }

        /// <summary>Stream state (messages, bytes, first/last sequence) as JSON.</summary>
        [ExcelFunction(Description = "Returns JetStream stream info as JSON", IsVolatile = true)]
        public static object NatsUDFJSStreamInfo(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadStreamInfo(endpoint, TextOf(stream));
        }

        /// <summary>
        /// Lists the streams as a matrix with a header row:
        /// <c>name</c>, <c>subjects</c>, <c>messages</c>, <c>bytes</c>,
        /// <c>consumers</c>, <c>storage</c>.
        /// </summary>
        [ExcelFunction(Description = "Lists JetStream streams", IsVolatile = true)]
        public static object NatsUDFJSStreams(
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            if (!TryPrepare(host, admin: false, out string endpoint, out string error))
                return error;

            return ReadStreams(endpoint);
        }

        /// <summary>Consumer progress (pending, redelivered, ack floor) as JSON.</summary>
        [ExcelFunction(Description = "Returns JetStream consumer info as JSON", IsVolatile = true)]
        public static object NatsUDFJSConsumerInfo(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "consumer", Description = "Consumer name")] object consumer,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireConsumer(consumer);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadConsumerInfo(endpoint, TextOf(stream), TextOf(consumer));
        }

        /// <summary>
        /// Lists a stream's consumers as a matrix with a header row:
        /// <c>name</c>, <c>durable</c>, <c>pending</c>, <c>ack_pending</c>,
        /// <c>waiting</c>, <c>delivered_seq</c>, <c>ack_floor_seq</c>.
        /// </summary>
        [ExcelFunction(Description = "Lists a stream's JetStream consumers", IsVolatile = true)]
        public static object NatsUDFJSConsumers(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            return ReadConsumers(endpoint, TextOf(stream));
        }

        // ==================================================================
        // §7.2 - writes (+ ...NonVolatile twins)
        // ==================================================================

        /// <summary>
        /// Publishes to a JetStream subject and returns the stored sequence as a
        /// number (or the fire-and-forget marker under <c>fireforget-all</c>).
        /// </summary>
        [ExcelFunction(Description = "Publishes to a JetStream subject and returns the sequence", IsVolatile = true)]
        public static object NatsUDFJSPublish(
            [ExcelArgument(Name = "subject", Description = "Concrete subject (no wildcards)")] object subject,
            [ExcelArgument(Name = "message", Description = "Payload text")] object message,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
            if (error != null)
                return error;
            if (!TryPayload(message, out string payload, out error))
                return error;
            if (!TryPrepare(host, admin: false, out string endpoint, out error))
                return error;

            string subjectText = TextOf(subject);
            return NatsUdfAsync.Run("NatsUDFJSPublish", host, new object[] { subject, message, host },
                () => WritePublish(endpoint, subjectText, payload));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFJSPublish"/>.</summary>
        [ExcelFunction(Description = "Publishes to a JetStream subject and returns the sequence", IsVolatile = false)]
        public static object NatsUDFJSPublishNonVolatile(
            [ExcelArgument(Name = "subject", Description = "Concrete subject (no wildcards)")] object subject,
            [ExcelArgument(Name = "message", Description = "Payload text")] object message,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFJSPublish(subject, message, host);
        }

        /// <summary>
        /// Publishes an Excel range converted to JSON (matrix → JSON, the
        /// <see cref="NatsJson.MatrixToJson"/> shape) and returns the sequence.
        /// </summary>
        [ExcelFunction(Description = "Publishes a range as JSON to a JetStream subject", IsVolatile = true)]
        public static object NatsUDFJSPublishJSON(
            [ExcelArgument(Name = "subject", Description = "Concrete subject (no wildcards)")] object subject,
            [ExcelArgument(Name = "range", Description = "Range to serialize as JSON")] object[,] range,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireSubject(subject, allowWildcard: false);
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

            string subjectText = TextOf(subject);
            return NatsUdfAsync.Run("NatsUDFJSPublishJSON", host, new object[] { subject, range, host },
                () => WritePublish(endpoint, subjectText, json ?? string.Empty));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFJSPublishJSON"/>.</summary>
        [ExcelFunction(Description = "Publishes a range as JSON to a JetStream subject", IsVolatile = false)]
        public static object NatsUDFJSPublishJSONNonVolatile(
            [ExcelArgument(Name = "subject", Description = "Concrete subject (no wildcards)")] object subject,
            [ExcelArgument(Name = "range", Description = "Range to serialize as JSON")] object[,] range,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFJSPublishJSON(subject, range, host);
        }

        /// <summary>Deletes one stored message by sequence. Destructive; gated.</summary>
        [ExcelFunction(Description = "Deletes a JetStream message by sequence", IsVolatile = true)]
        public static object NatsUDFJSDeleteMessage(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "seq", Description = "Stream sequence number")] object seq,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            if (!TryPositiveSequence(seq, out long sequence, out error))
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string streamText = TextOf(stream);
            return NatsUdfAsync.Run("NatsUDFJSDeleteMessage", host, new object[] { stream, seq, host },
                () => WriteDeleteMessage(endpoint, streamText, sequence));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFJSDeleteMessage"/>.</summary>
        [ExcelFunction(Description = "Deletes a JetStream message by sequence", IsVolatile = false)]
        public static object NatsUDFJSDeleteMessageNonVolatile(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "seq", Description = "Stream sequence number")] object seq,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFJSDeleteMessage(stream, seq, host);
        }

        /// <summary>
        /// Purges a whole stream, or only the messages matching
        /// <c>[filter]</c>, and returns the number of purged messages.
        /// Destructive; gated.
        /// </summary>
        [ExcelFunction(Description = "Purges a JetStream stream (optionally by subject filter)", IsVolatile = true)]
        public static object NatsUDFJSPurge(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "filter", Description = "Optional subject filter (wildcards allowed)")] object filter = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;

            string filterText = null;
            if (!IsBlank(filter))
            {
                // Purge accepts "*"/">" tokens in the filter; the value must be a subject.
                error = NatsUdfCommon.RequireSubject(filter, allowWildcard: true);
                if (error != null)
                    return error;
                filterText = TextOf(filter);
            }

            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string streamText = TextOf(stream);
            return NatsUdfAsync.Run("NatsUDFJSPurge", host, new object[] { stream, filter, host },
                () => WritePurge(endpoint, streamText, filterText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFJSPurge"/>.</summary>
        [ExcelFunction(Description = "Purges a JetStream stream (optionally by subject filter)", IsVolatile = false)]
        public static object NatsUDFJSPurgeNonVolatile(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "filter", Description = "Optional subject filter (wildcards allowed)")] object filter = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFJSPurge(stream, filter, host);
        }

        // ==================================================================
        // §7.3 - admin (opt-in; every function here is gated)
        // ==================================================================

        /// <summary>
        /// Creates a stream. <c>[subjects]</c> is a range, a single subject, or a
        /// comma/newline separated list; <c>[storage]</c> is <c>file</c>
        /// (default) or <c>memory</c>; <c>[retention]</c> is <c>limits</c>
        /// (default), <c>interest</c> or <c>workqueue</c>. Gated.
        /// </summary>
        [ExcelFunction(Description = "Creates a JetStream stream", IsVolatile = true)]
        public static object NatsUDFStreamCreate(
            [ExcelArgument(Name = "name", Description = "New stream name")] object name,
            [ExcelArgument(Name = "subjects", Description = "Subject range, single subject or comma list")] object subjects,
            [ExcelArgument(Name = "storage", Description = "file (default) or memory")] object storage = null,
            [ExcelArgument(Name = "retention", Description = "limits (default), interest or workqueue")] object retention = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(name);
            if (error != null)
                return error;
            error = ReadSubjects(subjects, out List<string> subjectList);
            if (error != null)
                return error;
            if (!TryParseStorage(storage, out StreamConfigStorage storageMode, out error))
                return error;
            if (!TryParseRetention(retention, out StreamConfigRetention retentionMode, out error))
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFStreamCreate", host, new object[] { name, subjects, storage, retention, host },
                () => WriteStreamCreate(endpoint, nameText, subjectList, storageMode, retentionMode));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFStreamCreate"/>.</summary>
        [ExcelFunction(Description = "Creates a JetStream stream", IsVolatile = false)]
        public static object NatsUDFStreamCreateNonVolatile(
            [ExcelArgument(Name = "name", Description = "New stream name")] object name,
            [ExcelArgument(Name = "subjects", Description = "Subject range, single subject or comma list")] object subjects,
            [ExcelArgument(Name = "storage", Description = "file (default) or memory")] object storage = null,
            [ExcelArgument(Name = "retention", Description = "limits (default), interest or workqueue")] object retention = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFStreamCreate(name, subjects, storage, retention, host);
        }

        /// <summary>
        /// Updates a stream's subject set. The signature mirrors
        /// <see cref="NatsUDFStreamCreate"/>, but <c>[storage]</c> and
        /// <c>[retention]</c> are immutable on an existing stream: supplying
        /// either is rejected with a clear error instead of silently ignoring it
        /// (delete and recreate the stream to change them). Gated.
        /// </summary>
        [ExcelFunction(Description = "Updates a JetStream stream's subjects", IsVolatile = true)]
        public static object NatsUDFStreamUpdate(
            [ExcelArgument(Name = "name", Description = "Existing stream name")] object name,
            [ExcelArgument(Name = "subjects", Description = "Replacement subject range, single subject or comma list")] object subjects,
            [ExcelArgument(Name = "storage", Description = "Immutable on update; must be blank")] object storage = null,
            [ExcelArgument(Name = "retention", Description = "Immutable on update; must be blank")] object retention = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(name);
            if (error != null)
                return error;
            error = ReadSubjects(subjects, out List<string> subjectList);
            if (error != null)
                return error;

            if (!IsBlank(storage))
                return NatsUdfCommon.Fail("storage is immutable; delete the stream and recreate it to change the storage backend");
            if (!IsBlank(retention))
                return NatsUdfCommon.Fail("retention is immutable; delete the stream and recreate it to change the retention policy");

            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFStreamUpdate", host, new object[] { name, subjects, storage, retention, host },
                () => WriteStreamUpdate(endpoint, nameText, subjectList));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFStreamUpdate"/>.</summary>
        [ExcelFunction(Description = "Updates a JetStream stream's subjects", IsVolatile = false)]
        public static object NatsUDFStreamUpdateNonVolatile(
            [ExcelArgument(Name = "name", Description = "Existing stream name")] object name,
            [ExcelArgument(Name = "subjects", Description = "Replacement subject range, single subject or comma list")] object subjects,
            [ExcelArgument(Name = "storage", Description = "Immutable on update; must be blank")] object storage = null,
            [ExcelArgument(Name = "retention", Description = "Immutable on update; must be blank")] object retention = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFStreamUpdate(name, subjects, storage, retention, host);
        }

        /// <summary>Deletes a stream and all of its messages. Destructive; gated.</summary>
        [ExcelFunction(Description = "Deletes a JetStream stream", IsVolatile = true)]
        public static object NatsUDFStreamDelete(
            [ExcelArgument(Name = "name", Description = "Stream name to delete")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(name);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFStreamDelete", host, new object[] { name, host },
                () => WriteStreamDelete(endpoint, nameText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFStreamDelete"/>.</summary>
        [ExcelFunction(Description = "Deletes a JetStream stream", IsVolatile = false)]
        public static object NatsUDFStreamDeleteNonVolatile(
            [ExcelArgument(Name = "name", Description = "Stream name to delete")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFStreamDelete(name, host);
        }

        /// <summary>
        /// Creates (or updates) a durable pull consumer. <c>[filter]</c> is an
        /// optional subject filter; <c>[deliverPolicy]</c> is one of <c>all</c>
        /// (default), <c>last</c>, <c>new</c>, <c>by_start_sequence</c>,
        /// <c>by_start_time</c> or <c>last_per_subject</c>. Gated.
        /// </summary>
        [ExcelFunction(Description = "Creates or updates a durable JetStream consumer", IsVolatile = true)]
        public static object NatsUDFConsumerCreate(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Durable consumer name")] object name,
            [ExcelArgument(Name = "filter", Description = "Optional subject filter")] object filter = null,
            [ExcelArgument(Name = "deliverPolicy", Description = "all (default), last, new, by_start_sequence, by_start_time, last_per_subject")] object deliverPolicy = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireConsumer(name);
            if (error != null)
                return error;

            string filterText = null;
            if (!IsBlank(filter))
            {
                error = NatsUdfCommon.RequireSubject(filter, allowWildcard: true);
                if (error != null)
                    return error;
                filterText = TextOf(filter);
            }

            if (!TryParseDeliverPolicy(deliverPolicy, out ConsumerConfigDeliverPolicy policy, out error))
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string streamText = TextOf(stream);
            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFConsumerCreate", host, new object[] { stream, name, filter, deliverPolicy, host },
                () => WriteConsumerCreate(endpoint, streamText, nameText, filterText, policy));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFConsumerCreate"/>.</summary>
        [ExcelFunction(Description = "Creates or updates a durable JetStream consumer", IsVolatile = false)]
        public static object NatsUDFConsumerCreateNonVolatile(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Durable consumer name")] object name,
            [ExcelArgument(Name = "filter", Description = "Optional subject filter")] object filter = null,
            [ExcelArgument(Name = "deliverPolicy", Description = "all (default), last, new, by_start_sequence, by_start_time, last_per_subject")] object deliverPolicy = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFConsumerCreate(stream, name, filter, deliverPolicy, host);
        }

        /// <summary>Deletes a consumer. Destructive; gated.</summary>
        [ExcelFunction(Description = "Deletes a JetStream consumer", IsVolatile = true)]
        public static object NatsUDFConsumerDelete(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Consumer name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireConsumer(name);
            if (error != null)
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string streamText = TextOf(stream);
            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFConsumerDelete", host, new object[] { stream, name, host },
                () => WriteConsumerDelete(endpoint, streamText, nameText));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFConsumerDelete"/>.</summary>
        [ExcelFunction(Description = "Deletes a JetStream consumer", IsVolatile = false)]
        public static object NatsUDFConsumerDeleteNonVolatile(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Consumer name")] object name,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFConsumerDelete(stream, name, host);
        }

        /// <summary>
        /// Pauses a consumer for <c>[pauseMs]</c> milliseconds (NATS server
        /// 2.11+). Destructive; gated.
        /// </summary>
        [ExcelFunction(Description = "Pauses a JetStream consumer", IsVolatile = true)]
        public static object NatsUDFConsumerPause(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Consumer name")] object name,
            [ExcelArgument(Name = "pauseMs", Description = "Pause duration in milliseconds")] object pauseMs = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            string error = NatsUdfCommon.RequireStream(stream);
            if (error != null)
                return error;
            error = NatsUdfCommon.RequireConsumer(name);
            if (error != null)
                return error;
            if (!TryToInt(pauseMs, "pauseMs", 1, MaxFetchMs, 0, out int pauseMillis, out error))
                return error;
            if (!TryPrepare(host, admin: true, out string endpoint, out error))
                return error;

            string streamText = TextOf(stream);
            string nameText = TextOf(name);
            return NatsUdfAsync.Run("NatsUDFConsumerPause", host, new object[] { stream, name, pauseMs, host },
                () => WriteConsumerPause(endpoint, streamText, nameText, pauseMillis));
        }

        /// <summary>Non-volatile twin of <see cref="NatsUDFConsumerPause"/>.</summary>
        [ExcelFunction(Description = "Pauses a JetStream consumer", IsVolatile = false)]
        public static object NatsUDFConsumerPauseNonVolatile(
            [ExcelArgument(Name = "stream", Description = "JetStream stream name")] object stream,
            [ExcelArgument(Name = "name", Description = "Consumer name")] object name,
            [ExcelArgument(Name = "pauseMs", Description = "Pause duration in milliseconds")] object pauseMs = null,
            [ExcelArgument(Name = "host", Description = "Optional cluster, server alias or URL")] object host = null)
        {
            return NatsUDFConsumerPause(stream, name, pauseMs, host);
        }

        // ==================================================================
        // Read implementations (bounded blocking only)
        // ==================================================================

        private static object ReadGetMessage(string endpoint, string stream, long sequence)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSStream handle = WaitResult(context.GetStreamAsync(stream).AsTask(), "JetStream stream info", stream, null);
                StreamMsgGetResponse response = WaitResult(
                    handle.GetAsync(new StreamMsgGetRequest { Seq = (ulong)sequence }).AsTask(),
                    "JetStream get message", stream, null);
                return NatsResultFormatter.BoundCellText(MessageValue(response));
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream get message", ex, stream, null);
            }
        }

        private static object ReadGetLastMessage(string endpoint, string stream, string filter)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSStream handle = WaitResult(context.GetStreamAsync(stream).AsTask(), "JetStream stream info", stream, null);
                StreamMsgGetResponse response = WaitResult(
                    handle.GetAsync(new StreamMsgGetRequest { LastBySubj = filter }).AsTask(),
                    "JetStream get last message", stream, null);
                return NatsResultFormatter.BoundCellText(MessageValue(response));
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream get last message", ex, stream, null);
            }
        }

        private static object ReadFetch(string endpoint, string stream, string consumer, int maxMessages, int expiresMs)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSConsumer handle = WaitResult(
                    context.GetConsumerAsync(stream, consumer).AsTask(),
                    "JetStream consumer info", stream, consumer);

                // The pull itself can legitimately take the whole window, so the
                // bound is the window plus the normal response budget.
                int bound = Better(expiresMs + JsTimeoutMs(), JsTimeoutMs());
                List<string[]> rows = WaitResult(
                    CollectFetchAsync(handle, maxMessages, expiresMs),
                    bound, "JetStream fetch", stream, consumer);

                return NatsResultFormatter.FormatRowsMatrix(rows);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream fetch", ex, stream, consumer);
            }
        }

        private static object ReadStreamInfo(string endpoint, string stream)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSStream handle = WaitResult(context.GetStreamAsync(stream).AsTask(), "JetStream stream info", stream, null);
                StreamInfo info = handle.Info;
                return NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatStreamInfoJson(
                    NameOr(info.Config?.Name, stream),
                    info.Config?.Subjects,
                    info.State.Messages,
                    info.State.Bytes,
                    ToLong(info.State.FirstSeq),
                    ToLong(info.State.LastSeq),
                    info.State.ConsumerCount,
                    StorageName(info.Config)));
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream stream info", ex, stream, null);
            }
        }

        private static object ReadStreams(string endpoint)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                List<INatsJSStream> streams = WaitResult(
                    CollectStreamsAsync(context), "JetStream list streams", null, null);

                var rows = new List<string[]>(streams.Count + 1)
                {
                    new[] { "name", "subjects", "messages", "bytes", "consumers", "storage" }
                };
                foreach (INatsJSStream stream in streams)
                {
                    StreamInfo info = stream.Info;
                    rows.Add(new[]
                    {
                        NameOr(info.Config?.Name, string.Empty),
                        SubjectsCsv(info.Config?.Subjects),
                        info.State.Messages.ToString(CultureInfo.InvariantCulture),
                        info.State.Bytes.ToString(CultureInfo.InvariantCulture),
                        info.State.ConsumerCount.ToString(CultureInfo.InvariantCulture),
                        StorageName(info.Config)
                    });
                }

                return NatsResultFormatter.FormatRowsMatrix(rows);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream list streams", ex, null, null);
            }
        }

        private static object ReadConsumerInfo(string endpoint, string stream, string consumer)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSConsumer handle = WaitResult(
                    context.GetConsumerAsync(stream, consumer).AsTask(),
                    "JetStream consumer info", stream, consumer);
                ConsumerInfo info = handle.Info;
                return NatsResultFormatter.BoundCellText(NatsResultFormatter.FormatConsumerInfoJson(
                    NameOr(info.StreamName, stream),
                    NameOr(info.Name, consumer),
                    !string.IsNullOrEmpty(info.Config?.DurableName),
                    ToLong(info.NumPending),
                    info.NumRedelivered,
                    info.NumAckPending,
                    info.NumWaiting,
                    ToLong(info.Delivered.StreamSeq),
                    ToLong(info.AckFloor.StreamSeq)));
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream consumer info", ex, stream, consumer);
            }
        }

        private static object ReadConsumers(string endpoint, string stream)
        {
            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                List<INatsJSConsumer> consumers = WaitResult(
                    CollectConsumersAsync(context, stream), "JetStream list consumers", stream, null);

                var rows = new List<string[]>(consumers.Count + 1)
                {
                    new[] { "name", "durable", "pending", "ack_pending", "waiting", "delivered_seq", "ack_floor_seq" }
                };
                foreach (INatsJSConsumer consumer in consumers)
                {
                    ConsumerInfo info = consumer.Info;
                    rows.Add(new[]
                    {
                        NameOr(info.Name, string.Empty),
                        string.IsNullOrEmpty(info.Config?.DurableName) ? "false" : "true",
                        info.NumPending.ToString(CultureInfo.InvariantCulture),
                        info.NumAckPending.ToString(CultureInfo.InvariantCulture),
                        info.NumWaiting.ToString(CultureInfo.InvariantCulture),
                        info.Delivered.StreamSeq.ToString(CultureInfo.InvariantCulture),
                        info.AckFloor.StreamSeq.ToString(CultureInfo.InvariantCulture)
                    });
                }

                return NatsResultFormatter.FormatRowsMatrix(rows);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream list consumers", ex, stream, null);
            }
        }

        /// <summary>
        /// Drains the paginated stream listing off the Excel thread; the wait for
        /// it is the caller's bounded <see cref="WaitResult{T}(Task{T}, int, string, string, string, string)"/>.
        /// </summary>
        private static Task<List<INatsJSStream>> CollectStreamsAsync(INatsJSContext context)
        {
            return Task.Run(async () =>
            {
                var streams = new List<INatsJSStream>();
                await foreach (INatsJSStream stream in context.ListStreamsAsync())
                    streams.Add(stream);
                return streams;
            });
        }

        /// <summary>Drains the paginated consumer listing off the Excel thread.</summary>
        private static Task<List<INatsJSConsumer>> CollectConsumersAsync(INatsJSContext context, string stream)
        {
            return Task.Run(async () =>
            {
                var consumers = new List<INatsJSConsumer>();
                await foreach (INatsJSConsumer consumer in context.ListConsumersAsync(stream))
                    consumers.Add(consumer);
                return consumers;
            });
        }

        /// <summary>Pulls one batch into N×4 rows (subject, sequence, timestamp, payload).</summary>
        private static Task<List<string[]>> CollectFetchAsync(INatsJSConsumer consumer, int maxMessages, int expiresMs)
        {
            return Task.Run(async () =>
            {
                var rows = new List<string[]>(maxMessages);
                var opts = new NatsJSFetchOpts
                {
                    MaxMsgs = maxMessages,
                    Expires = TimeSpan.FromMilliseconds(expiresMs)
                };

                await foreach (INatsJSMsg<byte[]> message in consumer.FetchAsync<byte[]>(opts))
                {
                    NatsJSMsgMetadata? metadata = message.Metadata;
                    long sequence = metadata.HasValue ? ToLong(metadata.Value.Sequence.Stream) : 0L;
                    string timestamp = metadata.HasValue
                        ? metadata.Value.Timestamp.UtcDateTime.ToString("o", CultureInfo.InvariantCulture)
                        : string.Empty;
                    rows.Add(new[]
                    {
                        message.Subject ?? string.Empty,
                        sequence.ToString(CultureInfo.InvariantCulture),
                        timestamp,
                        NatsResultFormatter.DecodePayload(message.Data) ?? string.Empty
                    });
                }

                return rows;
            });
        }

        // ==================================================================
        // Write implementations (every caller runs these through NatsUdfAsync)
        // ==================================================================

        private static object WritePublish(string endpoint, string subject, string payload)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFJSPublish", endpoint, ctx => ctx.PublishAsync(subject, payload).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            // A publish to a subject no stream listens on hangs until the ack
            // request times out; the bounded wait is what keeps that off the
            // Excel thread.
            try
            {
                PubAckResponse ack = WaitResult(
                    context.PublishAsync(subject, payload).AsTask(), "JetStream publish", null, null);
                if (ack == null)
                    return NatsUdfCommon.VerboseError("JetStream publish", "no acknowledgment returned", NatsUdfCommon.ErrorHint("timed out"));
                if (ack.Error != null)
                {
                    string cause = ApiCause(ack.Error);
                    return NatsUdfCommon.VerboseError("JetStream publish", cause, NatsUdfCommon.ErrorHint(cause));
                }

                return (double)ack.Seq;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream publish", ex, null, null);
            }
        }

        private static object WriteDeleteMessage(string endpoint, string stream, long sequence)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFJSDeleteMessage", endpoint,
                    ctx => ctx.DeleteMessageAsync(stream, new StreamMsgDeleteRequest { Seq = (ulong)sequence }).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                StreamMsgDeleteResponse response = WaitResult(
                    context.DeleteMessageAsync(stream, new StreamMsgDeleteRequest { Seq = (ulong)sequence }).AsTask(),
                    "JetStream delete message", stream, null);
                return response.Success
                    ? NatsUdfCommon.Ok
                    : NatsUdfCommon.VerboseError("JetStream delete message", "the server reported no success", null);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream delete message", ex, stream, null);
            }
        }

        private static object WritePurge(string endpoint, string stream, string filter)
        {
            var request = new StreamPurgeRequest { Filter = filter };

            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFJSPurge", endpoint, ctx => ctx.PurgeStreamAsync(stream, request).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                StreamPurgeResponse response = WaitResult(
                    context.PurgeStreamAsync(stream, request).AsTask(), "JetStream purge", stream, null);
                return (double)response.Purged;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream purge", ex, stream, null);
            }
        }

        private static object WriteStreamCreate(string endpoint, string name, List<string> subjects, StreamConfigStorage storage, StreamConfigRetention retention)
        {
            var config = new StreamConfig(name, subjects)
            {
                Storage = storage,
                Retention = retention
            };

            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFStreamCreate", endpoint, ctx => ctx.CreateStreamAsync(config).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                WaitResult(context.CreateStreamAsync(config).AsTask(), "JetStream stream create", name, null);
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream stream create", ex, name, null);
            }
        }

        private static object WriteStreamUpdate(string endpoint, string name, List<string> subjects)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
            {
                // A partial update would reset every field the caller did not
                // name, so the fire-and-forget path still reads the current
                // config first and updates from it.
                return FireAndForget("NatsUDFStreamUpdate", endpoint, async ctx =>
                {
                    INatsJSStream existing = await ctx.GetStreamAsync(name).ConfigureAwait(false);
                    StreamConfig updated = existing.Info.Config with { Subjects = subjects };
                    await ctx.UpdateStreamAsync(updated).ConfigureAwait(false);
                });
            }

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                INatsJSStream existing = WaitResult(
                    context.GetStreamAsync(name).AsTask(), "JetStream stream update", name, null);
                StreamConfig updated = existing.Info.Config with { Subjects = subjects };
                WaitResult(context.UpdateStreamAsync(updated).AsTask(), "JetStream stream update", name, null);
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream stream update", ex, name, null);
            }
        }

        private static object WriteStreamDelete(string endpoint, string name)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFStreamDelete", endpoint, ctx => ctx.DeleteStreamAsync(name).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                bool deleted = WaitResult(
                    context.DeleteStreamAsync(name).AsTask(), "JetStream stream delete", name, null);
                return deleted
                    ? NatsUdfCommon.Ok
                    : NatsUdfCommon.VerboseError("JetStream stream delete", "the server reported no success", null);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream stream delete", ex, name, null);
            }
        }

        private static object WriteConsumerCreate(string endpoint, string stream, string name, string filter, ConsumerConfigDeliverPolicy policy)
        {
            // The name constructor makes it durable (Name + DurableName) with an
            // explicit ack policy, which is what a pull consumer needs (PLAN S5.17).
            var config = new ConsumerConfig(name)
            {
                FilterSubject = string.IsNullOrWhiteSpace(filter) ? null : filter,
                DeliverPolicy = policy
            };

            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFConsumerCreate", endpoint, ctx => ctx.CreateOrUpdateConsumerAsync(stream, config).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                WaitResult(context.CreateOrUpdateConsumerAsync(stream, config).AsTask(),
                    "JetStream consumer create", stream, name);
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream consumer create", ex, stream, name);
            }
        }

        private static object WriteConsumerDelete(string endpoint, string stream, string name)
        {
            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFConsumerDelete", endpoint, ctx => ctx.DeleteConsumerAsync(stream, name).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                bool deleted = WaitResult(
                    context.DeleteConsumerAsync(stream, name).AsTask(),
                    "JetStream consumer delete", stream, name);
                return deleted
                    ? NatsUdfCommon.Ok
                    : NatsUdfCommon.VerboseError("JetStream consumer delete", "the server reported no success", null);
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream consumer delete", ex, stream, name);
            }
        }

        private static object WriteConsumerPause(string endpoint, string stream, string name, int pauseMs)
        {
            DateTimeOffset until = DateTimeOffset.UtcNow.AddMilliseconds(pauseMs);

            if (NatsUdfCommon.ShouldFireAndForget(AppConfig.SyncWrite, replyDependent: true))
                return FireAndForget("NatsUDFConsumerPause", endpoint, ctx => ctx.PauseConsumerAsync(stream, name, until).AsTask());

            if (!TryGetJetStream(endpoint, out INatsJSContext context, out string error))
                return error;

            try
            {
                WaitResult(context.PauseConsumerAsync(stream, name, until).AsTask(),
                    "JetStream consumer pause", stream, name);
                return NatsUdfCommon.Ok;
            }
            catch (Exception ex)
            {
                return MapJsError("JetStream consumer pause", ex, stream, name);
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
                var connection = NatsRuntime.Connections.GetConnection(endpoint);
                if (connection.ServerInfo == null)
                    NatsConnectionManager.WaitBounded(connection.PingAsync().AsTask(), JsTimeoutMs());
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("connection", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (!NatsRuntime.Connections.IsJetStreamEnabled(endpoint))
            {
                error = NatsUdfCommon.RequireJetStream(false);
                return false;
            }

            return true;
        }

        private static bool TryGetJetStream(string endpoint, out INatsJSContext context, out string error)
        {
            context = null;
            error = null;
            try
            {
                context = NatsRuntime.Connections.GetJetStream(endpoint);
            }
            catch (Exception ex)
            {
                error = NatsUdfCommon.VerboseError("JetStream", NatsUdfAsync.SafeMessage(ex), null);
                return false;
            }

            if (context == null)
            {
                error = NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Bounded wait for a JetStream result: the only place this file blocks.
        /// A timeout becomes a named, hint-carrying cell error; any other failure
        /// goes through <see cref="MapJsError"/>.
        /// </summary>
        private static T WaitResult<T>(Task<T> task, string what, string stream, string consumer)
        {
            return WaitResult(task, JsTimeoutMs(), what, stream, consumer);
        }

        private static T WaitResult<T>(Task<T> task, int timeoutMs, string what, string stream, string consumer)
        {
            try
            {
                return NatsConnectionManager.WaitBoundedResult(task, timeoutMs);
            }
            catch (TimeoutException)
            {
                throw new JsTimeoutException(what, timeoutMs);
            }
        }

        /// <summary>The blocking bound for one JetStream round trip.</summary>
        private static int JsTimeoutMs()
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
            if (configured > MaxJsTimeoutMs)
                configured = MaxJsTimeoutMs;
            return configured;
        }

        /// <summary>Fire-and-forget for a reply-dependent write forced by <c>fireforget-all</c>.</summary>
        private static object FireAndForget(string what, string endpoint, Func<INatsJSContext, Task> operation)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    INatsJSContext context = NatsRuntime.Connections.GetJetStream(endpoint);
                    await operation(context).ConfigureAwait(false);
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

        // ==================================================================
        // Argument parsing
        // ==================================================================

        private static bool TryPayload(object message, out string payload, out string error)
        {
            error = null;
            try
            {
                // A truly missing message is a legitimate empty NATS payload.
                payload = NatsUdfCommon.ToNatsString(message) ?? string.Empty;
                return true;
            }
            catch (ArgumentException ex)
            {
                payload = null;
                error = NatsUdfCommon.Fail(ex.Message);
                return false;
            }
        }

        private static bool TryPositiveSequence(object value, out long sequence, out string error)
        {
            error = null;
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
            {
                sequence = 0;
                error = NatsUdfCommon.Fail("a sequence number is required");
                return false;
            }
            if (!NatsUdfCommon.TryToLong(value, out sequence) || sequence <= 0)
            {
                sequence = 0;
                error = NatsUdfCommon.Fail("the sequence number must be a positive integer");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Integral argument with bounds. A missing/blank cell uses
        /// <paramref name="fallback"/> (which, when below <paramref name="min"/>,
        /// still marks the argument as required); a present value outside the
        /// bounds is a terse argument error.
        /// </summary>
        private static bool TryToInt(object value, string what, int min, int max, int fallback, out int result, out string error)
        {
            error = null;
            if (IsBlank(value))
            {
                result = fallback;
                if (fallback < min)
                {
                    error = NatsUdfCommon.Fail("a " + what + " is required");
                    return false;
                }
                return true;
            }

            if (!NatsUdfCommon.TryToInt(value, out result) || result < min || result > max)
            {
                result = 0;
                error = NatsUdfCommon.Fail(
                    "the " + what + " must be between " + min.ToString(CultureInfo.InvariantCulture)
                    + " and " + max.ToString(CultureInfo.InvariantCulture));
                return false;
            }

            return true;
        }

        /// <summary>
        /// Reads a subjects argument: a range (any shape), a single subject, or a
        /// comma/newline separated list. Blank cells are skipped.
        /// </summary>
        private static string ReadSubjects(object value, out List<string> subjects)
        {
            subjects = new List<string>();
            if (IsBlank(value))
                return NatsUdfCommon.Fail("a subjects range is required");
            if (value is ExcelError)
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorExcelErrorCell);

            if (value is object[,] range)
            {
                string rangeError = NatsUdfCommon.RequireRange(range, "subjects range");
                if (rangeError != null)
                    return rangeError;

                foreach (object cell in range)
                {
                    // Reject an Excel error cell explicitly: TextOf swallows it
                    // (ToNatsString throws and TextOf returns null), so a #N/A
                    // would be skipped like a blank and let the stream be created
                    // with fewer subjects than the sheet shows. RequireRange
                    // already checks the range, but it is re-read here, so the
                    // per-cell guard is repeated (PLAN 1.3).
                    if (cell is ExcelError)
                        return NatsUdfCommon.Fail(NatsUdfCommon.ErrorExcelErrorCell);

                    string text = TextOf(cell);
                    if (!string.IsNullOrWhiteSpace(text))
                        subjects.Add(text);
                }
            }
            else if (value is string)
            {
                // A single cell may hold one subject or a comma/newline separated
                // list; wildcards are allowed (a stream consumes them).
                string singleError = NatsUdfCommon.RequireSubject(value, allowWildcard: true);
                if (singleError != null)
                    return singleError;

                string text = TextOf(value);
                foreach (string part in text.Split(new[] { ',', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (!string.IsNullOrWhiteSpace(part))
                        subjects.Add(part.Trim());
                }
            }
            else
            {
                return NatsUdfCommon.Fail("a subjects range is required");
            }

            if (subjects.Count == 0)
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorNoEntries);
            return null;
        }

        private static bool TryParseStorage(object value, out StreamConfigStorage storage, out string error)
        {
            storage = StreamConfigStorage.File;
            error = null;
            if (IsBlank(value))
                return true;

            switch (NormalizeToken(TextOf(value)))
            {
                case StorageFile:
                    storage = StreamConfigStorage.File;
                    return true;
                case StorageMemory:
                    storage = StreamConfigStorage.Memory;
                    return true;
                default:
                    error = NatsUdfCommon.Fail("the storage must be 'file' or 'memory'");
                    return false;
            }
        }

        private static bool TryParseRetention(object value, out StreamConfigRetention retention, out string error)
        {
            retention = StreamConfigRetention.Limits;
            error = null;
            if (IsBlank(value))
                return true;

            switch (NormalizeToken(TextOf(value)))
            {
                case "limits":
                    retention = StreamConfigRetention.Limits;
                    return true;
                case "interest":
                    retention = StreamConfigRetention.Interest;
                    return true;
                case "workqueue":
                    retention = StreamConfigRetention.Workqueue;
                    return true;
                default:
                    error = NatsUdfCommon.Fail("the retention must be 'limits', 'interest' or 'workqueue'");
                    return false;
            }
        }

        private static bool TryParseDeliverPolicy(object value, out ConsumerConfigDeliverPolicy policy, out string error)
        {
            policy = ConsumerConfigDeliverPolicy.All;
            error = null;
            if (IsBlank(value))
                return true;

            switch (NormalizeToken(TextOf(value)))
            {
                case "all":
                    policy = ConsumerConfigDeliverPolicy.All;
                    return true;
                case "last":
                    policy = ConsumerConfigDeliverPolicy.Last;
                    return true;
                case "new":
                    policy = ConsumerConfigDeliverPolicy.New;
                    return true;
                case "bystartsequence":
                    policy = ConsumerConfigDeliverPolicy.ByStartSequence;
                    return true;
                case "bystarttime":
                    policy = ConsumerConfigDeliverPolicy.ByStartTime;
                    return true;
                case "lastpersubject":
                    policy = ConsumerConfigDeliverPolicy.LastPerSubject;
                    return true;
                default:
                    error = NatsUdfCommon.Fail(
                        "the deliverPolicy must be 'all', 'last', 'new', 'by_start_sequence', 'by_start_time' or 'last_per_subject'");
                    return false;
            }
        }

        /// <summary>Lowercase token with separators removed ("by_start_sequence" → "bystartsequence").</summary>
        private static string NormalizeToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;
            return text.Trim().ToLowerInvariant().Replace("_", string.Empty).Replace("-", string.Empty);
        }

        // ==================================================================
        // Error mapping (every cause is masked by VerboseError)
        // ==================================================================

        private static string MapJsError(string what, Exception ex, string stream, string consumer)
        {
            if (ex is JsTimeoutException timeout)
            {
                string cause = "timed out after " + timeout.TimeoutMs.ToString(CultureInfo.InvariantCulture) + "ms";
                return NatsUdfCommon.VerboseError(timeout.What, cause, NatsUdfCommon.ErrorHint("timed out"));
            }

            if (TryApiError(ex, out ApiError api))
            {
                string cause = ApiCause(api);
                string lower = cause.ToLowerInvariant();

                if (api.Code == 404)
                {
                    if (consumer != null && lower.Contains("consumer"))
                    {
                        return NatsUdfCommon.VerboseError(
                            what, NatsUdfCommon.ConsumerNotFound(consumer, stream), NatsUdfCommon.ErrorHint("consumer not found"));
                    }
                    if (lower.Contains("message"))
                        return NatsUdfCommon.NoValue;
                    if (stream != null)
                    {
                        return NatsUdfCommon.VerboseError(
                            what, NatsUdfCommon.StreamNotFound(stream), NatsUdfCommon.ErrorHint("stream not found"));
                    }
                }

                if (lower.Contains("jetstream") && lower.Contains("not enabled"))
                    return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);

                return NatsUdfCommon.VerboseError(what, cause, NatsUdfCommon.ErrorHint(cause));
            }

            string message = NatsUdfAsync.SafeMessage(ex);
            string text = (message ?? string.Empty).ToLowerInvariant();
            if (text.Contains("jetstream") && text.Contains("not enabled"))
                return NatsUdfCommon.Fail(NatsUdfCommon.ErrorJetStreamDisabled);
            if (consumer != null && text.Contains("consumer") && text.Contains("not found"))
            {
                return NatsUdfCommon.VerboseError(
                    what, NatsUdfCommon.ConsumerNotFound(consumer, stream), NatsUdfCommon.ErrorHint("consumer not found"));
            }
            if (stream != null && text.Contains("stream") && text.Contains("not found"))
            {
                return NatsUdfCommon.VerboseError(
                    what, NatsUdfCommon.StreamNotFound(stream), NatsUdfCommon.ErrorHint("stream not found"));
            }

            return NatsUdfCommon.VerboseError(what, message, NatsUdfCommon.ErrorHint(message));
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
        // Small helpers
        // ==================================================================

        private static string MessageValue(StreamMsgGetResponse response)
        {
            if (response == null || response.Message == null)
                return NatsUdfCommon.NoValue;
            byte[] data = response.Message.Data.ToArray();
            return NatsResultFormatter.DecodePayload(data) ?? string.Empty;
        }

        private static string SubjectsCsv(ICollection<string> subjects)
        {
            return subjects == null || subjects.Count == 0 ? string.Empty : string.Join(",", subjects);
        }

        private static string StorageName(StreamConfig config)
        {
            if (config == null)
                return null;
            return config.Storage == StreamConfigStorage.Memory ? StorageMemory : StorageFile;
        }

        private static string NameOr(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }

        private static long ToLong(ulong value)
        {
            return value > long.MaxValue ? long.MaxValue : (long)value;
        }

        private static int Better(int value, int floor)
        {
            return value < floor ? floor : value;
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
        private sealed class JsTimeoutException : Exception
        {
            public JsTimeoutException(string what, int timeoutMs)
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
