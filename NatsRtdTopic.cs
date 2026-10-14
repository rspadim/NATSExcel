using System;
using System.Collections.Generic;
using System.Globalization;

namespace NATSExcel
{
    /// <summary>
    /// The nine frozen RTD commands (DESIGN.md §6, PLAN.md §1.2) plus
    /// <see cref="Unknown"/>.
    /// <para>
    /// <see cref="Unknown"/> doubles as the result kind for a topic whose parse
    /// failed for any reason: <see cref="NatsRtdTopicSpec.Error"/> then carries
    /// the exact <c>#ERROR: ...</c> text the RTD server should surface. A caller
    /// therefore checks <see cref="NatsRtdTopicSpec.IsValid"/> (or the error
    /// text) before acting on a spec, never the kind alone.
    /// </para>
    /// </summary>
    internal enum NatsRtdKind
    {
        /// <summary><c>SUB</c> — subscribe (Core NATS, wildcards allowed).</summary>
        Sub,

        /// <summary><c>KV</c> — watch a JetStream Key/Value key.</summary>
        Kv,

        /// <summary><c>KVGET</c> — polled Key/Value read.</summary>
        KvGet,

        /// <summary><c>KVKEYS</c> — polled Key/Value key list.</summary>
        KvKeys,

        /// <summary><c>REQ</c> — request/reply issued every poll tick.</summary>
        Req,

        /// <summary><c>JS</c> — consume an existing JetStream consumer.</summary>
        Js,

        /// <summary><c>MSG</c> — polled "last message matching filter".</summary>
        Msg,

        /// <summary><c>OBJ</c> — watch Object Store metadata.</summary>
        Obj,

        /// <summary><c>RTDCLUSTER</c> — live cluster-health cell.</summary>
        Cluster,

        /// <summary>An unrecognized command, or any malformed topic.</summary>
        Unknown
    }

    /// <summary>
    /// Parsed form of a single <c>=RTD("NatsRtd", , &lt;command&gt;, ...)</c> topic.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This type is deliberately free of any Excel-DNA or COM dependency so it can
    /// be linked into the unit-test project (PLAN.md S3.2); <c>NatsRtd.cs</c> is
    /// the only RTD file allowed to touch Excel.
    /// </para>
    /// <para>
    /// The four <see cref="Arg1"/>..<see cref="Arg4"/> fields are the values of the
    /// documented positions, verbatim (only the command token is trimmed and
    /// upper-cased). A value's leading/trailing whitespace is significant in NATS
    /// — <c>"orders"</c> and <c>"orders "</c> are different subjects — so the
    /// parser never normalizes the value arguments. Position 3 of <c>REQ</c> is the
    /// one exception: it holds the clamped <c>everyMs</c> as invariant text, with
    /// the same integer exposed as <see cref="EveryMs"/>.
    /// </para>
    /// </remarks>
    internal sealed class NatsRtdTopicSpec
    {
        /// <summary>Minimum accepted <c>REQ</c> interval (PLAN.md §1.2): a smaller
        /// value would let a worksheet flood the server, so it is clamped up.</summary>
        public const int MinEveryMs = 100;

        /// <summary>
        /// Maximum UTF-8 byte length of one RTD slot (subject, queue group, bucket,
        /// filter, endpoint...). The parser is the only bound the RTD path has, and
        /// an over-long slot would exceed the server's <c>max_control_line</c>
        /// (default 4096 bytes) - a fatal server error. 2048 leaves room for the
        /// `SUB `/reply-queue prefix while refusing the pathological token.
        /// </summary>
        public const int MaxSlotBytes = 2048;

        /// <summary>
        /// The command list as it appears in the unknown-command error. The order is
        /// frozen (tests assert the exact text) and includes <c>RTDCLUSTER</c>, the
        /// ninth command (DESIGN.md §6 / §7.9).
        /// </summary>
        public const string CommandList = "[SUB, KV, KVGET, KVKEYS, REQ, JS, MSG, OBJ, RTDCLUSTER]";

        /// <summary>The parsed command; <see cref="Unknown"/> on any error.</summary>
        public NatsRtdKind Kind { get; private set; }

        /// <summary>
        /// The command token, trimmed and upper-cased (<c>ToUpperInvariant</c>).
        /// For an unknown command this still holds what the user typed so the
        /// error message and the log line can name it.
        /// </summary>
        public string Command { get; private set; }

        /// <summary>Position 1 of the command (subject / bucket / stream / …), or null.</summary>
        public string Arg1 { get; private set; }

        /// <summary>Position 2 (queueGroup / key / payload / consumer / filter / object / endpoint), or null.</summary>
        public string Arg2 { get; private set; }

        /// <summary>Position 3 (endpoint / everyMs / filter), or null.</summary>
        public string Arg3 { get; private set; }

        /// <summary>Position 4 (REQ interval? / JS endpoint), or null.</summary>
        public string Arg4 { get; private set; }

        /// <summary>
        /// The clamped <c>REQ</c> interval in milliseconds (0 for every other kind).
        /// Always <see cref="MinEveryMs"/> or greater when the topic is valid.
        /// </summary>
        public int EveryMs { get; private set; }

        /// <summary>
        /// The exact parse error, or null when the topic is well formed. The RTD
        /// server renders it as <c>#ERROR: &lt;Error&gt;</c>.
        /// </summary>
        public string Error { get; private set; }

        /// <summary>True when the topic parsed without error.</summary>
        public bool IsValid
        {
            get { return Error == null; }
        }

        /// <summary>
        /// The endpoint (cluster reference or URL) argument of the parsed kind, or
        /// null when the topic omits it. Convenience over the positional args, whose
        /// endpoint slot differs per command (PLAN.md §1.2): <c>SUB</c>/<c>KV</c>/
        /// <c>KVGET</c>/<c>MSG</c>/<c>OBJ</c> use position 3; <c>REQ</c>/<c>JS</c> use
        /// position 4; <c>KVKEYS</c>/<c>RTDCLUSTER</c> use position 2.
        /// </summary>
        public string Endpoint
        {
            get
            {
                switch (Kind)
                {
                    case NatsRtdKind.Sub:
                    case NatsRtdKind.Kv:
                    case NatsRtdKind.KvGet:
                    case NatsRtdKind.Msg:
                    case NatsRtdKind.Obj:
                        return Arg3;
                    case NatsRtdKind.Req:
                    case NatsRtdKind.Js:
                        return Arg4;
                    case NatsRtdKind.KvKeys:
                    case NatsRtdKind.Cluster:
                        return Arg2;
                    default:
                        return null;
                }
            }
        }

        private NatsRtdTopicSpec(
            NatsRtdKind kind,
            string command,
            string arg1,
            string arg2,
            string arg3,
            string arg4,
            int everyMs,
            string error)
        {
            Kind = kind;
            Command = command;
            Arg1 = arg1;
            Arg2 = arg2;
            Arg3 = arg3;
            Arg4 = arg4;
            EveryMs = everyMs;
            Error = error;
        }

        /// <summary>
        /// Parses one RTD topic into a <see cref="NatsRtdTopicSpec"/>. This method
        /// never throws on a malformed topic; it returns a spec with
        /// <see cref="Kind"/> = <see cref="NatsRtdKind.Unknown"/> and a non-null
        /// <see cref="Error"/> instead, so the RTD server can show
        /// <c>#ERROR: &lt;Error&gt;</c> without a try/catch.
        /// </summary>
        /// <param name="command">
        /// The command token (usually <c>topicInfo[0]</c>). When blank it is read
        /// from <paramref name="topicInfo"/>[0].
        /// </param>
        /// <param name="topicInfo">
        /// The RTD topic argument list, 1-based in the documented grammar: index 0
        /// is the command, index 1 is the first value, and so on. A caller that
        /// passes the value arguments only (command omitted from the list) is also
        /// accepted: the parser skips index 0 when it does not match
        /// <paramref name="command"/>. Reported topic positions are still the
        /// formula positions (the command is position 1), so the first value is
        /// always "the second argument".
        /// </param>
        /// <remarks>
        /// <para><b>Frozen grammar.</b> The command token is
        /// <c>ToUpperInvariant().Trim()</c>. Positions are fixed:</para>
        /// <list type="bullet">
        /// <item><description><c>SUB [1]subject [2]queueGroup [3]endpoint</c></description></item>
        /// <item><description><c>KV [1]bucket [2]key [3]endpoint</c></description></item>
        /// <item><description><c>KVGET [1]bucket [2]key [3]endpoint</c></description></item>
        /// <item><description><c>KVKEYS [1]bucket [2]endpoint</c></description></item>
        /// <item><description><c>REQ [1]subject [2]payload [3]everyMs [4]endpoint</c></description></item>
        /// <item><description><c>JS [1]stream [2]consumer [3]filter [4]endpoint</c></description></item>
        /// <item><description><c>MSG [1]stream [2]filter [3]endpoint</c></description></item>
        /// <item><description><c>OBJ [1]bucket [2]object [3]endpoint</c></description></item>
        /// <item><description><c>RTDCLUSTER [1]what [2]endpoint</c></description></item>
        /// </list>
        /// <para><b>Optional-middle positional rule.</b> <c>SUB</c>'s queue group,
        /// <c>JS</c>'s filter and <c>MSG</c>'s filter are optional tokens that sit
        /// between a leading value and the endpoint. Positions are fixed, so a filled
        /// middle token is always the middle value (spell out every token — including
        /// an empty middle placeholder — to reach the endpoint). The one exception the
        /// rule allows: when the middle token is blank (an empty placeholder, or the
        /// argument omitted) and the endpoint slot is filled, the trailing token is
        /// used as the <b>endpoint only if it "resolves as an endpoint"</b>, and as
        /// the middle value otherwise. Because this parser is config-free (it cannot
        /// consult <c>Servers</c> aliases), "resolves as an endpoint" is the
        /// deterministic stand-in: the token contains <c>://</c>, or it is a single
        /// non-blank token with no whitespace. A middle value that contains a space
        /// (a queue group such as <c>"my group"</c>) therefore stays in the middle
        /// slot. Ambiguity is inherent for a single-word token; spelling out every
        /// token is the documented workaround (PLAN.md §1.2).</para>
        /// <para><c>RTDCLUSTER</c> has no middle token (<c>what</c> is leading), so
        /// its positions stay strictly fixed: a lone token is <c>what</c>, never an
        /// endpoint.</para>
        /// <para>A non-blank token past the documented arity is an error; a
        /// blank/whitespace token is ignored because Excel leaves trailing empty
        /// arguments behind.</para>
        /// </remarks>
        internal static NatsRtdTopicSpec Parse(string command, IList<string> topicInfo)
        {
            string cmd = NormalizeToken(command);

            // topicInfo[0] is the command in the canonical RTD shape; use it when
            // the caller did not pass the command explicitly, and skip it when it
            // matches so only the value tokens are read positionally.
            if (cmd.Length == 0 && topicInfo != null && topicInfo.Count > 0)
                cmd = NormalizeToken(topicInfo[0]);
            int firstValue = (topicInfo != null && topicInfo.Count > 0 && NormalizeToken(topicInfo[0]) == cmd) ? 1 : 0;
            int valueCount = topicInfo == null ? 0 : Math.Max(0, topicInfo.Count - firstValue);

            NatsRtdKind kind;
            int arity;
            switch (cmd)
            {
                case "SUB": kind = NatsRtdKind.Sub; arity = 3; break;
                case "KV": kind = NatsRtdKind.Kv; arity = 3; break;
                case "KVGET": kind = NatsRtdKind.KvGet; arity = 3; break;
                case "KVKEYS": kind = NatsRtdKind.KvKeys; arity = 2; break;
                case "REQ": kind = NatsRtdKind.Req; arity = 4; break;
                case "JS": kind = NatsRtdKind.Js; arity = 4; break;
                case "MSG": kind = NatsRtdKind.Msg; arity = 3; break;
                case "OBJ": kind = NatsRtdKind.Obj; arity = 3; break;
                case "RTDCLUSTER": kind = NatsRtdKind.Cluster; arity = 2; break;
                default:
                    return Fail(cmd, "unknown command '" + cmd + "', expected one of " + CommandList);
            }

            // Fill the documented slots; anything past them is checked below.
            var slots = new string[arity];
            for (int v = 0; v < arity && v < valueCount; v++)
                slots[v] = topicInfo[firstValue + v];

            // Required values are rejected before the extra-argument check (so a
            // topic is refused before it can ever be registered and retried).
            string required = RequireSlots(cmd, kind, slots);
            if (required != null)
                return Fail(cmd, required);

            // Only a NON-BLANK token past the documented arity is an error: Excel
            // leaves trailing empty arguments on =RTD(...) calls.
            for (int v = arity; v < valueCount; v++)
            {
                string extra = topicInfo[firstValue + v];
                if (!string.IsNullOrWhiteSpace(extra))
                    return Fail(cmd, cmd + " accepts at most " + arity
                        + " argument(s); unexpected extra argument '" + extra
                        + "' at topic position " + (v + 2));
            }

            // Resolve the optional middle slot (see the <remarks> above).
            if (kind == NatsRtdKind.Sub)
                ResolveOptionalMiddle(slots, 1, 2);   // [1]subject [2]queueGroup [3]endpoint
            else if (kind == NatsRtdKind.Js)
                ResolveOptionalMiddle(slots, 2, 3);   // [1]stream [2]consumer [3]filter [4]endpoint
            else if (kind == NatsRtdKind.Msg)
                ResolveOptionalMiddle(slots, 1, 2);   // [1]stream [2]filter [3]endpoint

            int everyMs = 0;
            if (kind == NatsRtdKind.Req)
            {
                long raw;
                if (!long.TryParse(slots[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out raw))
                    return Fail(cmd, "REQ requires everyMs as a positive number of milliseconds");
                long clamped = raw < MinEveryMs ? MinEveryMs : raw;
                if (clamped > int.MaxValue) clamped = int.MaxValue;
                everyMs = (int)clamped;
                // Position 3 exposes the effective (clamped) interval, not the raw
                // token, so a spec built from "50" and one built from "100" behave
                // identically downstream.
                slots[2] = everyMs.ToString(CultureInfo.InvariantCulture);
            }

            return new NatsRtdTopicSpec(
                kind,
                cmd,
                slots[0],
                arity > 1 ? slots[1] : null,
                arity > 2 ? slots[2] : null,
                arity > 3 ? slots[3] : null,
                everyMs,
                null);
        }

        /// <summary>
        /// Validates the required values of the parsed kind. Returns the exact error
        /// message for the first blank required value, or null when every required
        /// value is present. The required-noun vocabulary is frozen: subject, bucket,
        /// key, stream, consumer, object.
        /// </summary>
        private static string RequireSlots(string cmd, NatsRtdKind kind, string[] slots)
        {
            switch (kind)
            {
                case NatsRtdKind.Sub:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "subject", 2);
                    break;
                case NatsRtdKind.Kv:
                case NatsRtdKind.KvGet:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "bucket", 2);
                    if (IsBlank(slots[1])) return RequiredMessage(cmd, "key", 3);
                    break;
                case NatsRtdKind.KvKeys:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "bucket", 2);
                    break;
                case NatsRtdKind.Req:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "subject", 2);
                    break;
                case NatsRtdKind.Js:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "stream", 2);
                    if (IsBlank(slots[1])) return RequiredMessage(cmd, "consumer", 3);
                    break;
                case NatsRtdKind.Msg:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "stream", 2);
                    break;
                case NatsRtdKind.Obj:
                    if (IsBlank(slots[0])) return RequiredMessage(cmd, "bucket", 2);
                    if (IsBlank(slots[1])) return RequiredMessage(cmd, "object", 3);
                    break;
                // RTDCLUSTER has no required value: both `what` and the endpoint are optional.
            }

            // Every non-blank slot lands in the SUB/JS control line, so cap its
            // UTF-8 byte length here (the parser is the only gate the RTD path has;
            // the UDF path uses RequireSubject/RequireText). A 32767-char subject
            // from one cell could otherwise exceed the server's max_control_line.
            for (int i = 0; i < slots.Length; i++)
            {
                if (!IsBlank(slots[i]) && System.Text.Encoding.UTF8.GetByteCount(slots[i]) > MaxSlotBytes)
                    return cmd + " argument " + (i + 2) + " is too long (max " + MaxSlotBytes + " bytes)";
            }
            return null;
        }

        /// <summary>
        /// Applies the optional-middle positional rule (see the <see cref="Parse"/>
        /// remarks). It fires only when the optional-middle slot is blank and the
        /// endpoint slot is not: the user left the middle value out (an empty
        /// placeholder or a shorter argument list), so the single trailing token is
        /// the endpoint when it resolves as one and the middle value otherwise.
        /// When the middle slot is filled, positions are fixed and nothing moves —
        /// <c>SUB(subject, "workers")</c> stays a queue group, not an endpoint.
        /// </summary>
        private static void ResolveOptionalMiddle(string[] slots, int middleIndex, int endpointIndex)
        {
            if (!IsBlank(slots[middleIndex]) || IsBlank(slots[endpointIndex]))
                return;
            // The middle was left blank; the trailing token is the endpoint only if
            // it resolves as one, otherwise it is the middle value.
            if (LooksLikeEndpoint(slots[endpointIndex]))
                return;
            slots[middleIndex] = slots[endpointIndex];
            slots[endpointIndex] = null;
        }

        /// <summary>
        /// The config-free stand-in for "resolves as an endpoint" (PLAN.md §1.2):
        /// a URL (contains <c>://</c>) or a single non-blank token with no
        /// whitespace (a <c>Servers</c> alias). A middle value with a space therefore
        /// never masquerades as an endpoint. This parser cannot consult
        /// <c>Servers</c>, so the alias case is approximated exactly this way.
        /// </summary>
        private static bool LooksLikeEndpoint(string token)
        {
            if (IsBlank(token))
                return false;
            if (token.IndexOf("://", StringComparison.Ordinal) >= 0)
                return true;
            for (int i = 0; i < token.Length; i++)
                if (char.IsWhiteSpace(token[i]))
                    return false;
            return true;
        }

        /// <summary>
        /// Builds the frozen required-value message:
        /// <c>{CMD} requires a {noun} as the {ordinal} argument</c>. The article is
        /// always "a" so the template stays mechanical and the tests can substitute
        /// the noun without a grammar special case.
        /// </summary>
        private static string RequiredMessage(string cmd, string noun, int topicPosition)
        {
            return cmd + " requires a " + noun
                + " as the " + Ordinal(topicPosition) + " argument";
        }

        /// <summary>Ordinal of a 1-based topic position, for the required-value message.</summary>
        private static string Ordinal(int topicPosition)
        {
            switch (topicPosition)
            {
                case 1: return "first";
                case 2: return "second";
                case 3: return "third";
                case 4: return "fourth";
                default: return topicPosition.ToString(CultureInfo.InvariantCulture);
            }
        }

        /// <summary>Builds a failed spec: <see cref="NatsRtdKind.Unknown"/> plus the exact error text.</summary>
        private static NatsRtdTopicSpec Fail(string cmd, string error)
        {
            return new NatsRtdTopicSpec(NatsRtdKind.Unknown, cmd, null, null, null, null, 0, error);
        }

        private static bool IsBlank(string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        private static string NormalizeToken(string token)
        {
            return (token ?? string.Empty).Trim().ToUpperInvariant();
        }
    }
}
