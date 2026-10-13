using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// The unbounded-blocking tripwire (PLAN V7.3 / golden rule 10 / risk R3):
    /// the NATS .NET client is async-only, so a synchronous wait on the Excel
    /// thread freezes Excel. <c>NatsConnectionManager.WaitBounded*</c> is the ONE
    /// bounded-blocking primitive; every other <c>.GetAwaiter().GetResult()</c>,
    /// <c>.Result</c> or <c>.Wait(</c> in the root production sources is a defect.
    ///
    /// This is a pure source scan (no build, no NATS, no Excel) so it runs
    /// everywhere and cannot be defeated by a runtime path that never executes.
    /// The matching analyzer tripwire is
    /// <c>Microsoft.VisualStudio.Threading.Analyzers</c> (VSTHRD002), wired in
    /// <c>NATSExcel.csproj</c>/.editorconfig; this test is the belt-and-braces
    /// scan that also catches a construct the analyzer may not resolve.
    ///
    /// Comments and string literals are ignored so prose (e.g. the XML doc in
    /// <c>NatsUdfCore.cs</c> that names the very pattern) never false-positives.
    /// Only the root <c>*.cs</c> files are scanned: <c>test\</c> and
    /// <c>tools\</c> legitimately block (their <c>Main</c> is not the Excel
    /// thread) and are excluded.
    /// </summary>
    public class NoUnboundedBlockingTests
    {
        /// <summary>
        /// The one file allowed to contain a synchronous wait, and only inside its
        /// bounded-wait helpers. Any other file is a failure.
        /// </summary>
        private const string BoundedWaitFile = "NatsConnectionManager.cs";

        /// <summary>
        /// The bounded-wait helper method names in
        /// <see cref="BoundedWaitFile"/>. A blocking call is allowed only between
        /// the declaration of one of these and the matching close brace.
        /// </summary>
        private static readonly string[] BoundedWaitMethods =
        {
            "WaitBounded",
            "WaitBoundedResult",
        };

        [Fact]
        public void RootProductionSources_ContainNoUnboundedBlockingOutsideTheBoundedWaitHelper()
        {
            string repoRoot = FindRepoRoot();
            string[] sources = Directory.GetFiles(repoRoot, "*.cs", SearchOption.TopDirectoryOnly);

            var offenders = new List<string>();
            foreach (string source in sources)
            {
                string fileName = Path.GetFileName(source);
                string[] lines = File.ReadAllLines(source);

                // A masking pass that blanks comments and string/char literals, so
                // a blocking token in prose or a message is not matched. Column
                // positions are preserved so the report can cite file:line.
                string[] masked = MaskCommentsAndStrings(lines);

                // Only the bounded-wait file may block, and only inside the bodies
                // of its bounded-wait helper methods. Compute those line ranges.
                HashSet<int> allowedLines = string.Equals(
                    fileName, BoundedWaitFile, StringComparison.OrdinalIgnoreCase)
                        ? BoundedWaitBodyLines(masked)
                        : new HashSet<int>();

                for (int i = 0; i < masked.Length; i++)
                {
                    string call = FirstBlockingCall(masked[i]);
                    if (call != null && !allowedLines.Contains(i))
                        offenders.Add(fileName + ":" + (i + 1) + " (" + call + ")");
                }
            }

            Assert.True(offenders.Count == 0,
                "unbounded blocking found outside the bounded-wait helper (golden rule 10, "
                + "PLAN V7.3). Route it through NatsConnectionManager.WaitBounded* or await "
                + "it asynchronously: " + string.Join(", ", offenders));
        }

        /// <summary>
        /// Pins the allow-list itself: changing <see cref="BoundedWaitFile"/>
        /// without moving the primitive is a review-visible failure, and the file
        /// must still really contain the helper.
        /// </summary>
        [Fact]
        public void TheOnlyBoundedWaitImplementation_StillLivesInNatsConnectionManager()
        {
            Assert.Equal("NatsConnectionManager.cs", BoundedWaitFile);

            string path = Path.Combine(FindRepoRoot(), BoundedWaitFile);
            Assert.True(File.Exists(path), "the bounded-wait file vanished: " + path);
            string text = File.ReadAllText(path);
            Assert.Contains("internal static bool WaitBounded", text);
            Assert.Contains("internal static T WaitBoundedResult", text);
        }

        /// <summary>The three blocking shapes the scan rejects.</summary>
        private static readonly string[] BlockingCalls =
        {
            ".GetAwaiter().GetResult()",
            ".Result",
            ".Wait(",
        };

        /// <summary>
        /// Returns the first blocking call on a masked line, or null. A blocking
        /// token must be a whole member access: <c>.ResultX</c> or
        /// <c>.WaitAsync(</c> are not matches (a word-boundary check follows the
        /// token), so the scan does not false-positive on a similarly named member.
        /// </summary>
        private static string FirstBlockingCall(string maskedLine)
        {
            foreach (string call in BlockingCalls)
            {
                int index = maskedLine.IndexOf(call, StringComparison.Ordinal);
                if (index < 0)
                    continue;

                int after = index + call.Length;
                bool boundary = after >= maskedLine.Length
                    || !(char.IsLetterOrDigit(maskedLine[after]) || maskedLine[after] == '_');
                if (boundary)
                    return call.TrimEnd('(', '.');
            }
            return null;
        }

        /// <summary>
        /// The set of line indices (0-based) that live inside the body of a
        /// <see cref="BoundedWaitMethods"/> helper in the bounded-wait file. Brace
        /// depth is counted from each helper's signature line so a blocking call is
        /// allowed only between its opening and closing braces.
        /// </summary>
        private static HashSet<int> BoundedWaitBodyLines(string[] maskedLines)
        {
            var allowed = new HashSet<int>();
            for (int i = 0; i < maskedLines.Length; i++)
            {
                bool isHelperSignature = BoundedWaitMethods.Any(
                    method => maskedLines[i].Contains(method));
                if (!isHelperSignature)
                    continue;

                // A helper's signature may span lines (attributes/generics); find
                // the first '{' at or after the signature that has no matching
                // '}' yet, then walk to its matching close brace.
                int openIndex = -1;
                for (int j = i; j < maskedLines.Length; j++)
                {
                    if (maskedLines[j].IndexOf('{') >= 0)
                    {
                        openIndex = j;
                        break;
                    }
                    // Stop at the next declaration end if no brace was found (a
                    // signature without a body cannot contain a call).
                    if (maskedLines[j].TrimEnd().EndsWith(";"))
                        break;
                }
                if (openIndex < 0)
                    continue;

                int depth = 0;
                for (int j = openIndex; j < maskedLines.Length; j++)
                {
                    foreach (char c in maskedLines[j])
                    {
                        if (c == '{')
                            depth++;
                        else if (c == '}')
                            depth--;
                    }
                    allowed.Add(j);
                    if (depth <= 0)
                        break;
                }
            }

            return allowed;
        }

        /// <summary>
        /// Replaces every character inside a line comment, a block comment, a
        /// string literal or a char literal with a space, leaving code and the
        /// original line/column layout (and therefore line numbers) intact.
        /// </summary>
        private static string[] MaskCommentsAndStrings(string[] lines)
        {
            var result = new string[lines.Length];
            bool inBlockComment = false;

            for (int i = 0; i < lines.Length; i++)
            {
                var output = new StringBuilder(lines[i].Length);
                for (int j = 0; j < lines[i].Length; j++)
                {
                    char c = lines[i][j];
                    char next = j + 1 < lines[i].Length ? lines[i][j + 1] : '\0';

                    if (inBlockComment)
                    {
                        output.Append(' ');
                        if (c == '*' && next == '/')
                        {
                            output.Append(' ');
                            j++;
                            inBlockComment = false;
                        }
                        continue;
                    }

                    // Line comment.
                    if (c == '/' && next == '/')
                    {
                        output.Append(' ', lines[i].Length - j);
                        break;
                    }

                    // Block comment.
                    if (c == '/' && next == '*')
                    {
                        output.Append("  ");
                        j++;
                        inBlockComment = true;
                        continue;
                    }

                    // Verbatim string (@"...) - may span lines and escape quotes by
                    // doubling them.
                    if (c == '@' && next == '"')
                    {
                        output.Append("  ");
                        j++;
                        while (j + 1 < lines[i].Length)
                        {
                            j++;
                            char sc = lines[i][j];
                            output.Append(' ');
                            if (sc == '"')
                            {
                                if (j + 1 < lines[i].Length && lines[i][j + 1] == '"')
                                {
                                    j++;
                                    output.Append(' ');
                                    continue; // escaped quote, stay in string
                                }
                                break; // closing quote
                            }
                        }
                        continue;
                    }

                    // Regular string or char literal (single line unless escaped).
                    if (c == '"' || c == '\'')
                    {
                        char quote = c;
                        output.Append(' ');
                        while (j + 1 < lines[i].Length)
                        {
                            j++;
                            char sc = lines[i][j];
                            output.Append(' ');
                            if (sc == '\\')
                            {
                                if (j + 1 < lines[i].Length)
                                {
                                    j++;
                                    output.Append(' ');
                                }
                                continue; // escaped char, still inside literal
                            }
                            if (sc == quote)
                                break; // closing quote
                        }
                        continue;
                    }

                    output.Append(c);
                }

                result[i] = output.ToString();
            }

            return result;
        }

        /// <summary>
        /// Walks up from the test binary until the repository root, anchored by
        /// <c>NATSExcel.csproj</c> (there is no solution file yet, PLAN F0.1).
        /// Mirrors <see cref="ProjectFileParityTests"/>.
        /// </summary>
        private static string FindRepoRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "NATSExcel.csproj")))
                    return directory.FullName;
                directory = directory.Parent;
            }
            throw new InvalidOperationException(
                "NATSExcel.csproj was not found above " + AppContext.BaseDirectory);
        }
    }
}
