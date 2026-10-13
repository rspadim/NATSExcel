using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NATS.Client.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NatsExcel.Tools
{
    /// <summary>
    /// NatsValidate - validates a NATSExcel.json and (optionally) tests the
    /// connection using the SAME NATS.Net client the add-in uses, so every
    /// credential style is really exercised: token, user/password, NKey seed,
    /// JWT + seed, a creds file and mTLS.
    ///
    /// The PowerShell validator (installer/validate-config.ps1) checks the file
    /// against the schema without a client; this tool covers the connection
    /// half that needs the client libraries.
    ///
    /// Exit codes:
    ///   0 = OK (and, with --test, the server answered)
    ///   1 = config problems
    ///   2 = the config could not be read/parsed
    ///   3 = the endpoint could not be resolved (missing cluster/alias)
    ///   4 = the connection test failed
    /// </summary>
    internal static class Program
    {
        private const string ConfigFileName = "NATSExcel.json";
        private const string DefaultUrl = "nats://localhost:4222";

        private static async Task<int> Main(string[] args)
        {
            var opts = CommandLine.Parse(args);
            if (opts.Help)
            {
                CommandLine.PrintHelp();
                return 0;
            }

            Console.WriteLine("======================================================================");
            Console.WriteLine(" NATSExcel - validator (C#)");
            Console.WriteLine("======================================================================");

            string path = opts.ConfigPath ?? FirstExistingPath();
            if (path == null)
            {
                Console.WriteLine("Config : (none found)");
                Console.WriteLine();
                Console.WriteLine("No NATSExcel.json was found; the add-in uses its built-in defaults");
                Console.WriteLine("(" + DefaultUrl + "). Validating that default instead.");
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine("Config : " + Path.GetFullPath(path));
            }

            ConfigRoot config;
            var problems = new List<string>();
            try
            {
                config = ConfigLoader.Load(path, problems);
            }
            catch (Exception ex)
            {
                Console.WriteLine("RESULT: FAIL - the file could not be read.");
                Console.WriteLine("  " + ex.Message);
                return 2;
            }

            foreach (var p in problems)
                Console.WriteLine("  - " + p);

            // Resolve the default cluster to a URL list + effective options.
            ResolvedEndpoint endpoint;
            try
            {
                endpoint = ConfigLoader.Resolve(config, null);
            }
            catch (ConfigReferenceException ex)
            {
                Console.WriteLine();
                Console.WriteLine("RESULT: FAIL - " + ex.Message);
                return 3;
            }

            Console.WriteLine();
            Console.WriteLine("Connection");
            Console.WriteLine("  urls      : " + ConfigLoader.MaskUrls(endpoint.Urls));
            if (!string.IsNullOrEmpty(endpoint.ClusterLabel))
                Console.WriteLine("  cluster   : " + endpoint.ClusterLabel
                    + (string.IsNullOrEmpty(endpoint.ExpectCluster) ? "" : " (expectCluster='" + endpoint.ExpectCluster + "')"));
            Console.WriteLine("  auth      : " + DescribeAuth(endpoint.Auth));
            Console.WriteLine("  tls       : " + (endpoint.Auth.Tls ? "yes" : "no"));
            Console.WriteLine("  timeout   : " + endpoint.TimeoutMs + " ms");
            Console.WriteLine("  jetstream : " + (config.JetStreamEnabled ? "enabled (config)" : "disabled (config)"));

            if (problems.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("RESULT: FAIL - " + problems.Count + " config problem(s).");
                return 1;
            }

            if (!opts.Test)
            {
                Console.WriteLine();
                Console.WriteLine("RESULT: OK - the configuration is valid.");
                Console.WriteLine("(run with --test to also connect)");
                return 0;
            }

            Console.WriteLine();
            Console.WriteLine("Testing the connection (NATS.Net client)...");
            // `await` (not a blocking .GetAwaiter().GetResult()): a validator is a
            // console app with no synchronization context, and the vs-threading
            // analyzers promote an unbounded blocking wait to a build error.
            var result = await ConnectionTester.TestAsync(endpoint, opts.TimeoutMs).ConfigureAwait(false);
            foreach (var line in result.Lines)
                Console.WriteLine("  " + line);

            Console.WriteLine();
            if (!result.Ok)
            {
                Console.WriteLine("RESULT: FAIL - " + result.Summary);
                return 4;
            }

            Console.WriteLine("RESULT: OK - connected as " + DescribeAuth(endpoint.Auth) + ".");
            Console.WriteLine("        " + result.Summary);
            return 0;
        }

        private static string DescribeAuth(AuthConfig auth)
        {
            if (auth == null)
                return "none";
            if (!string.IsNullOrEmpty(auth.CredsFile)) return "creds file";
            if (!string.IsNullOrEmpty(auth.Jwt)) return "jwt + nkey seed";
            if (!string.IsNullOrEmpty(auth.NkeySeed)) return "nkey seed";
            if (!string.IsNullOrEmpty(auth.Token)) return "token";
            if (!string.IsNullOrEmpty(auth.User)) return "user/password";
            return "none";
        }

        private static string FirstExistingPath()
        {
            foreach (var candidate in CandidatePaths())
                if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
                    return candidate;
            return null;
        }

        private static IEnumerable<string> CandidatePaths()
        {
            string here = AppDomain.CurrentDomain.BaseDirectory;
            yield return Path.Combine(here, ConfigFileName);
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ConfigFileName);
            yield return Path.Combine(here, "..", ConfigFileName);
            string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (!string.IsNullOrEmpty(windows))
                yield return Path.Combine(windows, ConfigFileName);
        }
    }
}

