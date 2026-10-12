using System;

namespace NatsExcel.Tools
{
    internal sealed class CommandLine
    {
        public string ConfigPath { get; private set; }
        public bool Test { get; private set; }
        public bool Help { get; private set; }
        public int TimeoutMs { get; private set; }

        public static CommandLine Parse(string[] args)
        {
            var o = new CommandLine();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                switch (a)
                {
                    case "--test":
                    case "-t":
                        o.Test = true;
                        break;
                    case "--help":
                    case "-h":
                    case "/?":
                        o.Help = true;
                        break;
                    case "--config":
                    case "-c":
                        if (i + 1 < args.Length) o.ConfigPath = args[++i];
                        break;
                    case "--timeout":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out var ms)) o.TimeoutMs = ms;
                        break;
                    default:
                        // A bare path argument names the config file.
                        if (!a.StartsWith("-") && !a.StartsWith("/")) o.ConfigPath = a;
                        break;
                }
            }
            return o;
        }

        public static void PrintHelp()
        {
            Console.WriteLine("NatsValidate - validate a NATSExcel.json and test the connection.");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  NatsValidate [config.json] [--test] [--timeout <ms>]");
            Console.WriteLine();
            Console.WriteLine("  (no --config)  search: this folder, %USERPROFILE%, .., C:\\Windows");
            Console.WriteLine("  --test, -t     also connect with the NATS.Net client");
            Console.WriteLine("  --timeout, ms  override the connect/timeout in milliseconds");
            Console.WriteLine();
            Console.WriteLine("Exit codes: 0 OK, 1 config problems, 2 unreadable, 3 dangling reference, 4 connection failed.");
        }
    }
}
