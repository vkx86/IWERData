using System;
using System.Collections.Generic;
using CommandLine;

namespace IWERDataCollector
{
    internal static class Program
    {
        private static int _exitCode = -1;

        private static int Main(string[] args)
        {
            Parser.Default.ParseArguments<CmdOptions>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);

            return _exitCode;
        }

        private static void RunOptions(CmdOptions opts)
        {
            try
            {
                Console.WriteLine("Collecting data...\n");
                new WorkbooksDataCollector(opts.InputDir).CollectData(opts.Year);
                Console.WriteLine("Data collection completed, enjoy the day!");
                //Console.ReadKey();
                _exitCode = 0;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error occured while collecting data:");
                Console.WriteLine(e);
            }
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            foreach (var error in errs)
            {
                Console.WriteLine(error);
            }
        }
    }
}
