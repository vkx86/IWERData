using System.Collections.Generic;
using CommandLine;
// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace IWERDataCollector
{
    // ReSharper disable once ClassNeverInstantiated.Global
    public class CmdOptions
    {
        [Option('i', "inputdir", Required = true, HelpText = "Input directory containing IWER workbooks")]
        public string InputDir { get; set; }

        [Option('y', "year", Required = true, HelpText = "Year(s) of interest")]
        public IEnumerable<string> Year { get; set; }
    }
}
