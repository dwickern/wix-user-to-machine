using System.Linq;
using System.Collections.Generic;
using System;
using CommandLine;
using CommandLine.Text;

namespace UninstallRelatedProducts
{
    /// <summary>
    /// Contains command-line options
    /// </summary>
    public class Options : CommandLineOptionsBase
    {
        [Option(null, "upgradecode", Required = true, HelpText = "Upgrade code of the products to uninstall.")]
        public string UpgradeCode { get; set; }

        [Option(null, "peruseronly", DefaultValue = false,
            HelpText = "Whether to only uninstall per-user products. " +
                "If not specified, both per-user and per-machine products will be uninstalled.")]
        public bool PerUserOnly { get; set; }

        [Option(null, "maxversion",
            HelpText = "Maximum product version to uninstall. " +
                "Products with a version greater than this number will be skipped. " +
                "If not specified, all related products will be uninstalled.")]
        public string MaxVersion { get; set; }

        [Option(null, "log", HelpText = "Path to a log file to write progress information.")]
        public string LogFilePath { get; set; }

        [Option(null, "quiet", DefaultValue = false, HelpText = "Quiet mode, no user interaction")]
        public bool Silent { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, c => HelpText.DefaultParsingErrorsHandler(this, c));
        }
    }
}