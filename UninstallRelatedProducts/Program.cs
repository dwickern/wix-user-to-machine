using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using CommandLine;
using CommandLine.Text;

namespace UninstallRelatedProducts
{
    class Program
    {
        class Options : CommandLineOptionsBase
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

        static void Main(string[] args)
        {
            var options = new Options();
            if (!CommandLineParser.Default.ParseArguments(args, options))
            {
                // Failed to parse arguments
                // Arguments help text is already printed to stdout
                Environment.Exit(-1);
            }

            using (var logger = new Logger(options.LogFilePath))
            {
                try
                {
                    // Additional arguments parsing
                    var upgradeCode = Guid.Parse(options.UpgradeCode);
                    var maxVersion = default(Version);
                    if (!string.IsNullOrEmpty(options.MaxVersion))
                        maxVersion = Version.Parse(options.MaxVersion);

                    var productCodes = Msi.GetRelatedProducts(upgradeCode).ToList();
                    logger.Log("Number of related products found: " + productCodes.Count);

                    foreach (var product in productCodes)
                    {
                        logger.Log("Product code: " + product);

                        if (maxVersion != null)
                        {
                            var version = Msi.GetVersion(product);
                            logger.Log("Product version: " + version);
                            if (version > maxVersion)
                            {
                                logger.Log("Skipping.");
                                continue;
                            }
                        }

                        if (options.PerUserOnly)
                        {
                            var allUsers = Msi.IsAllUsers(product);
                            logger.Log("All users: " + allUsers);
                            if (allUsers)
                            {
                                Console.WriteLine("Skipping.");
                                continue;
                            }
                        }

                        logger.Log("Uninstalling. Silently: " + options.Silent);
                        Msi.Uninstall(product, options.Silent);
                    }
                }
                catch (Exception e)
                {
                    logger.Log(e);
                    throw;
                }
            }
        }
    }
}
