using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
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
            using (var sw = File.CreateText(@"C:\uninstall.txt"))
            {
                try
                {
                    var options = new Options();
                    if (!CommandLineParser.Default.ParseArguments(args, options))
                    {
                        // Failed to parse arguments
                        // Arguments help text is already printed to stdout
                        Environment.Exit(-1);
                    }

                    var upgradeCode = Guid.Parse(options.UpgradeCode);
                    var productCodes = Msi.GetRelatedProducts(upgradeCode);
                    foreach (var product in productCodes)
                    {
                        sw.WriteLine("Uninstalling: " + product);
                        Console.WriteLine(product);
                        Msi.Uninstall(product, options.Silent);
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine(e.ToString());
                    Console.WriteLine(e.ToString());
                    throw;
                }
            }
        }
    }
}
