using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace UninstallRelatedProducts
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = Options.Parse(args);

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

                    logger.Log("Uninstall completed successfully");
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
