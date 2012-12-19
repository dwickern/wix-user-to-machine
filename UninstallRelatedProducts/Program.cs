using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace UninstallRelatedProducts
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var sw = File.CreateText(@"C:\uninstall.txt"))
            {
                sw.WriteLine("Args: " + args.Aggregate(string.Empty, (x, y) => x + y, _ => _));
                try
                {
                    var upgradeCode = Guid.Parse("6b124a3c-f3f9-480d-82f9-881e0b5663a1");
                    var productCodes = Msi.GetRelatedProducts(upgradeCode);
                    foreach (var product in productCodes)
                    {
                        sw.WriteLine("Uninstalling: " + product);
                        Console.WriteLine(product);
                        Msi.Uninstall(product, silent: true);
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
