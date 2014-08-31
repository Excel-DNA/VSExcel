using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.ExcelDnaTools
{
    class AddInLoader
    {
        // Puts an .xll next to the .dll, together with a trivial .dna file, and loads as an add-in into Excel
        // CONSIDER: Keep track of what we change in the directory, so we can clean up when unloading.
        // CONSIDER: Should Excel-DNA support setting the BaseDirectory explicitly, so that we would not need to copy the .xll into the project output directory.
        // TODO: Put the templates into a resource, or build from scratch using the executing .xll, or something.
        // TODO: 64-bit support
        // TODO: Only do this the first time...?
        public static void RegisterDll(string addInPath)
        {
            var xllDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            var masterXllPath = Path.Combine(xllDirectory, "AddInXll.master");
            var masterDnaPath = Path.Combine(xllDirectory, "AddInDna.master");

            var addInDirectory = Path.GetDirectoryName(addInPath);
            var externalLibraryPath = Path.GetFileName(addInPath);
            var addInXllPath = Path.Combine(addInDirectory, Path.ChangeExtension(externalLibraryPath, "xll"));
            var addInDnaPath = Path.ChangeExtension(addInXllPath, "dna");

            if (!File.Exists(addInXllPath))
            {
                File.Copy(masterXllPath, addInXllPath, false);
            }
            if (!File.Exists(addInDnaPath))
            {
                File.Copy(masterDnaPath, addInDnaPath, false);
                var dnaContent = File.ReadAllText(addInDnaPath);
                dnaContent = dnaContent.Replace("%AddIn_Path%", externalLibraryPath);
                File.WriteAllText(addInDnaPath, dnaContent);
            }

            ExcelIntegration.RegisterXLL(addInXllPath);
        }
    }
}
