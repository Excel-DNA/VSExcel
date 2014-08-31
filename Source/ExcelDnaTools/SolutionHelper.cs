using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System.IO;

namespace ExcelDna.ExcelDnaTools
{
    // Follows the '*Helper' anti-pattern
    class SolutionHelper
    {
        readonly ExcelDnaToolsPackage _package;
        readonly DTE _dte;
        public SolutionHelper(ExcelDnaToolsPackage package)
        {
            _package = package;
            var packageServiceProvider = (IServiceProvider)package;
            _dte = packageServiceProvider.GetService(typeof(DTE)) as DTE;
        }

        public string GetActiveTargetName()
        {
            // Getting the output path from the build: http://stackoverflow.com/questions/5626303/how-do-i-get-the-output-directories-from-the-last-build
            // and: http://stackoverflow.com/questions/5486593/getting-the-macro-value-of-projects-targetpath-via-dte
            // and: http://social.msdn.microsoft.com/Forums/vstudio/en-US/b6ec12c6-a0e7-419d-b43b-143a3cb4c326/getting-the-macro-value-of-projects-targetpath-via-dte?forum=vsx

            // TODO: This might not work for the default VB configuration, where the Advanced Configuration is not on.
            //       From  http://social.msdn.microsoft.com/Forums/vstudio/en-US/03d9d23f-e633-4a27-9b77-9029735cfa8d/how-to-get-the-right-output-path-from-envdteproject-by-code-if-show-advanced-build?forum=vsx
            //          If the Advanced configuration is disabled, the detection of output path is context sensitive during the build operation. 
            //          e.g. If you disable the advanced configuration in VS and do an F5 the active configuration selected is debug and if you 
            //          do a project build from the project node its release. So the VS shell picks up the configuration information in this case 
            //          based on context of the operation. The host application could do a similiar context detection from the dte object and 
            //          select the appropriate ActiveConfiguration.

            // Alternative is to get the IVsSolution of the SVsSolution service...?
            Array projs = (Array)_dte.ActiveSolutionProjects;
            if (projs.Length == 0)
                return null;

            var vsProject = (EnvDTE.Project)projs.GetValue(0);
            var fullPath = vsProject.Properties.Item("FullPath").Value.ToString();
            var outputPath = vsProject.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value.ToString();
            var outputDir = Path.Combine(fullPath, outputPath);
            var outputFileName = vsProject.Properties.Item("OutputFileName").Value.ToString();
            var assemblyPath = Path.Combine(outputDir, outputFileName);
            return assemblyPath;
        }

        public string GetActiveStartArguments()
        {
            Array projs = (Array)_dte.ActiveSolutionProjects;
            if (projs.Length == 0)
                return null;

            var vsProject = (EnvDTE.Project)projs.GetValue(0);
            var debugConfiguration = vsProject.ConfigurationManager.Cast<Configuration>().FirstOrDefault(c => c.ConfigurationName == "Debug");
            if (debugConfiguration != null)
            {
                var startArguments = debugConfiguration.Properties.Item("StartArguments").Value.ToString();
                // TODO: Check that this is valid
                return startArguments;
            }
            // TODO: Message / Log about things being wrong
            return null;
        }

    }
}
