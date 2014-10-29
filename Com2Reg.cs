using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;
using System.Threading;
using NKhil.Tools.Com2Reg.CommandLine;

namespace NKhil.Tools.Com2Reg
{
    internal class Com2Reg
    {
        #region DLL Imports

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int SearchPath(string path, string fileName, string extension, int numBufferChars,
            StringBuilder buffer, int[] filePart);

        #endregion

        private static int Main(string[] args)
        {
            SetConsoleUI();

            int returnCode;
            Com2RegOptions options;
            if (!ParseArguments(args, out options, out returnCode))
                return returnCode;

            return Run(options);
        }

// ReSharper disable once InconsistentNaming
        private static void SetConsoleUI()
        {
            Thread currentThread = Thread.CurrentThread;
            currentThread.CurrentUICulture = CultureInfo.CurrentUICulture.GetConsoleFallbackUICulture();
            if (((Environment.OSVersion.Platform != PlatformID.Win32Windows) &&
                 (Console.OutputEncoding.CodePage != currentThread.CurrentUICulture.TextInfo.OEMCodePage)) &&
                (Console.OutputEncoding.CodePage != currentThread.CurrentUICulture.TextInfo.ANSICodePage))
            {
                currentThread.CurrentUICulture = new CultureInfo("en-US");
            }
        }

        private static bool ParseArguments(string[] args, out Com2RegOptions options, out int returnCode)
        {
            options = new Com2RegOptions();
            returnCode = 0;

            CommandLine.CommandLine commandLine;
            try
            {
                commandLine = new CommandLine.CommandLine(args, new[] {"*asmpath", "codebase", "*regfile", "silent", "help", "?"});
            }
            catch (ApplicationException exception)
            {
                Output.WriteError(exception);
                returnCode = 100;
                return false;
            }

            if (commandLine.Arguments.Count == 0 && commandLine.Options.Count == 0)
            {
                PrintUsage();
                returnCode = 0;
                return false;
            }

            options.AssemblyNames = commandLine.Arguments.ToArray();

            foreach (Option option in commandLine.Options)
            {
                if (option.Name.Equals("asmpath", StringComparison.OrdinalIgnoreCase))
                    options.AssembliesPath = option.Value;
                else if (option.Name.Equals("codebase", StringComparison.OrdinalIgnoreCase))
                    options.SetCodeBase = true;
                else if (option.Name.Equals("regfile", StringComparison.OrdinalIgnoreCase))
                    options.RegFileName = option.Value;
                else if (option.Name.Equals("silent", StringComparison.OrdinalIgnoreCase))
                    options.SilentMode = true;
                else if (option.Name.Equals("?") || option.Name.Equals("help"))
                {
                    PrintUsage();
                    returnCode = 0;
                    return false;
                }
                else
                {
                    Output.WriteError(string.Format("An invalid option has been specified: {0}", option.Name));
                    returnCode = 100;
                    return false;
                }
            }

            if (options.AssemblyNames.Length != 0)
                return true;

            Output.WriteError("No input files has been specified");
            returnCode = 100;
            return false;
        }

        private static int Run(Com2RegOptions options)
        {
            int returnCode;

            try
            {
                for (int i = 0; i < options.AssemblyNames.Length; i++)
                {
                    string originalAssemblyName = options.AssemblyNames[i];
                    string fullPath = Path.GetFullPath(originalAssemblyName);
                    if (File.Exists(fullPath))
                        options.AssemblyNames[i] = fullPath;
                    else
                    {
// ReSharper disable once InconsistentNaming
                        const int MAX_PATH = 260;
                        StringBuilder buffer = new StringBuilder(MAX_PATH + 1);

                        if (SearchPath(null, originalAssemblyName, null, buffer.Capacity + 1, buffer, null) == 0)
                        {
                            throw new ApplicationException(
                                string.Format("Unable to locate input assembly '{0}' or one of its dependencies.",
                                    originalAssemblyName));
                        }

                        options.AssemblyNames[i] = buffer.ToString();
                    }

                    options.AssemblyNames[i] = new FileInfo(options.AssemblyNames[i]).FullName;
                }

                if (!string.IsNullOrEmpty(options.RegFileName))
                    options.RegFileName = new FileInfo(options.RegFileName).FullName;

                if (string.IsNullOrEmpty(options.RegFileName))
                {
                    if (options.AssemblyNames.Length == 1)
                        options.RegFileName = Path.ChangeExtension(options.AssemblyNames[0], ".reg");
                    else
                        options.RegFileName = Path.Combine(Environment.CurrentDirectory, "com2reg.reg");
                }

                string regFileDirectoryName = Path.GetDirectoryName(options.RegFileName);
                if (!string.IsNullOrEmpty(regFileDirectoryName) && !Directory.Exists(regFileDirectoryName))
                    Directory.CreateDirectory(regFileDirectoryName);

                AppDomainSetup appDomainSetup = new AppDomainSetup
                {
                    ApplicationBase = Path.GetDirectoryName(options.AssemblyNames[0])
                };

                AppDomain domain = AppDomain.CreateDomain("com2reg", null, appDomainSetup);
                if (domain == null)
                    throw new ApplicationException("Error creating an app domain to perform the registration");

                ObjectHandle handle = domain.CreateInstanceFrom(Assembly.GetExecutingAssembly().CodeBase,
                    "NKhil.Tools.Com2Reg.RemoteCom2Reg");
                RemoteCom2Reg com2Reg = (RemoteCom2Reg) handle.Unwrap();

                returnCode = com2Reg.Run(options.AssemblyNames, options.AssembliesPath, options.RegFileName,
                    options.SetCodeBase, options.SilentMode);
            }
            catch (Exception ex)
            {
                Output.WriteError(ex);
                returnCode = 100;
            }

            return returnCode;
        }

        private static void PrintUsage()
        {
            Output.WriteInfo("Generates a registry script from one or more assemblies or native COM libraries");
            Output.WriteInfo("");
            Output.WriteInfo("Syntax: com2reg AssemblyName [AssemblyName [ ...]] [Options]");
            Output.WriteInfo("\nOptions:");
            Output.WriteInfo("    /regfile:FileName        Generate a reg file with the specified name.");
            Output.WriteInfo("    /codebase                Set the code base in the registry");
            Output.WriteInfo("    /asmpath:Directory       Look for assembly references here");
            Output.WriteInfo("    /silent                  Silent mode. Prevents displaying of");
            Output.WriteInfo("                             success messages");
            Output.WriteInfo("    /? or /help              Display this usage message");
            Output.WriteInfo("");
        }
    }
}