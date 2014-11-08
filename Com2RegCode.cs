using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using NKhil.Tools.Com2Reg.CommandLine;
using NKhil.Tools.Com2Reg.Interop;

namespace NKhil.Tools.Com2Reg
{
    internal static class Com2RegCode
    {
        internal static int Run(Com2RegOptions options)
        {
            Output.SetSilent(options.SilentMode);

            int returnCode = 0;

            List<Assembly> assembliesToRegister = new List<Assembly>();
            List<SafeComLibraryHandle> comLibrariesToRegister = new List<SafeComLibraryHandle>();
            try
            {
// ReSharper disable RedundantAssignment
// ReSharper disable once NotAccessedVariable
                AssemblyResolver assemblyResolver;
                if (!string.IsNullOrEmpty(options.AssembliesPath))
                    assemblyResolver = new AssemblyResolver(options.AssembliesPath);
// ReSharper restore RedundantAssignment

                foreach (string assemblyName in options.AssemblyNames)
                {
                    SafeComLibraryHandle comLibrary = null;
                    try
                    {
                        comLibrary = new SafeComLibraryHandle(assemblyName);
                        comLibrariesToRegister.Add(comLibrary);
                        continue;
                    }
                    catch (Exception)
                    {
                        if (comLibrary != null)
                            comLibrary.Dispose();
                    }

                    Assembly assembly;
                    try
                    {
                        if (!string.IsNullOrEmpty(options.AssembliesPath))
                            assembly = Assembly.ReflectionOnlyLoadFrom(assemblyName);
                        else
                            assembly = Assembly.LoadFrom(assemblyName);
                    }
                    catch (BadImageFormatException ex)
                    {
                        throw new ApplicationException(
                            string.Format("Failed to load '{0}' because it is neither a valid .NET assembly, nor a native {1}-bit COM library", assemblyName, IntPtr.Size * 8),
                            ex);
                    }
                    catch (FileNotFoundException ex)
                    {
                        throw new ApplicationException(
                            string.Format("Unable to locate input assembly '{0}' or one of its dependencies.",
                                assemblyName), ex);
                    }

                    if (string.Compare(assemblyName, options.RegFileName, StringComparison.OrdinalIgnoreCase) == 0)
                        throw new ApplicationException("Registry file would overwrite the input file");

                    if (options.SetCodeBase && (assembly.GetName().GetPublicKey().Length == 0))
                    {
                        Output.WriteWarning(
                            "Registering an unsigned assembly with /codebase can cause your assembly to interfere with other applications that may be installed on the same computer. The /codebase switch is intended to be used only with signed assemblies. Please give your assembly a strong name and re-register it.");
                        Output.WriteWarning(string.Format("Please consider signing the following assembly: {0}",
                            assemblyName));
                    }

                    assembliesToRegister.Add(assembly);
                }


                if (GenerateRegScript(assembliesToRegister, comLibrariesToRegister, options))
                    Output.WriteInfo(string.Format("Registry script '{0}' generated successfully", options.RegFileName));
                else
                    Output.WriteWarning("No registry script will be produced since there are no types to register");
            }
            catch (ReflectionTypeLoadException ex)
            {
                Output.WriteError("The following exceptions were thrown while loading the types in the assembly:");

                Exception[] loaderExceptions = ex.LoaderExceptions;
                for (int i = 0; i < loaderExceptions.Length; i++)
                {
                    try
                    {
                        Output.WriteError(string.Format("Exception[{0}] = {1}", i, loaderExceptions[i]));
                    }
                    catch (Exception innerException)
                    {
                        Output.WriteError(string.Format("Exception[{0}] ==>> {1}", i, innerException));
                    }
                }

                returnCode = 100;
            }
            catch (Exception ex)
            {
                Output.WriteError(ex);
                returnCode = 100;
            }
            finally
            {
                foreach (SafeComLibraryHandle safeHandle in comLibrariesToRegister)
                    safeHandle.Dispose();
            }

            return returnCode;
        }

        private static bool GenerateRegScript(
            IEnumerable<Assembly> assembliesToRegister,
            IEnumerable<SafeComLibraryHandle> comLibrariesToRegister,
            Com2RegOptions options)
        {
            const string overrideRootKeyName = @"Software\NKhil\Com2Reg";

            RegistrationServices regAsm = new RegistrationServices();

            try
            {
                using (new SafeRegistryOverride(NativeRegistry.ClassesRoot,
                    NativeRegistry.CurrentUser.CreateSubKey(overrideRootKeyName + @"\HKEY_CLASSES_ROOT")))
                using (new SafeRegistryOverride(NativeRegistry.LocalMachine,
                    NativeRegistry.CurrentUser.CreateSubKey(overrideRootKeyName + @"\HKEY_LOCAL_MACHINE")))
                using (new SafeRegistryOverride(NativeRegistry.CurrentUser,
                    NativeRegistry.CurrentUser.CreateSubKey(overrideRootKeyName + @"\HKEY_CURRENT_USER")))
                {
                    RegisterAssemblies(regAsm, assembliesToRegister, options.SetCodeBase);

                    using (var oleAut32 = new OleAut32LibraryHandle())
                    {
                        oleAut32.OaEnablePerUserTLibRegistration();

                        RegisterComLibraries(comLibrariesToRegister);
                    }
                }

                using (RegistryKey overrideRootKey = Registry.CurrentUser.OpenSubKey(overrideRootKeyName))
                    return GenerateRegScript(overrideRootKey, options.RegFileName);
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(overrideRootKeyName);
            }
        }

        private static void RegisterComLibraries(IEnumerable<SafeComLibraryHandle> comLibrariesToRegister)
        {
            foreach (SafeComLibraryHandle comLibrary in comLibrariesToRegister)
            {
                try
                {
                    comLibrary.DllRegisterServer();
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(
                        string.Format("An error occurred during registration of COM library: {0}", comLibrary.FileName),
                        ex);
                }
            }
        }

        private static void RegisterAssemblies(RegistrationServices regAsm, IEnumerable<Assembly> assemblies, bool setCodeBase)
        {
            AssemblyRegistrationFlags flags = setCodeBase
                ? AssemblyRegistrationFlags.SetCodeBase
                : AssemblyRegistrationFlags.None;

            foreach (Assembly assembly in assemblies)
            {
                try
                {
                    if (!regAsm.RegisterAssembly(assembly, flags))
                    {
                        Output.WriteWarning(string.Format("the assembly contains no eligible types: {0}",
                            assembly.CodeBase));
                    }
                }
                catch (TargetInvocationException ex)
                {
                    throw new ApplicationException(
                        string.Format("An error occurred inside the user defined Register functions in {0}", assembly.CodeBase),
                        ex);
                }
            }
        }

        private static bool GenerateRegScript(RegistryKey rootKey, string regFileName)
        {
            bool regScriptGenerated;
            StringBuilder sb = new StringBuilder();
            using (TextWriter regScriptWriter = new StringWriter(sb))
            {
                regScriptWriter.WriteLine("REGEDIT4");
                regScriptWriter.WriteLine();

                regScriptGenerated = WriteKeyToRegFile(regScriptWriter, rootKey, rootKey.Name);
            }

            if (regScriptGenerated)
            {
                using (TextWriter regFileWriter = new StreamWriter(regFileName, false, new UTF8Encoding(false)))
                    regFileWriter.Write(sb.ToString());
            }

            return regScriptGenerated;
        }

        private static string EscapeRegistryString(string s)
        {
            return s.Replace(@"\", @"\\").Replace("\"", "\\\"");
        }

        private static bool WriteKeyToRegFile(TextWriter regScriptWriter, RegistryKey key, string rootKeyName)
        {
            bool isEmpty = true;

            if (key.Name != rootKeyName)
            {
                string overridenKeyName = key.Name.Substring(rootKeyName.Length + 1);

                bool isHive = !overridenKeyName.Contains(@"\");
                bool writeKey = key.ValueCount != 0 || (key.SubKeyCount == 0 && !isHive);
                if (writeKey)
                {
                    isEmpty = false;

                    regScriptWriter.WriteLine("[{0}]", overridenKeyName);

                    foreach (string valueName in key.GetValueNames())
                    {
                        string quotedValueName = !string.IsNullOrEmpty(valueName)
                            ? string.Format("\"{0}\"", EscapeRegistryString(valueName))
                            : "@";

                        string valueTypeAndData;
                        RegistryValueKind valueKind = key.GetValueKind(valueName);
                        switch (valueKind)
                        {
                            case RegistryValueKind.String:
                                string stringData = (string) key.GetValue(valueName, string.Empty);
                                valueTypeAndData = string.Format("\"{0}\"", EscapeRegistryString(stringData));
                                break;

                            case RegistryValueKind.DWord:
                                int intData = (int) key.GetValue(valueName, 0);
                                valueTypeAndData = string.Format("dword:{0:X8}", intData);
                                break;

                            case RegistryValueKind.MultiString:
                                string[] multiStringData = (string[]) key.GetValue(valueName, new string[0]);

                                string formatedMultiStringData = string.Join(
                                    ",",
                                    multiStringData.Select(s => string.Format("\"{0}\"", EscapeRegistryString(s)))
                                        .ToArray());

                                valueTypeAndData = string.Format("multiStringData:{0}", formatedMultiStringData);
                                break;

                            case RegistryValueKind.Binary:
                                byte[] binaryData = (byte[]) key.GetValue(valueName, new byte[0]);
                                string formattedBinaryData = string.Join(
                                    ",", binaryData.Select(b => b.ToString("X2")).ToArray());

                                valueTypeAndData = string.Format("multiStringData:{0}", formattedBinaryData);
                                break;

                            default:
                                throw new ApplicationException(
                                    string.Format(
                                        "Registry value kind is not supported! Key: '{0}', Value: '{1}', Kind: {2}",
                                        overridenKeyName, valueName, valueKind));
                        }

                        regScriptWriter.WriteLine("{0}={1}", quotedValueName, valueTypeAndData);
                    }

                    regScriptWriter.WriteLine();
                }
            }

            foreach (string subKeyName in key.GetSubKeyNames())
            {
                using (RegistryKey childKey = key.OpenSubKey(subKeyName))
                {
                    if (WriteKeyToRegFile(regScriptWriter, childKey, rootKeyName))
                        isEmpty = false;
                }
            }

            return !isEmpty;
        }
    }
}
