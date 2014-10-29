using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NKhil.Tools.Com2Reg
{
    internal class AssemblyResolver
    {
        #region Fields

        private readonly string[] m_assemblyPaths;

        #endregion

        #region Constructors

        internal AssemblyResolver(string assemblyPaths)
        {
            m_assemblyPaths = !string.IsNullOrEmpty(assemblyPaths)
                ? assemblyPaths.Split(new[] {';'})
                : null;

            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += ResolveAssembly;
        }

        #endregion

        #region Private Methods

        private Assembly ResolveAssembly(object sender, ResolveEventArgs args)
        {
            string assemblyName = args.Name;

            Assembly assembly;
            if (m_assemblyPaths == null)
                assembly = Assembly.ReflectionOnlyLoad(AppDomain.CurrentDomain.ApplyPolicy(assemblyName));
            else
            {
                AssemblyName name = new AssemblyName(assemblyName);
                assemblyName = name.Name;

                string dllAssemblyFileName = assemblyName + ".dll";
                assembly = m_assemblyPaths.Select(path => path + @"\" + dllAssemblyFileName)
                    .Where(File.Exists)
                    .Select(Assembly.ReflectionOnlyLoadFrom)
                    .FirstOrDefault();

                if (assembly == null)
                {
                    string exeAssemblyFileName = assemblyName + ".exe";
                    assembly = m_assemblyPaths.Select(lstPath => lstPath + @"\" + exeAssemblyFileName)
                        .Where(File.Exists)
                        .Select(Assembly.ReflectionOnlyLoadFrom)
                        .FirstOrDefault();
                }
            }

            if (assembly == null)
                throw new FileNotFoundException(assemblyName);

            return assembly;
        }

        #endregion
    }
}
