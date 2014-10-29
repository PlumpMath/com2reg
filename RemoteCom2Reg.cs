using System;

namespace NKhil.Tools.Com2Reg
{
    internal class RemoteCom2Reg : MarshalByRefObject
    {
        internal int Run(string[] assemblyNames, string assembliesPath, string regFileName, bool setCodebase, bool silentMode)
        {
            Com2RegOptions options = new Com2RegOptions
            {
                AssemblyNames = assemblyNames,
                AssembliesPath = assembliesPath,
                RegFileName = regFileName,
                SetCodeBase = setCodebase,
                SilentMode = silentMode
            };

            return Com2RegCode.Run(options);
        }
    }
}