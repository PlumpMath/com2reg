namespace NKhil.Tools.Com2Reg
{
    internal sealed class Com2RegOptions
    {
        internal string[] AssemblyNames { get; set; }

        internal string AssembliesPath { get; set; }

        internal string RegFileName { get; set; }

        internal bool SetCodeBase { get; set; }

        internal bool SilentMode { get; set; }
    }
}
