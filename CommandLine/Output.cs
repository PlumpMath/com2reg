using System;

namespace NKhil.Tools.Com2Reg.CommandLine
{
    internal static class Output
    {
        #region Static

        private static bool s_silent;

        #endregion

        #region Internal Methods

        internal static void SetSilent(bool silent)
        {
            s_silent = silent;
        }

        internal static void WriteError(Exception e)
        {
            WriteError(e.ToString());
        }

        internal static void WriteError(string message)
        {
            WriteError(message, 0);
        }

        internal static void WriteError(string message, int errorNumber)
        {
            const string format = "com2reg : error CTR{1:0000} : {0}";
            Console.Error.WriteLine(format, message, errorNumber);
        }

        internal static void WriteInfo(string message)
        {
            if (!s_silent)
                Console.WriteLine(message);
        }

        internal static void WriteWarning(string message)
        {
            if (s_silent)
                return;

            const string format = "com2reg : warning CTR0000 : {0}";
            Console.Error.WriteLine(format, message);
        }

        #endregion
    }
}
