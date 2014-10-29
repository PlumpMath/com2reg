using System;
using System.ComponentModel;

namespace NKhil.Tools.Com2Reg.Interop
{
    internal sealed class SafeRegistryOverride : IDisposable
    {
        #region Fields

        private NativeRegistryKey m_sourceKey;
        private NativeRegistryKey m_destinationKey;
        private readonly bool m_disposeDestinationKey;

        #endregion

        #region Constructors

        internal SafeRegistryOverride(NativeRegistryKey sourceKey, NativeRegistryKey destinationKey, bool disposeDestinationKey = true)
        {
            m_sourceKey = sourceKey;
            m_destinationKey = destinationKey;
            m_disposeDestinationKey = disposeDestinationKey;

            int errorCode = NativeWinAPI.RegOverridePredefKey(sourceKey.Handle, destinationKey.Handle);
            if (errorCode != (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(errorCode);
        }

        #endregion

        #region Implementation of IDisposable

        ~SafeRegistryOverride()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (m_sourceKey != null)
                NativeWinAPI.RegOverridePredefKey(m_sourceKey.Handle, UIntPtr.Zero);

            if (disposing && m_disposeDestinationKey && m_destinationKey != null)
            {
                m_destinationKey.Dispose();
                m_destinationKey = null;
            }
        }

        #endregion
    }
}
