using System;
using System.Runtime.InteropServices;

namespace NKhil.Tools.Com2Reg.Interop
{
    internal class SafeComLibraryHandle : SafeLibraryHandle
    {
        #region Delegates

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = false)]
        private delegate int DllRegisterServerDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = false)]
        private delegate int DllUnregisterServerDelegate();

        #endregion

        #region Fields

        private DllRegisterServerDelegate m_dllRegisterServer;
        private DllUnregisterServerDelegate m_dllUnregisterServer;

        #endregion

        #region Constructors

        internal SafeComLibraryHandle(IntPtr hModule)
            : base(hModule)
        {
            Initialize();
        }

        internal SafeComLibraryHandle(string fileName)
            : base(fileName)
        {
            Initialize();
        }

        #endregion

        #region Public Methods

        public void DllRegisterServer()
        {
            if (IsDisposed)
                throw new ObjectDisposedException("");

            int hr = m_dllRegisterServer();
            Marshal.ThrowExceptionForHR(hr);
        }

        public void DllUnregisterServer()
        {
            if (IsDisposed)
                throw new ObjectDisposedException("");

            int hr = m_dllUnregisterServer();
            Marshal.ThrowExceptionForHR(hr);
        }

        #endregion

        #region Private Method

        private void Initialize()
        {
            m_dllRegisterServer =
                (DllRegisterServerDelegate) GetFunctionDelegate<DllRegisterServerDelegate>("DllRegisterServer");

            m_dllUnregisterServer =
                (DllUnregisterServerDelegate) GetFunctionDelegate<DllUnregisterServerDelegate>("DllUnregisterServer");
        }

        #endregion
    }
}