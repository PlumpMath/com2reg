using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace NKhil.Tools.Com2Reg.Interop
{
    internal class OleAut32LibraryHandle : SafeLibraryHandle
    {
        #region Delegates

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = false)]
        private delegate void OaEnablePerUserTLibRegistrationDelegate();

        #endregion

        #region Fields

        private OaEnablePerUserTLibRegistrationDelegate m_oaEnablePerUserTLibRegistration;

        #endregion

        #region Constructors

        internal OleAut32LibraryHandle()
            : base("OleAut32.dll")
        {
            Initialize();
        }

        #endregion

        #region Public Methods

        public void OaEnablePerUserTLibRegistration()
        {
            if (IsDisposed)
                throw new ObjectDisposedException("");

            if (m_oaEnablePerUserTLibRegistration != null)
                m_oaEnablePerUserTLibRegistration();
        }

        #endregion

        #region Private Method

        private void Initialize()
        {
            try
            {
                m_oaEnablePerUserTLibRegistration =
                    (OaEnablePerUserTLibRegistrationDelegate)GetFunctionDelegate<OaEnablePerUserTLibRegistrationDelegate>(
                        "OaEnablePerUserTLibRegistration");
            }
            catch (Win32Exception ex)
            {
                const int ERROR_PROC_NOT_FOUND = 127;
                if (ex.ErrorCode == ERROR_PROC_NOT_FOUND)
                    return;

                throw;
            }
        }

        #endregion
    }
}
