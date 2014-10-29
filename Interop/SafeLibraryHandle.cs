using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace NKhil.Tools.Com2Reg.Interop
{
    internal class SafeLibraryHandle : IDisposable
    {
        #region DLL Imports

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string fileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        #endregion

        #region Fields

        private IntPtr m_hModule;
        private readonly string m_fileName;
        private readonly Dictionary<string, Delegate> m_functionDelegates = new Dictionary<string, Delegate>();
        
        #endregion

        #region Properties

        public bool IsDisposed
        {
            get { return m_hModule == IntPtr.Zero; }
        }

        protected IntPtr HModule
        {
            get { return m_hModule; }
        }

        public string FileName
        {
            get { return m_fileName; }
        }

        #endregion

        #region Constructors

        internal SafeLibraryHandle(IntPtr hModule)
        {
            m_hModule = hModule;
        }

        internal SafeLibraryHandle(string fileName)
        {
            m_fileName = fileName;

            m_hModule = LoadLibrary(fileName);
            if (m_hModule != IntPtr.Zero)
                return;

            int errorCode = Marshal.GetLastWin32Error();
            if (errorCode != 0)
                throw new Win32Exception(errorCode);
        }

        #endregion

        #region Public Methods

        internal Delegate GetFunctionDelegate<T>(string functionName)
        {
            if (m_hModule == IntPtr.Zero)
                throw new ObjectDisposedException(GetType().Name);

            Delegate functionDelegate;
            if (m_functionDelegates.TryGetValue(functionName, out functionDelegate))
                return functionDelegate;

            IntPtr ptr = GetProcAddress(m_hModule, functionName);
            if (ptr == IntPtr.Zero)
                throw new Win32Exception(Marshal.GetLastWin32Error());

            functionDelegate = Marshal.GetDelegateForFunctionPointer(ptr, typeof(T));
            m_functionDelegates.Add(functionName, functionDelegate);

            return functionDelegate;
        }

        #endregion

        #region Implementation of IDisposable

        ~SafeLibraryHandle()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (m_hModule == IntPtr.Zero)
                return;

            FreeLibrary(m_hModule);
            m_hModule = IntPtr.Zero;
        }

        #endregion
    }
}