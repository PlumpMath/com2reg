using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32;

namespace NKhil.Tools.Com2Reg.Interop
{
    internal enum NativeRegistryOptions
    {
        Default = 0x0,
        RegistryView64 = 0x0100,
        RegistryView32 = 0x0200
    }

// ReSharper disable InconsistentNaming
    internal enum NativeRegistryHive : uint
    {
        HKEY_CLASSES_ROOT = 0x80000000,
        HKEY_CURRENT_USER = 0x80000001,
        HKEY_LOCAL_MACHINE = 0x80000002,
        HKEY_USERS = 0x80000003,
        HKEY_PERFORMANCE_DATA = 0x80000004,
        HKEY_PERFORMANCE_TEXT = 0x80000050,
        HKEY_PERFORMANCE_NLSTEXT = 0x80000060,
        HKEY_CURRENT_CONFIG = 0x80000005,
        HKEY_DYN_DATA = 0x80000006
    }
// ReSharper restore InconsistentNaming

    internal static class NativeRegistry
    {
        #region Static

        internal static readonly NativeRegistryKey ClassesRoot =
            new NativeRegistryKey(NativeRegistryHive.HKEY_CLASSES_ROOT);

        internal static readonly NativeRegistryKey CurrentConfig =
            new NativeRegistryKey(NativeRegistryHive.HKEY_CURRENT_CONFIG);

        internal static readonly NativeRegistryKey CurrentUser =
            new NativeRegistryKey(NativeRegistryHive.HKEY_CURRENT_USER);

        internal static readonly NativeRegistryKey DynData = new NativeRegistryKey(NativeRegistryHive.HKEY_DYN_DATA);

        internal static readonly NativeRegistryKey LocalMachine =
            new NativeRegistryKey(NativeRegistryHive.HKEY_LOCAL_MACHINE);

        internal static readonly NativeRegistryKey PerformanceData =
            new NativeRegistryKey(NativeRegistryHive.HKEY_PERFORMANCE_DATA);

        internal static readonly NativeRegistryKey Users = new NativeRegistryKey(NativeRegistryHive.HKEY_USERS);

        #endregion
    }

    internal sealed class NativeRegistryKey : IDisposable
    {
        #region Fields

        private readonly NativeRegistryOptions m_options;
        private readonly UIntPtr m_hKey;
        private string m_name;
        private bool m_isOpened;

        #endregion

        #region Properties

        internal UIntPtr Handle
        {
            get { return m_hKey; }
        }

        internal NativeRegistryOptions Options
        {
            get { return m_options; }
        }

        internal string Name
        {
            get { return m_name; }
            private set { m_name = value.Trim('\\'); }
        }

        internal string RealName
        {
            get { return NativeWinAPI.GetKeyNameByHandle(m_hKey); }
        }

        internal int SubKeyCount
        {
            get
            {
                int resultCode;
                int classLength = 0;
                int subKeyCount, maxSubKeyLength, maxClassLength, valueCount, maxValueNameLength, maxValueLength;
                if (
                    (resultCode =
                        NativeWinAPI.RegQueryInfoKey(m_hKey, null, ref classLength, IntPtr.Zero, out subKeyCount,
                            out maxSubKeyLength, out maxClassLength, out valueCount, out maxValueNameLength,
                            out maxValueLength, IntPtr.Zero, IntPtr.Zero)) != (int)NativeWinAPI.ErrorCode.SUCCESS)
                    throw new Win32Exception(resultCode);
                return subKeyCount;
            }
        }

        internal int ValueCount
        {
            get
            {
                int resultCode;
                int classLength = 0;
                int subKeyCount, maxSubKeyLength, maxClassLength, valueCount, maxValueNameLength, maxValueLength;
                if (
                    (resultCode =
                        NativeWinAPI.RegQueryInfoKey(m_hKey, null, ref classLength, IntPtr.Zero, out subKeyCount,
                            out maxSubKeyLength, out maxClassLength, out valueCount, out maxValueNameLength,
                            out maxValueLength, IntPtr.Zero, IntPtr.Zero)) != (int)NativeWinAPI.ErrorCode.SUCCESS)
                    throw new Win32Exception(resultCode);
                return valueCount;
            }
        }

        #endregion

        #region Constructors

        internal NativeRegistryKey(NativeRegistryHive registryHive, NativeRegistryOptions options = NativeRegistryOptions.Default)
        {
            m_name = Enum.GetName(typeof(NativeRegistryHive), registryHive);

            if (UIntPtr.Size == 8)
                m_hKey = new UIntPtr((ulong) (int) registryHive);
            else
                m_hKey = new UIntPtr((uint) registryHive);

            m_options = options;
            m_isOpened = false;
        }

        private NativeRegistryKey(UIntPtr hKey, NativeRegistryOptions options = NativeRegistryOptions.Default)
        {
            m_hKey = hKey;
            m_options = options;
            m_isOpened = m_hKey != UIntPtr.Zero;
        }

        #endregion

        #region internal Methods

        internal NativeRegistryKey Reopen(NativeRegistryOptions options)
        {
            return new NativeRegistryKey(m_hKey, options) { Name = m_name };
        }

        internal NativeRegistryKey OpenSubKey(string name)
        {
            return OpenSubKey(name, RegistryRights.ReadKey, m_options);
        }

        internal NativeRegistryKey OpenSubKey(string name, RegistryRights rights)
        {
            return OpenSubKey(name, rights, m_options);
        }

        internal NativeRegistryKey OpenSubKey(string name, NativeRegistryOptions options)
        {
            return OpenSubKey(name, RegistryRights.ReadKey, options);
        }

        internal NativeRegistryKey OpenSubKey(string name, RegistryRights rights, NativeRegistryOptions options)
        {
            int resultCode;
            UIntPtr hSubKey;
            if ((resultCode = NativeWinAPI.RegOpenKeyEx(m_hKey, name, 0, (int)rights | (int)options, out hSubKey)) !=
                (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            return new NativeRegistryKey(hSubKey, options) { Name = m_name + @"\" + name };
        }

        internal NativeRegistryKey CreateSubKey(string name, bool isVolatile = false)
        {
            return CreateSubKey(name, RegistryRights.FullControl, m_options, isVolatile);
        }

        internal NativeRegistryKey CreateSubKey(string name, RegistryRights rights, bool isVolatile = false)
        {
            return CreateSubKey(name, rights, m_options, isVolatile);
        }

        internal NativeRegistryKey CreateSubKey(string name, NativeRegistryOptions options, bool isVolatile = false)
        {
            return CreateSubKey(name, RegistryRights.FullControl, options, isVolatile);
        }

        internal NativeRegistryKey CreateSubKey(string name, RegistryRights rights, NativeRegistryOptions options, bool isVolatile = false)
        {
            UIntPtr hSubKey;
            int dwDisposition;
            int resultCode = NativeWinAPI.RegCreateKeyEx(m_hKey, name, 0, null, isVolatile ? 0x1 : 0x0,
                (int) rights | (int) options, IntPtr.Zero, out hSubKey, out dwDisposition);

            if (resultCode != (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            return new NativeRegistryKey(hSubKey, options) { Name = m_name + @"\" + name };
        }

        internal string[] GetSubKeyNames()
        {
            int resultCode;
            int classLength = 0;
            int subKeyCount, maxSubKeyLength, maxClassLength, valueCount, maxValueNameLength, maxValueLength;
            if (
                (resultCode =
                    NativeWinAPI.RegQueryInfoKey(m_hKey, null, ref classLength, IntPtr.Zero, out subKeyCount,
                        out maxSubKeyLength, out maxClassLength, out valueCount, out maxValueNameLength,
                        out maxValueLength, IntPtr.Zero, IntPtr.Zero)) != (int)NativeWinAPI.ErrorCode.SUCCESS)

                throw new Win32Exception(resultCode);

            if (subKeyCount == 0)
                return new string[0];

            int i = 0;
            List<string> subKeys = new List<string>(subKeyCount);
            while (resultCode == (int)NativeWinAPI.ErrorCode.SUCCESS)
            {
                int subKeyNameLength = maxSubKeyLength + 1;
                StringBuilder subKeyName = new StringBuilder(subKeyNameLength);
                resultCode = NativeWinAPI.RegEnumKeyEx(m_hKey, i++, subKeyName, ref subKeyNameLength, IntPtr.Zero, null,
                    IntPtr.Zero, IntPtr.Zero);

                if (resultCode != (int)NativeWinAPI.ErrorCode.SUCCESS)
                {
                    if (resultCode != (int)NativeWinAPI.ErrorCode.ERROR_NO_MORE_ITEMS)
                        throw new Win32Exception(resultCode);

                    break;
                }

                subKeys.Add(subKeyName.ToString());
            }

            subKeys.Sort();
            return subKeys.ToArray();
        }

        internal string[] GetValueNames()
        {
            int resultCode;
            int classLength = 0;
            int subKeyCount, maxSubKeyLength, maxClassLength, valueCount, maxValueNameLength, maxValueLength;
            if (
                (resultCode =
                    NativeWinAPI.RegQueryInfoKey(m_hKey, null, ref classLength, IntPtr.Zero, out subKeyCount,
                        out maxSubKeyLength, out maxClassLength, out valueCount, out maxValueNameLength,
                        out maxValueLength, IntPtr.Zero, IntPtr.Zero)) != (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            if (valueCount == 0)
                return new string[0];

            int i = 0;
            List<string> values = new List<string>(valueCount);
            while (resultCode == (int)NativeWinAPI.ErrorCode.SUCCESS)
            {
                RegistryValueKind type;
                int valueNameLength = maxValueNameLength + 1;
                StringBuilder valueName = new StringBuilder(valueNameLength);
                resultCode = NativeWinAPI.RegEnumValue(m_hKey, i++, valueName, ref valueNameLength, IntPtr.Zero,
                    out type, IntPtr.Zero, IntPtr.Zero);

                if (resultCode != (int)NativeWinAPI.ErrorCode.SUCCESS)
                {
                    if (resultCode != (int)NativeWinAPI.ErrorCode.ERROR_NO_MORE_ITEMS)
                        throw new Win32Exception(resultCode);

                    break;
                }

                values.Add(valueName.ToString());
            }

            values.Sort();
            return values.ToArray();
        }

        internal object GetValue(string name)
        {
            RegistryValueKind valueKind;
            return GetValue(name, out valueKind);
        }

        internal object GetValue(string name, object defaultValue)
        {
            RegistryValueKind valueKind;
            return GetValue(name, defaultValue, out valueKind);
        }

        internal RegistryValueKind GetValueKind(string name)
        {
            int resultCode;
            int dataSize = 0;
            RegistryValueKind valueKind;
            if (
                (resultCode =
                    NativeWinAPI.RegQueryValueEx(m_hKey, name, IntPtr.Zero, out valueKind, IntPtr.Zero, ref dataSize)) !=
                (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            return valueKind;
        }

        internal object GetValue(string name, out RegistryValueKind valueKind)
        {
            return GetValue(name, RegistryValueOptions.None, out valueKind);
        }

        internal object GetValue(string name, object defaultValue, out RegistryValueKind valueKind)
        {
            return GetValue(name, defaultValue, RegistryValueOptions.None, out valueKind);
        }

        internal object GetValue(string name, RegistryValueOptions options, out RegistryValueKind valueKind)
        {
            int resultCode;
            int dataSize = 0;
            if (
                (resultCode =
                    NativeWinAPI.RegQueryValueEx(m_hKey, name, IntPtr.Zero, out valueKind, IntPtr.Zero, ref dataSize)) !=
                (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            byte[] binaryData = new byte[dataSize];
            if (
                (resultCode =
                    NativeWinAPI.RegQueryValueEx(m_hKey, name, IntPtr.Zero, out valueKind, binaryData, ref dataSize)) !=
                (int)NativeWinAPI.ErrorCode.SUCCESS)
                throw new Win32Exception(resultCode);

            switch (valueKind)
            {
                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                case RegistryValueKind.MultiString:
                    string stringData = Encoding.Unicode.GetString(binaryData).TrimEnd('\0');

                    if (valueKind == RegistryValueKind.MultiString)
                        return stringData.Split('\0');

                    return valueKind == RegistryValueKind.String ||
                           options == RegistryValueOptions.DoNotExpandEnvironmentNames
                        ? stringData
                        : Environment.ExpandEnvironmentVariables(stringData);
                case RegistryValueKind.DWord:
                    return BitConverter.ToInt32(binaryData, 0);
                case RegistryValueKind.QWord:
                    return BitConverter.ToInt64(binaryData, 0);
                case RegistryValueKind.Binary:
                    return binaryData;
                default:
                    throw new Exception(string.Format("Unknown value kind: {0}", valueKind));
            }
        }

        internal object GetValue(string name, object defaultValue, RegistryValueOptions options,
            out RegistryValueKind valueKind)
        {
            int dataSize = 0;
            if ((NativeWinAPI.RegQueryValueEx(m_hKey, name, IntPtr.Zero, out valueKind, IntPtr.Zero, ref dataSize)) != (int)NativeWinAPI.ErrorCode.SUCCESS)
                return defaultValue;

            byte[] binaryData = new byte[dataSize];
            if ((NativeWinAPI.RegQueryValueEx(m_hKey, name, IntPtr.Zero, out valueKind, binaryData, ref dataSize)) != (int)NativeWinAPI.ErrorCode.SUCCESS)
                return defaultValue;

            switch (valueKind)
            {
                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                case RegistryValueKind.MultiString:
                    string stringData = Encoding.Unicode.GetString(binaryData).TrimEnd('\0');

                    if (valueKind == RegistryValueKind.MultiString)
                        return stringData.Split('\0');

                    return valueKind == RegistryValueKind.String ||
                           options == RegistryValueOptions.DoNotExpandEnvironmentNames
                        ? stringData
                        : Environment.ExpandEnvironmentVariables(stringData);
                case RegistryValueKind.DWord:
                    return BitConverter.ToInt32(binaryData, 0);
                case RegistryValueKind.QWord:
                    return BitConverter.ToInt64(binaryData, 0);
                case RegistryValueKind.Binary:
                    return binaryData;
                default:
                    return defaultValue;
            }
        }

        internal void Close()
        {
            if (m_isOpened)
            {
                NativeWinAPI.RegCloseKey(m_hKey);
                m_isOpened = false;
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Close();
        }

        #endregion
    }

    internal static class NativeWinAPI
    {
        #region Enums

        internal enum KeyInformationClass
        {
            KeyBasicInformation = 0,
            KeyNodeInformation = 1,
            KeyFullInformation = 2,
            KeyNameInformation = 3,
            KeyCachedInformation = 4,
            KeyFlagsInformation = 5,
            KeyVirtualizationInformation = 6,
            KeyHandleTagsInformation = 7,
            MaxKeyInfoClass = 8
        }

// ReSharper disable InconsistentNaming
        internal enum ErrorCode : uint
        {
            SUCCESS = 0,
            ERROR_NO_MORE_ITEMS = 259,
            STATUS_BUFFER_TOO_SMALL = 0xC0000023,
            STATUS_BUFFER_OVERFLOW = 0x80000005,
        }
// ReSharper restore InconsistentNaming

        #endregion

        #region internal Static Fields

        internal static readonly string CurrentUserSid = GetCurrentUserSid(".default");
        internal static readonly string CurrentUserClasses = CurrentUserSid + "_classes";

        #endregion

        #region internal External Methods

        [DllImport("ntdll.dll", SetLastError = true)]
        internal static extern int NtQueryKey(UIntPtr keyHandle, KeyInformationClass keyInformationClass,
            IntPtr keyInformation, int bufferSize, out int resultSize);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegOpenKeyEx(UIntPtr hKey, string lpSubKey, uint ulOptions, int samDesired,
            out UIntPtr phkResult);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegCreateKeyEx(UIntPtr hKey, string lpSubKey, uint reserved, string lpClass,
            int dwOptions, int samDesired,
            IntPtr lpSecurityAttributes, out UIntPtr phkResult, out int lpdwDisposition);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegDeleteKeyEx(UIntPtr hKey, string lpSubKey, int samDesired, int reserved);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegQueryInfoKey(UIntPtr hKey, string lpClass, ref int lpcbClass,
            IntPtr lpReservedMustBeZero, out int lpcSubKeys, out int lpcbMaxSubKeyLen, out int lpcbMaxClassLen,
            out int lpcValues, out int lpcbMaxValueNameLen, out int lpcbMaxValueLen, IntPtr lpcbSecurityDescriptor,
            IntPtr lpftLastWriteTime);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegEnumKeyEx(UIntPtr hKey, int dwIndex, StringBuilder lpName, ref int lpcbName,
            IntPtr lpReserved, string lpClass, IntPtr lpcbClass, IntPtr lpftLastWriteTime);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        internal static extern int RegEnumValue(UIntPtr hKey, int dwIndex, StringBuilder valueName, ref int lpcchValueName,
            IntPtr reserved, out RegistryValueKind type, IntPtr data, IntPtr dataLength);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueEx")]
        internal static extern int RegQueryValueEx(UIntPtr hKey,
            string valueName, IntPtr reserved, out RegistryValueKind type,
            IntPtr zero, ref int dataSize);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueEx")]
        internal static extern int RegQueryValueEx(UIntPtr hKey,
            string valueName, IntPtr reserved, out RegistryValueKind type,
            [Out] byte[] data, ref int dataSize);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueEx")]
        internal static extern int RegQueryValueEx(UIntPtr hKey,
            string valueName, IntPtr reserved, out RegistryValueKind type,
            ref int data, ref int dataSize);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegCloseKey(UIntPtr hKey);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern int RegOverridePredefKey(UIntPtr hkey, UIntPtr hNewKey);

        #endregion

        #region Common internal Static Methods

        internal static string GetKeyNameByHandle(UIntPtr hKey)
        {
            NativeRegistryHive registryHive = (NativeRegistryHive)hKey.ToUInt64();
            if (Enum.IsDefined(typeof(NativeRegistryHive), registryHive))
                return Enum.GetName(typeof(NativeRegistryHive), registryHive);

            int ptrSize = 0;
            IntPtr ptr = Marshal.AllocHGlobal(ptrSize);

            ErrorCode resultCode =
                (ErrorCode)NtQueryKey(hKey, KeyInformationClass.KeyNameInformation, ptr, ptrSize, out ptrSize);
            Marshal.FreeHGlobal(ptr);

            if (resultCode != ErrorCode.STATUS_BUFFER_TOO_SMALL &&
                resultCode != ErrorCode.STATUS_BUFFER_OVERFLOW)
                throw new Win32Exception((int)resultCode);

            ptr = Marshal.AllocHGlobal(ptrSize);
            resultCode =
                (ErrorCode)NtQueryKey(hKey, KeyInformationClass.KeyNameInformation, ptr, ptrSize, out ptrSize);

            string keyName = null;
            if (resultCode == ErrorCode.SUCCESS)
            {
                int keyNameLength = Marshal.ReadInt32(ptr) / sizeof(char);
                IntPtr keyNamePtr = new IntPtr(ptr.ToInt64() + sizeof(int));
                keyName = Marshal.PtrToStringUni(keyNamePtr, keyNameLength);
            }

            Marshal.FreeHGlobal(ptr);
            if (resultCode != ErrorCode.SUCCESS)
                Marshal.ThrowExceptionForHR((int)resultCode);

            keyName = FromWin32Path(keyName);
            return keyName;
        }

        #endregion

        #region Private Static Methods

        private static string GetCurrentUserSid(string defaultValue)
        {
            WindowsIdentity currentIdentity = WindowsIdentity.GetCurrent();
            return currentIdentity == null || currentIdentity.User == null
                ? defaultValue
                : currentIdentity.User.Value;
        }

        private static string FromWin32Path(string win32RegistryPath)
        {
            string keyName = win32RegistryPath;
            LinkedList<string> keyNameParts =
                new LinkedList<string>(keyName.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries));
            if (keyNameParts.First == null ||
                !keyNameParts.First.Value.Equals(@"registry", StringComparison.InvariantCultureIgnoreCase))
                return keyName;

            keyNameParts.RemoveFirst();
            if (keyNameParts.First != null)
            {
                if (keyNameParts.First.Value.Equals(@"user", StringComparison.InvariantCultureIgnoreCase))
                {
                    keyNameParts.RemoveFirst();
                    if (keyNameParts.First == null)
                        keyNameParts.AddFirst(@"HKEY_USERS");
                    else if (keyNameParts.First.Value.Equals(CurrentUserSid, StringComparison.InvariantCultureIgnoreCase))
                    {
                        keyNameParts.RemoveFirst();
                        keyNameParts.AddFirst(@"HKEY_CURRENT_USER");
                    }
                    else if (keyNameParts.First.Value.Equals(CurrentUserClasses,
                        StringComparison.InvariantCultureIgnoreCase))
                    {
                        keyNameParts.RemoveFirst();
                        keyNameParts.AddFirst(@"HKEY_CLASSES_ROOT");
                    }
                    else
                        keyNameParts.AddFirst(@"HKEY_USERS");
                }
                else if (keyNameParts.First.Value.Equals(@"machine", StringComparison.InvariantCultureIgnoreCase))
                {
                    keyNameParts.RemoveFirst();
                    if (keyNameParts.First != null && keyNameParts.First.Next != null &&
                        keyNameParts.First.Value.Equals(@"SOFTWARE", StringComparison.InvariantCultureIgnoreCase) &&
                        keyNameParts.First.Next.Value.Equals(@"Classes", StringComparison.InvariantCultureIgnoreCase))
                    {
                        keyNameParts.RemoveFirst();
                        keyNameParts.RemoveFirst();
                    }
                    else
                        keyNameParts.AddFirst("HKEY_LOCAL_MACHINE");
                }
            }

            string[] keyNamePartsArray = new string[keyNameParts.Count];
            keyNameParts.CopyTo(keyNamePartsArray, 0);
            return string.Join(@"\", keyNamePartsArray);
        }

        #endregion
    }
}