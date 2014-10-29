using System;
using System.Globalization;

namespace NKhil.Tools.Com2Reg.CommandLine
{
    internal class Abbrevs
    {
        #region Fields

        private readonly string[] m_options;
        private readonly bool[] m_canHaveValue;
        private readonly bool[] m_requiresValue;

        #endregion

        #region Constructors

        internal Abbrevs(string[] options)
        {
            m_options = new string[options.Length];
            m_requiresValue = new bool[options.Length];
            m_canHaveValue = new bool[options.Length];

            for (int i = 0; i < options.Length; i++)
            {
                string option = options[i].ToLower(CultureInfo.InvariantCulture);

                if (option.StartsWith("*", StringComparison.Ordinal))
                {
                    m_requiresValue[i] = true;
                    m_canHaveValue[i] = true;
                    option = option.Substring(1);
                }
                else if (option.StartsWith("+", StringComparison.Ordinal))
                {
                    m_requiresValue[i] = false;
                    m_canHaveValue[i] = true;
                    option = option.Substring(1);
                }

                m_options[i] = option;
            }
        }

        #endregion

        #region Internal Methods

        internal string Lookup(string strOpt, out bool requiresValue, out bool canHaveValue)
        {
            int index = -1;
            bool found = false;

            for (int i = 0; i < m_options.Length; i++)
            {
                if (strOpt.Equals(m_options[i], StringComparison.OrdinalIgnoreCase))
                {
                    requiresValue = m_requiresValue[i];
                    canHaveValue = m_canHaveValue[i];
                    return m_options[i];
                }

                if (!m_options[i].StartsWith(strOpt, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (found)
                    throw new ApplicationException(string.Format("Ambiguous option: /{0}", strOpt));
                    
                found = true;
                index = i;
            }

            if (!found)
                throw new ApplicationException(string.Format("Unknown option: /{0}", strOpt));

            requiresValue = m_requiresValue[index];
            canHaveValue = m_canHaveValue[index];
            return m_options[index];
        }

        #endregion
    }
}
