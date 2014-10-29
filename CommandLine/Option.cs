namespace NKhil.Tools.Com2Reg.CommandLine
{
    internal class Option
    {
        #region Fields

        private readonly string m_name;
        private readonly string m_value;

        #endregion

        #region Properties

        internal string Name
        {
            get
            {
                return m_name;
            }
        }

        internal string Value
        {
            get
            {
                return m_value;
            }
        }

        #endregion

        #region Constructors

        internal Option(string name, string value)
        {
            m_name = name;
            m_value = value;
        }

        #endregion
    }
}
