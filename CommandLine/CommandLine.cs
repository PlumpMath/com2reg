using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace NKhil.Tools.Com2Reg.CommandLine
{
    internal class CommandLine
    {
        #region Fields

        private readonly string[] m_arguments;
        private readonly Option[] m_options;

        #endregion

        #region Properties

        internal ICollection<string> Arguments
        {
            get { return new ReadOnlyCollection<string>(m_arguments); }
        }

        internal ICollection<Option> Options
        {
            get { return new ReadOnlyCollection<Option>(m_options); }
        }

        #endregion

        #region Constructors

        internal CommandLine(string[] args, string[] validOpts)
        {
            Abbrevs validOptions = new Abbrevs(validOpts);
            List<string> arguments = new List<string>(args.Length);
            List<Option> options = new List<Option>(args.Length);

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].StartsWith("/", StringComparison.Ordinal) ||
                    args[i].StartsWith("-", StringComparison.Ordinal))
                {
                    string optionName;
                    bool requiresValue;
                    bool canHaveValue;
                    string optionValue = null;

                    int equalSignPos = args[i].IndexOfAny(new[] {':', '='});
                    if (equalSignPos == -1)
                        optionName = args[i].Substring(1);
                    else
                        optionName = args[i].Substring(1, equalSignPos - 1);

                    optionName = validOptions.Lookup(optionName, out requiresValue, out canHaveValue);

                    if (!canHaveValue && (equalSignPos != -1))
                        throw new ApplicationException(string.Format("The /{0} option does not require a value",
                            optionName));

                    if (requiresValue && (equalSignPos == -1))
                        throw new ApplicationException(string.Format("The /{0} option requires a value", optionName));

                    if (canHaveValue && (equalSignPos != -1))
                    {
                        if (equalSignPos == (args[i].Length - 1))
                        {
                            if ((i + 1) == args.Length)
                            {
                                throw new ApplicationException(string.Format("The /{0} option requires a value",
                                    optionName));
                            }

                            if (args[i + 1].StartsWith("/", StringComparison.Ordinal) ||
                                args[i + 1].StartsWith("-", StringComparison.Ordinal))
                            {
                                throw new ApplicationException(string.Format("The /{0} option requires a value",
                                    optionName));
                            }

                            optionValue = args[i + 1];
                            i++;
                        }
                        else
                            optionValue = args[i].Substring(equalSignPos + 1);
                    }

                    options.Add(new Option(optionName, optionValue));
                }
                else
                    arguments.Add(args[i]);
            }

            m_arguments = arguments.ToArray();
            m_options = options.ToArray();
        }

        #endregion
    }
}
