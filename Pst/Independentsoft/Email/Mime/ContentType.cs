using System;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class ContentType.
    /// </summary>
    public class ContentType
    {
        private string type;
        private string subtype;
        private ParameterList parameters = new ParameterList();

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentType"/> class.
        /// </summary>
        public ContentType()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentType"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="subtype">The subtype.</param>
        public ContentType(string type, string subtype)
        {
            this.type = type;
            this.subtype = subtype;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentType"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="subtype">The subtype.</param>
        /// <param name="charset">The charset.</param>
        public ContentType(string type, string subtype, string charset) : this(type, subtype)
        {
            CharSet = charset;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentType"/> class.
        /// </summary>
        /// <param name="contentType">Type of the content.</param>
        public ContentType(string contentType)
        {
            Parse(contentType);
        }

        private void Parse(string contentType)
        {
            if (contentType != null)
            {
                string[] splitted = Split(contentType);

                if (splitted.Length > 0)
                {
                    splitted[0] = splitted[0].Trim();

                    string[] first = splitted[0].Split(new char[] { '/' });

                    if (first.Length == 1)
                    {
                        this.type = first[0];
                    }
                    else if (first.Length == 2)
                    {
                        this.type = first[0];
                        this.subtype = first[1];
                    }
                }

                for (int i = 1; i < splitted.Length; i++)
                {
                    splitted[i] = splitted[i].Trim();

                    if (splitted[i].Length > 0)
                    {
                        int equalIndex = splitted[i].IndexOf("=");

                        if (equalIndex > -1)
                        {
                            string parameterName = splitted[i].Substring(0, equalIndex);
                            string parameterValue = splitted[i].Substring(equalIndex + 1);

                            parameterName = parameterName.Trim();
                            parameterValue = parameterValue.Trim(new char[] { '"', ' ' });

                            Parameter parameter = new Parameter(parameterName, parameterValue);
                            parameters.Add(parameter);
                        }
                    }
                }
            }
        }

        private static string[] Split(string input)
        {
            IList<String> list = new List<String>();

            int separatorIndex = input.IndexOf(";");

            while (separatorIndex > -1)
            {
                int openQuoteIndex = input.IndexOf("\"");
                int closeQuoteIndex = input.IndexOf("\"", openQuoteIndex + 1);

                if (openQuoteIndex < separatorIndex && separatorIndex < closeQuoteIndex)
                {
                    separatorIndex = input.IndexOf(";", closeQuoteIndex + 1);
                }
                else
                {
                    string line = input.Substring(0, separatorIndex);
                    input = input.Substring(separatorIndex + 1);

                    list.Add(line);

                    separatorIndex = input.IndexOf(";");
                }
            }

            list.Add(input);

            string[] splitted = new string[list.Count];

            for (int i = 0; i < splitted.Length; i++)
            {
                splitted[i] = list[i];
            }

            return splitted;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            if (type == null || subtype == null)
            {
                return "";
            }

            string stringValue = type + "/" + subtype;

            for (int i = 0; i < parameters.Count; i++)
            {
                if (parameters[i].Name != null && parameters[i].Value != null)
                {
                    stringValue += "; " + parameters[i].Name + "=" + parameters[i].Value;
                }
            }

            return stringValue;
        }

        #region Properties

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The type.</value>
        public string Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
            }
        }

        /// <summary>
        /// Gets or sets the type of the sub.
        /// </summary>
        /// <value>The type of the sub.</value>
        public string SubType
        {
            get
            {
                return subtype;
            }
            set
            {
                subtype = value;
            }
        }

        /// <summary>
        /// Gets or sets the character set.
        /// </summary>
        /// <value>The character set.</value>
        public string CharSet
        {
            get
            {
                Parameter charsetParameter = parameters["charset"];

                if (charsetParameter != null)
                {
                    return charsetParameter.Value;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value != null)
                {
                    parameters.Remove("charset");

                    Parameter charsetParameter = new Parameter("charset", value);
                    parameters.Add(charsetParameter);
                }
                else
                {
                    parameters.Remove("charset");
                }
            }
        }

        /// <summary>
        /// Gets the parameters.
        /// </summary>
        /// <value>The parameters.</value>
        public ParameterList Parameters
        {
            get
            {
                return parameters;
            }
        }

        #endregion
    }
}
