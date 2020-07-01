using System;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class ContentDisposition.
    /// </summary>
    public class ContentDisposition
    {
        private ContentDispositionType type = ContentDispositionType.Inline;
        private ParameterList parameters = new ParameterList();

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentDisposition"/> class.
        /// </summary>
        public ContentDisposition()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentDisposition"/> class.
        /// </summary>
        /// <param name="type">The type.</param>
        public ContentDisposition(ContentDispositionType type)
        {
            this.type = type;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentDisposition"/> class.
        /// </summary>
        /// <param name="contentDisposition">The content disposition.</param>
        public ContentDisposition(string contentDisposition)
        {
            Parse(contentDisposition);
        }

        private void Parse(string contentDisposition)
        {
            if (contentDisposition != null)
            {
                string[] splitted = contentDisposition.Split(new char[] { ';' });
                
                for (int i = 0; i < splitted.Length - 1; i++)
                {
                    int firstQuote1 = splitted[i].IndexOf("\"");

                    if (firstQuote1 > -1)
                    {
                        int secondQuote1 = splitted[i].IndexOf("\"", firstQuote1 + 1);

                        int firstQuote2 = splitted[i + 1].IndexOf("\"");

                        if (firstQuote2 > -1)
                        {
                            int secondQuote2 = splitted[i + 1].IndexOf("\"", firstQuote2 + 1);

                            if (firstQuote1 > 0 && secondQuote1 == -1 && firstQuote2 > 0 && secondQuote2 == -1)
                            {
                                splitted[i] = splitted[i] + ";" + splitted[i + 1];
                                splitted[i + 1] = "";
                            }
                        }
                    }
                }

                if (splitted.Length > 0)
                {
                    this.type = EnumUtil.ParseContentDispositionType(splitted[0]);
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

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            string stringValue = EnumUtil.ParseContentDispositionType(type);

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
        public ContentDispositionType Type
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
