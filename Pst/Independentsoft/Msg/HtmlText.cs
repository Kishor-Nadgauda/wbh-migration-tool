using System;

namespace Independentsoft.Msg
{
    internal class HtmlText
    {
        private string text;
        private System.Text.Encoding encoding = System.Text.Encoding.UTF8;

        internal HtmlText()
        {
        }

        internal HtmlText(string text, System.Text.Encoding encoding)
        {
            this.text = text;
            this.encoding = encoding;
        }

        #region Properties

        internal string Text
        {
            get
            {
                return text;
            }
        }

        internal System.Text.Encoding Encoding
        {
            get
            {
                return encoding;
            }
        }

        internal byte[] GetBytes()
        {
            if (text != null)
            {
                return encoding.GetBytes(text);
            }
            else
            {
                return null;
            }
        }

        #endregion
    }
}
