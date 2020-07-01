using System;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Represents a mailbox.
    /// </summary>
    public class Mailbox
    {
        private string emailAddress;
        private string name;

        /// <summary>
        /// Initializes a new instance of the Mailbox.
        /// </summary>
        public Mailbox()
        {
        }

        /// <summary>
        /// Initializes a new instance of the Mailbox.
        /// </summary>
        /// <param name="emailAddress">Email address of mailbox owner.</param>
        /// <param name="name">Name of mailbox owner.</param>
        public Mailbox(string emailAddress, string name)
        {
            this.emailAddress = emailAddress;
            this.name = name;
        }

        /// <summary>
        /// Initializes a new instance of the Mailbox.
        /// </summary>
        /// <param name="mailbox">Email address and name of mailbox owner.</param>
        public Mailbox(string mailbox)
        {
            Parse(mailbox);
        }

        internal void Parse(string mailbox)
        {
            if (mailbox != null)
            {
                mailbox = mailbox.Trim();
                mailbox = mailbox.Replace("\t", " ");

                int startBracketIndex = mailbox.LastIndexOf("<");
                int endBracketIndex = mailbox.LastIndexOf(">");
                int lastSpaceIndex = mailbox.LastIndexOf(" ");

                string emailAddressString = null;
                string nameString = null;

                if (startBracketIndex > -1 && endBracketIndex > startBracketIndex)
                {
                    nameString = mailbox.Substring(0, startBracketIndex);
                    emailAddressString = mailbox.Substring(startBracketIndex + 1, endBracketIndex - startBracketIndex - 1);
                }
                else if (lastSpaceIndex == -1)
                {
                    emailAddressString = mailbox;
                }
                else
                {
                    nameString = mailbox.Substring(0, lastSpaceIndex);
                    emailAddressString = mailbox.Substring(lastSpaceIndex + 1);
                }

                if (nameString != null)
                {
                    nameString = nameString.Trim(new char[] { ' ', '"' });
                }

                this.emailAddress = emailAddressString;
                this.name = nameString;
            }
        }

        /// <summary>
        /// Returns a String that represents the current Mailbox.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string nameString = null;
            string emailAddressString = null;

            if (name != null && name.Length > 0)
            {
                nameString = "\"" + name.Trim() + "\" ";
            }

            if (emailAddress != null && emailAddress.Length > 0)
            {
                emailAddressString = "<" + emailAddress.Trim() + ">";
            }

            if (nameString != null)
            {
                return nameString + " " + emailAddressString;
            }
            else
            {
                return emailAddressString;
            }
        }

        #region Properties

        /// <summary>
        /// Gets or sets display name.
        /// </summary>
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Gets or sets email address.
        /// </summary>
        public string EmailAddress
        {
            get
            {
                return emailAddress;
            }
            set
            {
                emailAddress = value;
            }
        }

        #endregion
    }
}
