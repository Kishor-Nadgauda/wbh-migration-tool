using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Recipient.
    /// </summary>
    public class Recipient
    {
        internal string displayName;
        internal string emailAddress;
        internal string addressType;
        internal ObjectType objectType = ObjectType.None;
        internal RecipientType recipientType = RecipientType.None;
        internal DisplayType displayType = DisplayType.None;
        internal byte[] entryId;
        internal byte[] instanceKey;
        internal byte[] searchKey;
        internal bool responsibility;
        internal string smtpAddress;
        internal string displayName7Bit;
        internal string transmitableDisplayName;
        internal bool sendRichInfo;
        internal int sendInternetEncoding;
        internal string originatingAddressType;
        internal string originatingEmailAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="Recipient"/> class.
        /// </summary>
        public Recipient()
        {
        }

        #region Properties

        /// <summary>
        /// Gets the display name.
        /// </summary>
        /// <value>The display name.</value>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
        }

        /// <summary>
        /// Gets the email address.
        /// </summary>
        /// <value>The email address.</value>
        public string EmailAddress
        {
            get
            {
                return emailAddress;
            }
        }

        /// <summary>
        /// Gets the type of the address.
        /// </summary>
        /// <value>The type of the address.</value>
        public string AddressType
        {
            get
            {
                return addressType;
            }
        }

        /// <summary>
        /// Gets the type of the object.
        /// </summary>
        /// <value>The type of the object.</value>
        public ObjectType ObjectType
        {
            get
            {
                return objectType;
            }
        }

        /// <summary>
        /// Gets the type of the recipient.
        /// </summary>
        /// <value>The type of the recipient.</value>
        public RecipientType RecipientType
        {
            get
            {
                return recipientType;
            }
        }

        /// <summary>
        /// Gets the display type.
        /// </summary>
        /// <value>The display type.</value>
        public DisplayType DisplayType
        {
            get
            {
                return displayType;
            }
        }

        /// <summary>
        /// Gets the entry identifier.
        /// </summary>
        /// <value>The entry identifier.</value>
        public byte[] EntryId
        {
            get
            {
                return entryId;
            }
        }

        /// <summary>
        /// Gets the instance key.
        /// </summary>
        /// <value>The instance key.</value>
        public byte[] InstanceKey
        {
            get
            {
                return instanceKey;
            }
        }

        /// <summary>
        /// Gets the search key.
        /// </summary>
        /// <value>The search key.</value>
        public byte[] SearchKey
        {
            get
            {
                return searchKey;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this <see cref="Recipient"/> is responsibility.
        /// </summary>
        /// <value><c>true</c> if responsibility; otherwise, <c>false</c>.</value>
        public bool Responsibility
        {
            get
            {
                return responsibility;
            }
        }

        /// <summary>
        /// Gets the SMTP address.
        /// </summary>
        /// <value>The SMTP address.</value>
        public string SmtpAddress
        {
            get
            {
                return smtpAddress;
            }
        }

        /// <summary>
        /// Gets the display name7 bit.
        /// </summary>
        /// <value>The display name7 bit.</value>
        public string DisplayName7Bit
        {
            get
            {
                return displayName7Bit;
            }
        }

        /// <summary>
        /// Gets the display name of the transmitable.
        /// </summary>
        /// <value>The display name of the transmitable.</value>
        public string TransmitableDisplayName
        {
            get
            {
                return transmitableDisplayName;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [send rich information].
        /// </summary>
        /// <value><c>true</c> if [send rich information]; otherwise, <c>false</c>.</value>
        public bool SendRichInfo
        {
            get
            {
                return sendRichInfo;
            }
        }

        /// <summary>
        /// Gets the send internet encoding.
        /// </summary>
        /// <value>The send internet encoding.</value>
        public int SendInternetEncoding
        {
            get
            {
                return sendInternetEncoding;
            }
        }

        /// <summary>
        /// Gets the type of the originating address.
        /// </summary>
        /// <value>The type of the originating address.</value>
        public string OriginatingAddressType
        {
            get
            {
                return originatingAddressType;
            }
        }

        /// <summary>
        /// Gets the originating email address.
        /// </summary>
        /// <value>The originating email address.</value>
        public string OriginatingEmailAddress
        {
            get
            {
                return originatingEmailAddress;
            }
        }

        #endregion
    }
}
