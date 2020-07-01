using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents a user or resource, generally a mail message addressee.
    /// </summary>
    public class Recipient
    {
        private string displayName;
        private string emailAddress;
        private string addressType;
        private ObjectType objectType = ObjectType.None;
        private RecipientType recipientType = RecipientType.None;
        private DisplayType displayType = DisplayType.None;
        private byte[] entryId;
        private byte[] instanceKey;
        private byte[] searchKey;
        private bool responsibility;
        private string smtpAddress;
        private string displayName7Bit;
        private string transmitableDisplayName;
        private bool sendRichInfo;
        private int sendInternetEncoding;
        private string originatingAddressType;
        private string originatingEmailAddress;


        /// <summary>
        /// Initializes a new instance of the Recipient class. 
        /// </summary>
        public Recipient()
        {
        }

        #region Properties

        /// <summary>
        /// Contains the display name of the recipient.
        /// </summary>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
            set
            {
                displayName = value;
            }
        }

        /// <summary>
        /// Contains the recipient's e-mail address.
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

        /// <summary>
        /// Contains the recipient's e-mail address type, such as Simple Mail Transfer Protocol (SMTP).
        /// </summary>
        public string AddressType
        {
            get
            {
                return addressType;
            }
            set
            {
                addressType = value;
            }
        }

        /// <summary>
        /// Contains the type of the recipient.
        /// </summary>
        public ObjectType ObjectType
        {
            get
            {
                return objectType;
            }
            set
            {
                objectType = value;
            }
        }

        /// <summary>
        /// Contains the recipient type for a message recipient.
        /// </summary>
        public RecipientType RecipientType
        {
            get
            {
                return recipientType;
            }
            set
            {
                recipientType = value;
            }
        }

        /// <summary>
        /// Contains a value used to associate an icon with a particular row of a table.
        /// </summary>
        public DisplayType DisplayType
        {
            get
            {
                return displayType;
            }
            set
            {
                displayType = value;
            }
        }

        /// <summary>
        /// Contains the EntryID of the recipient.
        /// </summary>
        public byte[] EntryId
        {
            get
            {
                return entryId;
            }
            set
            {
                entryId = value;
            }
        }

        /// <summary>
        /// Contains a value that uniquely identifies a row in a table.
        /// </summary>
        public byte[] InstanceKey
        {
            get
            {
                return instanceKey;
            }
            set
            {
                instanceKey = value;
            }
        }

        /// <summary>
        /// Contains a binary-comparable key that identifies correlated objects for a search.
        /// </summary>
        public byte[] SearchKey
        {
            get
            {
                return searchKey;
            }
            set
            {
                searchKey = value;
            }
        }

        /// <summary>
        /// Contains true if some transport provider has already accepted responsibility for delivering the message to this recipient, and false if the MAPI spooler considers that this transport provider should accept responsibility.
        /// </summary>
        public bool Responsibility
        {
            get
            {
                return responsibility;
            }
            set
            {
                responsibility = value;
            }
        }

        /// <summary>
        /// Contains SMTP email address.
        /// </summary>
        public string SmtpAddress
        {
            get
            {
                return smtpAddress;
            }
            set
            {
                smtpAddress = value;
            }
        }

        /// <summary>
        /// Contains a 7-bit ASCII representation of the recipient's display name.
        /// </summary>
        public string DisplayName7Bit
        {
            get
            {
                return displayName7Bit;
            }
            set
            {
                displayName7Bit = value;
            }
        }

        /// <summary>
        /// Contains a recipient's display name in a secure form that cannot be changed.
        /// </summary>
        public string TransmitableDisplayName
        {
            get
            {
                return transmitableDisplayName;
            }
            set
            {
                transmitableDisplayName = value;
            }
        }

        /// <summary>
        /// Contains true if the recipient can receive all message content, including Rich Text Format (RTF) and Object Linking and Embedding (OLE) objects.
        /// </summary>
        public bool SendRichInfo
        {
            get
            {
                return sendRichInfo;
            }
            set
            {
                sendRichInfo = value;
            }
        }

        /// <summary>
        /// Contains a bitmask of encoding preferences.
        /// </summary>
        public int SendInternetEncoding
        {
            get
            {
                return sendInternetEncoding;
            }
            set
            {
                sendInternetEncoding = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string OriginatingAddressType
        {
            get
            {
                return originatingAddressType;
            }
            set
            {
                originatingAddressType = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string OriginatingEmailAddress
        {
            get
            {
                return originatingEmailAddress;
            }
            set
            {
                originatingEmailAddress = value;
            }
        }

        #endregion
    }
}