using System;
using System.IO;
using System.Collections.Generic;
using Independentsoft.Email.Mime;
using Independentsoft.IO.StructuredStorage;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents Outlook message file.
    /// </summary>
    public class Message
    {
        private IDictionary<String, Property> propertyTable;

        private string messageClass = "IPM.Note";
        private string subject;
        private string subjectPrefix;
        private string conversationTopic;
        private string displayBcc;
        private string displayCc;
        private string displayTo;
        private string originalDisplayTo;
        private string replyTo;
        private string normalizedSubject;
        private string body;
        private byte[] rtfCompressed;
        private byte[] searchKey;
        private byte[] changeKey;
        private byte[] entryId;
        private byte[] readReceiptEntryId;
        private byte[] readReceiptSearchKey;
        private DateTime creationTime;
        private DateTime lastModificationTime;
        private DateTime messageDeliveryTime;
        private DateTime clientSubmitTime;
        private DateTime deferredDeliveryTime;
        private DateTime providerSubmitTime;
        private DateTime reportTime;
        private string reportText;
        private string creatorName;
        private string lastModifierName;
        private string internetMessageId;
        private string inReplyTo;
        private string internetReferences;
        private uint messageCodePage;
        private uint iconIndex;
        private uint messageSize;
        private uint internetCodePage;
        private byte[] conversationIndex;
        private bool isHidden;
        private bool isReadOnly;
        private bool isSystem;
        private bool disableFullFidelity;
        private bool hasAttachment;
        private bool rtfInSync;
        private bool readReceiptRequested;
        private bool deliveryReportRequested;
        private byte[] bodyHtml;
        private Sensitivity sensitivity = Sensitivity.None;
        private Importance importance = Importance.None;
        private Priority priority = Priority.None;
        private FlagIcon flagIcon = FlagIcon.None;
        private FlagStatus flagStatus = FlagStatus.None;
        private ObjectType objectType = ObjectType.None;
        private string receivedRepresentingAddressType;
        private string receivedRepresentingEmailAddress;
        private byte[] receivedRepresentingEntryId;
        private string receivedRepresentingName;
        private byte[] receivedRepresentingSearchKey;
        private string receivedByAddressType;
        private string receivedByEmailAddress;
        private byte[] receivedByEntryId;
        private string receivedByName;
        private byte[] receivedBySearchKey;
        private string senderAddressType;
        private string senderEmailAddress;
        private string extendedSenderEmailAddress;
        private string senderSmtpAddress;
        private byte[] senderEntryId;
        private string senderName;
        private byte[] senderSearchKey;
        private string sentRepresentingAddressType;
        private string sentRepresentingEmailAddress;
        private string sentRepresentingSmtpAddress;
        private byte[] sentRepresentingEntryId;
        private string sentRepresentingName;
        private byte[] sentRepresentingSearchKey;
        private string transportMessageHeaders;
        private DateTime lastVerbExecutionTime;
        private LastVerbExecuted lastVerbExecuted = LastVerbExecuted.None;
        private IList<MessageFlag> messageFlags = new List<MessageFlag>();
        private IList<StoreSupportMask> storeSupportMasks = new List<StoreSupportMask>();

        //general named properties
        private string outlookVersion;
        private uint outlookInternalVersion;
        private DateTime commonStartTime;
        private DateTime commonEndTime;
        private DateTime flagDueBy;
        private bool isRecurring;
        private DateTime reminderTime;
        private uint reminderMinutesBeforeStart;
        private IList<string> companies = new List<string>();
        private IList<string> contactNames = new List<string>();
        private IList<string> keywords = new List<string>();
        private string billingInformation;
        private string mileage;
        private string reminderSoundFile;
        private bool isPrivate;
        private bool isReminderSet;
        private bool reminderOverrideDefault;
        private bool reminderPlaySound;
        private string internetAccountName;

        //appointments named properties
        private DateTime appointmentStartTime;
        private DateTime appointmentEndTime;
        private bool isAllDayEvent;
        private string location;
        private BusyStatus busyStatus = BusyStatus.None;
        private MeetingStatus meetingStatus = MeetingStatus.None;
        private ResponseStatus responseStatus = ResponseStatus.None;
        private RecurrenceType recurrenceType = RecurrenceType.None;
        private string appointmentMessageClass;
        private string timeZone;
        private string recurrencePatternDescription;
        internal RecurrencePattern recurrencePattern;
        private byte[] guid;
        private int label = -1;
        private uint duration;

        //tasks named properties
        private DateTime taskStartDate;
        private DateTime taskDueDate;
        private string owner;
        private string delegator;
        private double percentComplete;
        private uint actualWork;
        private uint totalWork;
        private bool isTeamTask;
        private bool isComplete;
        private DateTime dateCompleted;
        private TaskStatus taskStatus = TaskStatus.None;
        private TaskOwnership taskOwnership = TaskOwnership.None;
        private TaskDelegationState taskDelegationState = TaskDelegationState.None;

        //notes named properties
        private uint noteWidth;
        private uint noteHeight;
        private uint noteLeft;
        private uint noteTop;
        private NoteColor noteColor = NoteColor.None;

        //journals named properties
        private DateTime journalStartTime;
        private DateTime journalEndTime;
        private string journalType;
        private string journalTypeDescription;
        private uint journalDuration;

        //contacts properties
        private DateTime birthday;
        private IList<string> childrenNames = new List<string>();
        private string assistentName;
        private string assistentPhone;
        private string businessPhone;
        private string businessPhone2;
        private string businessFax;
        private string businessHomePage;
        private string callbackPhone;
        private string carPhone;
        private string cellularPhone;
        private string companyMainPhone;
        private string companyName;
        private string computerNetworkName;
        private string customerId;
        private string departmentName;
        private string displayName;
        private string displayNamePrefix;
        private string ftpSite;
        private string generation;
        private string givenName;
        private string governmentId;
        private string hobbies;
        private string homePhone2;
        private string homeAddressCity;
        private string homeAddressCountry;
        private string homeAddressPostalCode;
        private string homeAddressPostOfficeBox;
        private string homeAddressState;
        private string homeAddressStreet;
        private string homeFax;
        private string homePhone;
        private string initials;
        private string isdn;
        private string managerName;
        private string middleName;
        private string nickname;
        private string officeLocation;
        private string otherAddressCity;
        private string otherAddressCountry;
        private string otherAddressPostalCode;
        private string otherAddressState;
        private string otherAddressStreet;
        private string otherPhone;
        private string pager;
        private string personalHomePage;
        private string postalAddress;
        private string businessAddressCountry;
        private string businessAddressCity;
        private string businessAddressPostalCode;
        private string businessAddressPostOfficeBox;
        private string businessAddressState;
        private string businessAddressStreet;
        private string primaryFax;
        private string primaryPhone;
        private string profession;
        private string radioPhone;
        private string spouseName;
        private string surname;
        private string telex;
        private string title;
        private string ttyTddPhone;
        private DateTime weddingAnniversary;
        private Gender gender = Gender.None;

        //contacts named properties
        private SelectedMailingAddress selectedMailingAddress = SelectedMailingAddress.None;
        private bool contactHasPicture;
        private string fileAs;
        private string instantMessengerAddress;
        private string internetFreeBusyAddress;
        private string businessAddress;
        private string homeAddress;
        private string otherAddress;
        private string email1Address;
        private string email2Address;
        private string email3Address;
        private string email1DisplayName;
        private string email2DisplayName;
        private string email3DisplayName;
        private string email1DisplayAs;
        private string email2DisplayAs;
        private string email3DisplayAs;
        private string email1Type;
        private string email2Type;
        private string email3Type;
        private byte[] email1EntryId;
        private byte[] email2EntryId;
        private byte[] email3EntryId;

        private IList<Recipient> recipients = new List<Recipient>();
        private IList<Attachment> attachments = new List<Attachment>();
        private ExtendedPropertyList extendedProperties = new ExtendedPropertyList();
        private IList<NamedProperty> namedProperties = new List<NamedProperty>();

        private string stringTypeMask = "001E";
        private uint stringTypeHexMask = 0x001E;
        private string multiValueStringTypeMask = "101E";
        private uint multiValueStringTypeHexMask = 0x101E;

        private string unicodeTypeMask = "001F";
        private uint unicodeTypeHexMask = 0x001F;
        private string multiValueUnicodeTypeMask = "101F";
        private uint multiValueUnicodeTypeHexMask = 0x101F;


        private System.Text.Encoding encoding = System.Text.Encoding.UTF8;
        private System.Text.Encoding encodingUnicode = System.Text.Encoding.Unicode;
        private bool isEmbedded;

        /// <summary>
        /// Initializes a new instance of the Message class.
        /// </summary>
        public Message()
        {
            storeSupportMasks.Add(StoreSupportMask.Attachments);
            storeSupportMasks.Add(StoreSupportMask.Categorize);
            storeSupportMasks.Add(StoreSupportMask.Create);
            storeSupportMasks.Add(StoreSupportMask.EntryIdUnique);
            storeSupportMasks.Add(StoreSupportMask.Html);
            storeSupportMasks.Add(StoreSupportMask.ItemProc);
            storeSupportMasks.Add(StoreSupportMask.Modify);
            storeSupportMasks.Add(StoreSupportMask.MultiValueProperties);
            storeSupportMasks.Add(StoreSupportMask.Notify);
            storeSupportMasks.Add(StoreSupportMask.Ole);
            storeSupportMasks.Add(StoreSupportMask.Pusher);
            storeSupportMasks.Add(StoreSupportMask.ReadOnly);
            storeSupportMasks.Add(StoreSupportMask.Restrictions);
            storeSupportMasks.Add(StoreSupportMask.Rtf);
            storeSupportMasks.Add(StoreSupportMask.Search);
            storeSupportMasks.Add(StoreSupportMask.Sort);
            storeSupportMasks.Add(StoreSupportMask.Submit);
            storeSupportMasks.Add(StoreSupportMask.UncompressedRtf);
            storeSupportMasks.Add(StoreSupportMask.Unicode);
        }

        /// <summary>
        /// Initializes a new instance of the Message class based on the supplied file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public Message(string filePath)
            : this()
        {
            Open(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the Message class based on the supplied stream.
        /// </summary>
        /// <param name="stream">A stream.</param>
        public Message(System.IO.Stream stream)
            : this()
        {
            Open(stream);
        }

        /// <summary>
        /// Initializes a new instance of the Message class from the specified MIME message.
        /// </summary>
        /// <param name="mimeMessage">The MIME message.</param>
        public Message(Independentsoft.Email.Mime.Message mimeMessage)
            : this()
        {
            ConvertFromMimeMessage(mimeMessage);
        }

        /// <summary>
        /// Loads message from the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Open(string filePath)
        {
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);

            using (fileStream)
            {
                Open(fileStream);
            }
        }

        /// <summary>
        /// Loads message from the specified stream.
        /// </summary>
        /// <param name="stream">An input stream.</param>
        public void Open(System.IO.Stream stream)
        {
            Parse(stream);
        }

        private void Parse(System.IO.Stream stream)
        {
            CompoundFile compoundFile = new CompoundFile(stream);

            Independentsoft.IO.StructuredStorage.Storage nameIdStorage = (Independentsoft.IO.StructuredStorage.Storage)compoundFile.Root.DirectoryEntries["__nameid_version1.0"];

            Independentsoft.IO.StructuredStorage.Stream guidStream = (Independentsoft.IO.StructuredStorage.Stream)nameIdStorage.DirectoryEntries["__substg1.0_00020102"];
            Independentsoft.IO.StructuredStorage.Stream entryStream = (Independentsoft.IO.StructuredStorage.Stream)nameIdStorage.DirectoryEntries["__substg1.0_00030102"];
            Independentsoft.IO.StructuredStorage.Stream stringStream = (Independentsoft.IO.StructuredStorage.Stream)nameIdStorage.DirectoryEntries["__substg1.0_00040102"];

            IDictionary<String, String> namedPropertyIdTable = new Dictionary<String, String>();
            namedProperties = new List<NamedProperty>();

            if (entryStream != null)
            {
                uint entryCount = (uint)entryStream.Size / 8;
                byte[] entryBuffer = entryStream.Buffer;

                for (int i = 0; i < entryCount; i++)
                {
                    uint nameOrStringOffset = BitConverter.ToUInt32(entryBuffer, i * 8);
                    uint indexAndKind = BitConverter.ToUInt32(entryBuffer, i * 8 + 4);

                    ushort propertyIndex = (ushort)(indexAndKind >> 16);
                    ushort guidAndKind = (ushort)((indexAndKind << 16) >> 16);
                    ushort guidIndex = (ushort)(guidAndKind >> 1);
                    ushort propertyKind = (ushort)(guidAndKind << 15);

                    NamedProperty namedProperty = new NamedProperty();

                    if (propertyKind == 0) //numerical named property
                    {
                        uint propertyId = nameOrStringOffset;

                        namedProperty.Id = propertyId;
                        namedProperty.Type = NamedPropertyType.Numerical;

                    }
                    else //string named property
                    {
                        int stringOffset = (int)nameOrStringOffset;

                        byte[] stringBuffer = stringStream.Buffer;
                        uint nameLength = BitConverter.ToUInt32(stringBuffer, stringOffset);

                        string name = System.Text.Encoding.Unicode.GetString(stringBuffer, stringOffset + 4, (int)nameLength);

                        namedProperty.Name = name;
                        namedProperty.Type = NamedPropertyType.String;
                    }

                    if (guidIndex == 1) //PS_MAPI
                    {
                        namedProperty.Guid = StandardPropertySet.Mapi;
                    }
                    else if (guidIndex == 2) //PS_PUBLIC_STRINGS
                    {
                        namedProperty.Guid = StandardPropertySet.PublicStrings;
                    }
                    else if (guidIndex > 2)
                    {
                        int guidOffset = guidIndex - 3;

                        byte[] guidBuffer = guidStream.Buffer;
                        byte[] guid = new byte[16];

                        System.Array.Copy(guidBuffer, guidOffset * 16, guid, 0, 16);

                        namedProperty.Guid = guid;
                    }

                    if (namedProperty.Id > 0)
                    {
                        string namedPropertyHex = Util.ConvertNamedPropertyToHex(namedProperty.Id, namedProperty.Guid);
                        string newPropertyId = String.Format("{0:X4}", 0x8000 + propertyIndex);

                        if (!namedPropertyIdTable.ContainsKey(namedPropertyHex))
                        {
                            namedPropertyIdTable.Add(namedPropertyHex, newPropertyId);
                        }
                    }
                    else if (namedProperty.Name != null)//string named properties
                    {
                        string namedPropertyHex = Util.ConvertNamedPropertyToHex(namedProperty.Name, namedProperty.Guid);
                        string newPropertyId = String.Format("{0:X4}", 0x8000 + propertyIndex);

                        if (!namedPropertyIdTable.ContainsKey(namedPropertyHex))
                        {
                            namedPropertyIdTable.Add(namedPropertyHex, newPropertyId);
                        }
                    }

                    namedProperties.Add(namedProperty);
                }
            }

            //Create extended properties from named properties
            for (int n = 0; n < namedProperties.Count; n++)
            {
                NamedProperty currentNamedProperty = namedProperties[n];

                if (currentNamedProperty.Name != null)
                {
                    ExtendedPropertyName extendedPropertyName = new ExtendedPropertyName(currentNamedProperty.Name, currentNamedProperty.Guid);
                    ExtendedProperty extendedProperty = new ExtendedProperty(extendedPropertyName);
                    extendedProperties.Add(extendedProperty);
                }
                else if (currentNamedProperty.Id > 0)
                {
                    ExtendedPropertyId extendedPropertyId = new ExtendedPropertyId((int)currentNamedProperty.Id, currentNamedProperty.Guid);
                    ExtendedProperty extendedProperty = new ExtendedProperty(extendedPropertyId);
                    extendedProperties.Add(extendedProperty);
                }
            }

            ParseMessageProperties(compoundFile.Root.DirectoryEntries, namedPropertyIdTable);
        }

        private Message(DirectoryEntryList directoryEntries, bool isEmbedded)
        {
            this.isEmbedded = isEmbedded;
            ParseMessageProperties(directoryEntries);
        }

        private void ParseMessageProperties(DirectoryEntryList directoryEntries)
        {
            ParseMessageProperties(directoryEntries, null);
        }

        private void ParseMessageProperties(DirectoryEntryList directoryEntries, IDictionary<String, String> namedPropertyIdTable)
        {
            int headerSize = 24;

            propertyTable = new Dictionary<String, Property>();

            Independentsoft.IO.StructuredStorage.Stream propertiesStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__properties_version1.0"];

            //first 24 bytes (embedded messages) or 32 bytes is header

            uint recipientCount = 0;
            uint attachmentCount = 0;

            if (propertiesStream != null)
            {
                uint reserved1 = BitConverter.ToUInt32(propertiesStream.Buffer, 0);
                uint reserved2 = BitConverter.ToUInt32(propertiesStream.Buffer, 4);
                uint nextRecipientId = BitConverter.ToUInt32(propertiesStream.Buffer, 8);
                uint nextAttachmentId = BitConverter.ToUInt32(propertiesStream.Buffer, 12);
                recipientCount = BitConverter.ToUInt32(propertiesStream.Buffer, 16);
                attachmentCount = BitConverter.ToUInt32(propertiesStream.Buffer, 20);
            }

            if (!isEmbedded)
            {
                if (propertiesStream != null)
                {
                    uint reserved3 = BitConverter.ToUInt32(propertiesStream.Buffer, 24);
                    uint reserved4 = BitConverter.ToUInt32(propertiesStream.Buffer, 28);
                }

                headerSize = 32;
            }

            if (propertiesStream != null && propertiesStream.Buffer != null)
            {
                //first 24 bytes (embedded messages) or 32 bytes is header
                for (int i = headerSize; i <= propertiesStream.Buffer.Length - 16; i += 16)
                {
                    byte[] propertyBuffer = new byte[16];

                    System.Array.Copy(propertiesStream.Buffer, i, propertyBuffer, 0, 16);

                    Property property = new Property(propertyBuffer);

                    if (property.Size > 0)
                    {
                        string propertyStreamName = "__substg1.0_" + String.Format("{0:X8}", property.Tag);

                        Independentsoft.IO.StructuredStorage.Stream currentPropertyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[propertyStreamName];

                        if (currentPropertyStream != null && currentPropertyStream.Buffer != null && currentPropertyStream.Buffer.Length > 0)
                        {
                            property.Value = new byte[currentPropertyStream.Buffer.Length];
                            System.Array.Copy(currentPropertyStream.Buffer, 0, property.Value, 0, property.Value.Length);
                        }
                    }

                    string propertyTagHex = String.Format("{0:X8}", property.Tag);

                    try
                    {
                        propertyTable.Add(propertyTagHex, property);
                    }
                    catch (ArgumentException) //ignore duplicate keys
                    {
                    }
                }
            }

            Property internetCodePageProperty = propertyTable.ContainsKey("3FDE0003") ? propertyTable["3FDE0003"] : null;

            if (internetCodePageProperty != null && internetCodePageProperty.Value != null)
            {
                internetCodePage = BitConverter.ToUInt32(internetCodePageProperty.Value, 0);
            }

            Property messageCodePageProperty = propertyTable.ContainsKey("3FFD0003") ? propertyTable["3FFD0003"] : null;

            if (messageCodePageProperty != null && messageCodePageProperty.Value != null)
            {
                messageCodePage = BitConverter.ToUInt32(messageCodePageProperty.Value, 0);
            }

            Property storeSupportMaskProperty = propertyTable.ContainsKey("340D0003") ? propertyTable["340D0003"] : null;

            if (storeSupportMaskProperty != null && storeSupportMaskProperty.Value != null)
            {
                uint storeSupportMaskValue = BitConverter.ToUInt32(storeSupportMaskProperty.Value, 0);

                storeSupportMasks = EnumUtil.ParseStoreSupportMask(storeSupportMaskValue);

                if ((storeSupportMaskValue & 0x00040000) == 0x00040000) //is Unicode
                {
                    encoding = System.Text.Encoding.Unicode;
                    stringTypeMask = "001F";
                    stringTypeHexMask = 0x001F;
                    multiValueStringTypeMask = "101F";
                    multiValueStringTypeHexMask = 0x101F;
                }
            }
            else if (internetCodePage > 0) // Set encoding
            {
                try
                {
                    //if (internetCodePage == 65001)
                    //{
                    //    internetCodePage = 65000; //UTF-7 encoding
                    //}

                    encoding = System.Text.Encoding.GetEncoding((int)internetCodePage);
                }
                catch
                {
                }
            }
            else if (messageCodePage > 0) // Set encoding
            {
                try
                {
                    encoding = System.Text.Encoding.GetEncoding((int)messageCodePage);
                }
                catch
                {
                }
            }

            Property creationTimeProperty = propertyTable.ContainsKey("30070040") ? propertyTable["30070040"] : null;

            if (creationTimeProperty != null && creationTimeProperty.Value != null)
            {
                uint creationTimeLow = BitConverter.ToUInt32(creationTimeProperty.Value, 0);
                ulong creationTimeHigh = BitConverter.ToUInt32(creationTimeProperty.Value, 4);

                if (creationTimeHigh > 0)
                {
                    long ticks = creationTimeLow + (long)(creationTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        creationTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property lastModificationTimeProperty = propertyTable.ContainsKey("30080040") ? propertyTable["30080040"] : null;

            if (lastModificationTimeProperty != null && lastModificationTimeProperty.Value != null)
            {
                uint lastModificationTimeLow = BitConverter.ToUInt32(lastModificationTimeProperty.Value, 0);
                ulong lastModificationTimeHigh = BitConverter.ToUInt32(lastModificationTimeProperty.Value, 4);

                if (lastModificationTimeHigh > 0)
                {
                    long ticks = lastModificationTimeLow + (long)(lastModificationTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        lastModificationTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property messageDeliveryTimeProperty = propertyTable.ContainsKey("0E060040") ? propertyTable["0E060040"] : null;

            if (messageDeliveryTimeProperty != null && messageDeliveryTimeProperty.Value != null)
            {
                uint messageDeliveryTimeLow = BitConverter.ToUInt32(messageDeliveryTimeProperty.Value, 0);
                ulong messageDeliveryTimeHigh = BitConverter.ToUInt32(messageDeliveryTimeProperty.Value, 4);

                if (messageDeliveryTimeHigh > 0)
                {
                    long ticks = messageDeliveryTimeLow + (long)(messageDeliveryTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        messageDeliveryTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property clientSubmitTimeProperty = propertyTable.ContainsKey("00390040") ? propertyTable["00390040"] : null;

            if (clientSubmitTimeProperty != null && clientSubmitTimeProperty.Value != null)
            {
                uint clientSubmitTimeLow = BitConverter.ToUInt32(clientSubmitTimeProperty.Value, 0);
                ulong clientSubmitTimeHigh = BitConverter.ToUInt32(clientSubmitTimeProperty.Value, 4);

                if (clientSubmitTimeHigh > 0)
                {
                    long ticks = clientSubmitTimeLow + (long)(clientSubmitTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        clientSubmitTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property deferredDeliveryTimeProperty = propertyTable.ContainsKey("000F0040") ? propertyTable["000F0040"] : null;

            if (deferredDeliveryTimeProperty != null && deferredDeliveryTimeProperty.Value != null)
            {
                uint deferredDeliveryTimeLow = BitConverter.ToUInt32(deferredDeliveryTimeProperty.Value, 0);
                ulong deferredDeliveryTimeHigh = BitConverter.ToUInt32(deferredDeliveryTimeProperty.Value, 4);

                if (deferredDeliveryTimeHigh > 0)
                {
                    long ticks = deferredDeliveryTimeLow + (long)(deferredDeliveryTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        deferredDeliveryTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property providerSubmitTimeProperty = propertyTable.ContainsKey("00480040") ? propertyTable["00480040"] : null;

            if (providerSubmitTimeProperty != null && providerSubmitTimeProperty.Value != null)
            {
                uint providerSubmitTimeLow = BitConverter.ToUInt32(providerSubmitTimeProperty.Value, 0);
                ulong providerSubmitTimeHigh = BitConverter.ToUInt32(providerSubmitTimeProperty.Value, 4);

                if (providerSubmitTimeHigh > 0)
                {
                    long ticks = providerSubmitTimeLow + (long)(providerSubmitTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        providerSubmitTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property reportTimeProperty = propertyTable.ContainsKey("00320040") ? propertyTable["00320040"] : null;

            if (reportTimeProperty != null && reportTimeProperty.Value != null)
            {
                uint reportTimeLow = BitConverter.ToUInt32(reportTimeProperty.Value, 0);
                ulong reportTimeHigh = BitConverter.ToUInt32(reportTimeProperty.Value, 4);

                if (reportTimeHigh > 0)
                {
                    long ticks = reportTimeLow + (long)(reportTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        reportTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property lastVerbExecutionTimeProperty = propertyTable.ContainsKey("10820040") ? propertyTable["10820040"] : null;

            if (lastVerbExecutionTimeProperty != null && lastVerbExecutionTimeProperty.Value != null)
            {
                uint lastVerbExecutionTimeLow = BitConverter.ToUInt32(lastVerbExecutionTimeProperty.Value, 0);
                ulong lastVerbExecutionTimeHigh = BitConverter.ToUInt32(lastVerbExecutionTimeProperty.Value, 4);

                if (lastVerbExecutionTimeHigh > 0)
                {
                    long ticks = lastVerbExecutionTimeLow + (long)(lastVerbExecutionTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        lastVerbExecutionTime = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property iconIndexProperty = propertyTable.ContainsKey("10800003") ? propertyTable["10800003"] : null;

            if (iconIndexProperty != null && iconIndexProperty.Value != null)
            {
                iconIndex = BitConverter.ToUInt32(iconIndexProperty.Value, 0);
            }

            Property messageSizeProperty = propertyTable.ContainsKey("0E080003") ? propertyTable["0E080003"] : null;

            if (messageSizeProperty != null && messageSizeProperty.Value != null)
            {
                messageSize = BitConverter.ToUInt32(messageSizeProperty.Value, 0);
            }

            Property messageFlagsProperty = propertyTable.ContainsKey("0E070003") ? propertyTable["0E070003"] : null;

            if (messageFlagsProperty != null && messageFlagsProperty.Value != null)
            {
                uint messageFlagsValue = BitConverter.ToUInt32(messageFlagsProperty.Value, 0);

                messageFlags = EnumUtil.ParseMessageFlag(messageFlagsValue);
            }


            Property isHiddenProperty = propertyTable.ContainsKey("10F4000B") ? propertyTable["10F4000B"] : null;

            if (isHiddenProperty != null && isHiddenProperty.Value != null)
            {
                ushort isHiddenValue = BitConverter.ToUInt16(isHiddenProperty.Value, 0);

                if (isHiddenValue > 0)
                {
                    isHidden = true;
                }
            }

            Property isReadOnlyProperty = propertyTable.ContainsKey("10F6000B") ? propertyTable["10F6000B"] : null;

            if (isReadOnlyProperty != null && isReadOnlyProperty.Value != null)
            {
                ushort isReadOnlyValue = BitConverter.ToUInt16(isReadOnlyProperty.Value, 0);

                if (isReadOnlyValue > 0)
                {
                    isReadOnly = true;
                }
            }

            Property isSystemProperty = propertyTable.ContainsKey("10F5000B") ? propertyTable["10F5000B"] : null;

            if (isSystemProperty != null && isSystemProperty.Value != null)
            {
                ushort isSystemValue = BitConverter.ToUInt16(isSystemProperty.Value, 0);

                if (isSystemValue > 0)
                {
                    isSystem = true;
                }
            }

            Property disableFullFidelityProperty = propertyTable.ContainsKey("10F2000B") ? propertyTable["10F2000B"] : null;

            if (disableFullFidelityProperty != null && disableFullFidelityProperty.Value != null)
            {
                ushort disableFullFidelityValue = BitConverter.ToUInt16(disableFullFidelityProperty.Value, 0);

                if (disableFullFidelityValue > 0)
                {
                    disableFullFidelity = true;
                }
            }

            Property hasAttachmentProperty = propertyTable.ContainsKey("0E1B000B") ? propertyTable["0E1B000B"] : null;

            if (hasAttachmentProperty != null && hasAttachmentProperty.Value != null)
            {
                ushort hasAttachmentValue = BitConverter.ToUInt16(hasAttachmentProperty.Value, 0);

                if (hasAttachmentValue > 0)
                {
                    hasAttachment = true;
                }
            }

            Property rtfInSyncProperty = propertyTable.ContainsKey("0E1F000B") ? propertyTable["0E1F000B"] : null;

            if (rtfInSyncProperty != null && rtfInSyncProperty.Value != null)
            {
                ushort rtfInSyncValue = BitConverter.ToUInt16(rtfInSyncProperty.Value, 0);

                if (rtfInSyncValue > 0)
                {
                    rtfInSync = true;
                }
            }

            Property readReceiptRequestedProperty = propertyTable.ContainsKey("0029000B") ? propertyTable["0029000B"] : null;

            if (readReceiptRequestedProperty != null && readReceiptRequestedProperty.Value != null)
            {
                ushort readReceiptRequestedValue = BitConverter.ToUInt16(readReceiptRequestedProperty.Value, 0);

                if (readReceiptRequestedValue > 0)
                {
                    readReceiptRequested = true;
                }
            }

            Property deliveryReportRequestedProperty = propertyTable.ContainsKey("0023000B") ? propertyTable["0023000B"] : null;

            if (deliveryReportRequestedProperty != null && deliveryReportRequestedProperty.Value != null)
            {
                ushort deliveryReportRequestedValue = BitConverter.ToUInt16(deliveryReportRequestedProperty.Value, 0);

                if (deliveryReportRequestedValue > 0)
                {
                    deliveryReportRequested = true;
                }
            }

            Property sensitivityProperty = propertyTable.ContainsKey("00360003") ? propertyTable["00360003"] : null;

            if (sensitivityProperty != null && sensitivityProperty.Value != null)
            {
                uint sensitivityValue = BitConverter.ToUInt32(sensitivityProperty.Value, 0);

                sensitivity = EnumUtil.ParseSensitivity(sensitivityValue);
            }

            Property lastVerbExecutedProperty = propertyTable.ContainsKey("10810003") ? propertyTable["10810003"] : null;

            if (lastVerbExecutedProperty != null && lastVerbExecutedProperty.Value != null)
            {
                uint lastVerbExecutedPropertyValue = BitConverter.ToUInt32(lastVerbExecutedProperty.Value, 0);

                lastVerbExecuted = EnumUtil.ParseLastVerbExecuted(lastVerbExecutedPropertyValue);
            }

            Property importanceProperty = propertyTable.ContainsKey("00170003") ? propertyTable["00170003"] : null;

            if (importanceProperty != null && importanceProperty.Value != null)
            {
                uint importanceValue = BitConverter.ToUInt32(importanceProperty.Value, 0);

                importance = EnumUtil.ParseImportance(importanceValue);
            }

            Property priorityProperty = propertyTable.ContainsKey("00260003") ? propertyTable["00260003"] : null;

            if (priorityProperty != null && priorityProperty.Value != null)
            {
                uint priorityValue = BitConverter.ToUInt32(priorityProperty.Value, 0);

                priority = EnumUtil.ParsePriority(priorityValue);
            }

            Property flagIconProperty = propertyTable.ContainsKey("10950003") ? propertyTable["10950003"] : null;

            if (flagIconProperty != null && flagIconProperty.Value != null)
            {
                uint flagIconValue = BitConverter.ToUInt32(flagIconProperty.Value, 0);

                flagIcon = EnumUtil.ParseFlagIcon(flagIconValue);
            }

            Property flagStatusProperty = propertyTable.ContainsKey("10900003") ? propertyTable["10900003"] : null;

            if (flagStatusProperty != null && flagStatusProperty.Value != null)
            {
                uint flagStatusValue = BitConverter.ToUInt32(flagStatusProperty.Value, 0);

                flagStatus = EnumUtil.ParseFlagStatus(flagStatusValue);
            }

            Property objectTypeProperty = propertyTable.ContainsKey("0FFE0003") ? propertyTable["0FFE0003"] : null;

            if (objectTypeProperty != null && objectTypeProperty.Value != null)
            {
                uint objectTypeValue = BitConverter.ToUInt32(objectTypeProperty.Value, 0);

                objectType = EnumUtil.ParseObjectType(objectTypeValue);
            }

            uint outlookVersionNamedPropertyId = 0x8554;
            byte[] outlookVersionNamedPropertyGuid = StandardPropertySet.Common;

            string outlookVersionHex = Util.ConvertNamedPropertyToHex(outlookVersionNamedPropertyId, outlookVersionNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(outlookVersionHex) && namedPropertyIdTable[outlookVersionHex] != null)
            {
                string outlookVersionPropertyId = namedPropertyIdTable[outlookVersionHex];

                outlookVersionPropertyId = outlookVersionPropertyId + stringTypeMask;

                Property outlookVersionProperty = propertyTable.ContainsKey(outlookVersionPropertyId) ? propertyTable[outlookVersionPropertyId] : null;

                if (outlookVersionProperty != null && outlookVersionProperty.Value != null)
                {
                    outlookVersion = encoding.GetString(outlookVersionProperty.Value, 0, outlookVersionProperty.Value.Length);
                }
            }

            uint outlookInternalVersionNamedPropertyId = 0x8552;
            byte[] outlookInternalVersionNamedPropertyGuid = StandardPropertySet.Common;

            string outlookInternalVersionHex = Util.ConvertNamedPropertyToHex(outlookInternalVersionNamedPropertyId, outlookInternalVersionNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(outlookInternalVersionHex) && namedPropertyIdTable[outlookInternalVersionHex] != null)
            {
                string outlookInternalVersionPropertyId = namedPropertyIdTable[outlookInternalVersionHex];

                outlookInternalVersionPropertyId = outlookInternalVersionPropertyId + "0003";

                Property outlookInternalVersionProperty = propertyTable.ContainsKey(outlookInternalVersionPropertyId) ? propertyTable[outlookInternalVersionPropertyId] : null;

                if (outlookInternalVersionProperty != null && outlookInternalVersionProperty.Value != null)
                {
                    outlookInternalVersion = BitConverter.ToUInt32(outlookInternalVersionProperty.Value, 0);
                }
            }

            uint commonStartTimeNamedPropertyId = 0x8516;
            byte[] commonStartTimeNamedPropertyGuid = StandardPropertySet.Common;

            string commonStartTimeHex = Util.ConvertNamedPropertyToHex(commonStartTimeNamedPropertyId, commonStartTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(commonStartTimeHex) && namedPropertyIdTable[commonStartTimeHex] != null)
            {
                string commonStartTimePropertyId = namedPropertyIdTable[commonStartTimeHex];

                commonStartTimePropertyId = commonStartTimePropertyId + "0040";

                Property commonStartTimeProperty = propertyTable.ContainsKey(commonStartTimePropertyId) ? propertyTable[commonStartTimePropertyId] : null;

                if (commonStartTimeProperty != null && commonStartTimeProperty.Value != null)
                {
                    uint commonStartTimeLow = BitConverter.ToUInt32(commonStartTimeProperty.Value, 0);
                    ulong commonStartTimeHigh = BitConverter.ToUInt32(commonStartTimeProperty.Value, 4);

                    if (commonStartTimeHigh > 0)
                    {
                        long ticks = commonStartTimeLow + (long)(commonStartTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            commonStartTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint commonEndTimeNamedPropertyId = 0x8517;
            byte[] commonEndTimeNamedPropertyGuid = StandardPropertySet.Common;

            string commonEndTimeHex = Util.ConvertNamedPropertyToHex(commonEndTimeNamedPropertyId, commonEndTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(commonEndTimeHex) && namedPropertyIdTable[commonEndTimeHex] != null)
            {
                string commonEndTimePropertyId = namedPropertyIdTable[commonEndTimeHex];

                commonEndTimePropertyId = commonEndTimePropertyId + "0040";

                Property commonEndTimeProperty = propertyTable.ContainsKey(commonEndTimePropertyId) ? propertyTable[commonEndTimePropertyId] : null;

                if (commonEndTimeProperty != null && commonEndTimeProperty.Value != null)
                {
                    uint commonEndTimeLow = BitConverter.ToUInt32(commonEndTimeProperty.Value, 0);
                    ulong commonEndTimeHigh = BitConverter.ToUInt32(commonEndTimeProperty.Value, 4);

                    if (commonEndTimeHigh > 0)
                    {
                        long ticks = commonEndTimeLow + (long)(commonEndTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            commonEndTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint flagDueByNamedPropertyId = 0x8560;
            byte[] flagDueByNamedPropertyGuid = StandardPropertySet.Common;

            string flagDueByHex = Util.ConvertNamedPropertyToHex(flagDueByNamedPropertyId, flagDueByNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(flagDueByHex) && namedPropertyIdTable[flagDueByHex] != null)
            {
                string flagDueByPropertyId = namedPropertyIdTable[flagDueByHex];

                flagDueByPropertyId = flagDueByPropertyId + "0040";

                Property flagDueByProperty = propertyTable.ContainsKey(flagDueByPropertyId) ? propertyTable[flagDueByPropertyId] : null;

                if (flagDueByProperty != null && flagDueByProperty.Value != null)
                {
                    uint flagDueByLow = BitConverter.ToUInt32(flagDueByProperty.Value, 0);
                    ulong flagDueByHigh = BitConverter.ToUInt32(flagDueByProperty.Value, 4);

                    if (flagDueByHigh > 0)
                    {
                        long ticks = flagDueByLow + (long)(flagDueByHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            flagDueBy = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint companiesNamedPropertyId = 0x8539;
            byte[] companiesNamedPropertyGuid = StandardPropertySet.Common;

            string companiesHex = Util.ConvertNamedPropertyToHex(companiesNamedPropertyId, companiesNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(companiesHex) && namedPropertyIdTable[companiesHex] != null)
            {
                string companiesPropertyId = namedPropertyIdTable[companiesHex];

                companiesPropertyId = companiesPropertyId + multiValueStringTypeMask;

                Property companiesLenghtProperty = propertyTable.ContainsKey(companiesPropertyId) ? propertyTable[companiesPropertyId] : null;

                if (companiesLenghtProperty != null && companiesLenghtProperty.Value != null)
                {
                    uint companiesLenghtValue = companiesLenghtProperty.Size / 4;

                    companies = new List<String>();

                    for (int i = 0; i < companiesLenghtValue; i++)
                    {
                        string valueStreamName = "__substg1.0_" + companiesPropertyId + "-" + String.Format("{0:X8}", i);

                        Independentsoft.IO.StructuredStorage.Stream valueStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[valueStreamName];

                        if (valueStream != null && valueStream.Buffer != null)
                        {
                            companies.Add(encoding.GetString(valueStream.Buffer, 0, valueStream.Buffer.Length - encoding.GetBytes("\0").Length));
                        }
                    }
                }
            }

            uint contactNamesNamedPropertyId = 0x853A;
            byte[] contactNamesNamedPropertyGuid = StandardPropertySet.Common;

            string contactNamesHex = Util.ConvertNamedPropertyToHex(contactNamesNamedPropertyId, contactNamesNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(contactNamesHex) && namedPropertyIdTable[contactNamesHex] != null)
            {
                string contactNamesPropertyId = namedPropertyIdTable[contactNamesHex];

                contactNamesPropertyId = contactNamesPropertyId + multiValueStringTypeMask;

                Property contactNamesLenghtProperty = propertyTable.ContainsKey(contactNamesPropertyId) ? propertyTable[contactNamesPropertyId] : null;

                if (contactNamesLenghtProperty != null && contactNamesLenghtProperty.Value != null)
                {
                    uint contactNamesLenghtValue = contactNamesLenghtProperty.Size / 4;

                    contactNames = new List<String>();

                    for (int i = 0; i < contactNamesLenghtValue; i++)
                    {
                        string valueStreamName = "__substg1.0_" + contactNamesPropertyId + "-" + String.Format("{0:X8}", i);

                        Independentsoft.IO.StructuredStorage.Stream valueStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[valueStreamName];

                        if (valueStream != null && valueStream.Buffer != null)
                        {
                            contactNames.Add(encoding.GetString(valueStream.Buffer, 0, valueStream.Buffer.Length - encoding.GetBytes("\0").Length));
                        }
                    }
                }
            }

            string keywordsNamedPropertyId = "Keywords";
            byte[] keywordsNamedPropertyGuid = StandardPropertySet.PublicStrings;

            string keywordsHex = Util.ConvertNamedPropertyToHex(keywordsNamedPropertyId, keywordsNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(keywordsHex) && namedPropertyIdTable[keywordsHex] != null)
            {
                string keywordsPropertyId = namedPropertyIdTable[keywordsHex];

                keywordsPropertyId = keywordsPropertyId + multiValueStringTypeMask;

                Property keywordsLenghtProperty = propertyTable.ContainsKey(keywordsPropertyId) ? propertyTable[keywordsPropertyId] : null;

                if (keywordsLenghtProperty != null && keywordsLenghtProperty.Value != null)
                {
                    uint keywordsLenghtValue = keywordsLenghtProperty.Size / 4;

                    keywords = new List<String>();

                    for (int i = 0; i < keywordsLenghtValue; i++)
                    {
                        string valueStreamName = "__substg1.0_" + keywordsPropertyId + "-" + String.Format("{0:X8}", i);

                        Independentsoft.IO.StructuredStorage.Stream valueStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[valueStreamName];

                        int encodingLength = encoding.GetBytes("\0").Length;

                        if (valueStream != null && valueStream.Buffer != null && valueStream.Buffer.Length >= encodingLength)
                        {
                            keywords.Add(encoding.GetString(valueStream.Buffer, 0, valueStream.Buffer.Length - encodingLength));
                        }
                    }
                }
            }

            uint billingInformationNamedPropertyId = 0x8535;
            byte[] billingInformationNamedPropertyGuid = StandardPropertySet.Common;

            string billingInformationHex = Util.ConvertNamedPropertyToHex(billingInformationNamedPropertyId, billingInformationNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(billingInformationHex) && namedPropertyIdTable[billingInformationHex] != null)
            {
                string billingInformationPropertyId = namedPropertyIdTable[billingInformationHex];

                billingInformationPropertyId = billingInformationPropertyId + stringTypeMask;

                Property billingInformationProperty = propertyTable.ContainsKey(billingInformationPropertyId) ? propertyTable[billingInformationPropertyId] : null;

                if (billingInformationProperty != null && billingInformationProperty.Value != null)
                {
                    billingInformation = encoding.GetString(billingInformationProperty.Value, 0, billingInformationProperty.Value.Length);
                }
            }

            uint mileageNamedPropertyId = 0x8534;
            byte[] mileageNamedPropertyGuid = StandardPropertySet.Common;

            string mileageHex = Util.ConvertNamedPropertyToHex(mileageNamedPropertyId, mileageNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(mileageHex) && namedPropertyIdTable[mileageHex] != null)
            {
                string mileagePropertyId = namedPropertyIdTable[mileageHex];

                mileagePropertyId = mileagePropertyId + stringTypeMask;

                Property mileageProperty = propertyTable.ContainsKey(mileagePropertyId) ? propertyTable[mileagePropertyId] : null;

                if (mileageProperty != null && mileageProperty.Value != null)
                {
                    mileage = encoding.GetString(mileageProperty.Value, 0, mileageProperty.Value.Length);
                }
            }

            uint internetAccountNameNamedPropertyId = 0x8580;
            byte[] internetAccountNameNamedPropertyGuid = StandardPropertySet.Common;

            string internetAccountNameHex = Util.ConvertNamedPropertyToHex(internetAccountNameNamedPropertyId, internetAccountNameNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(internetAccountNameHex) && namedPropertyIdTable[internetAccountNameHex] != null)
            {
                string internetAccountNamePropertyId = namedPropertyIdTable[internetAccountNameHex];

                internetAccountNamePropertyId = internetAccountNamePropertyId + stringTypeMask;

                Property internetAccountNameProperty = propertyTable.ContainsKey(internetAccountNamePropertyId) ? propertyTable[internetAccountNamePropertyId] : null;

                if (internetAccountNameProperty != null && internetAccountNameProperty.Value != null)
                {
                    internetAccountName = encoding.GetString(internetAccountNameProperty.Value, 0, internetAccountNameProperty.Value.Length);
                }
            }

            uint reminderSoundFileNamedPropertyId = 0x851F;
            byte[] reminderSoundFileNamedPropertyGuid = StandardPropertySet.Common;

            string reminderSoundFileHex = Util.ConvertNamedPropertyToHex(reminderSoundFileNamedPropertyId, reminderSoundFileNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(reminderSoundFileHex) && namedPropertyIdTable[reminderSoundFileHex] != null)
            {
                string reminderSoundFilePropertyId = namedPropertyIdTable[reminderSoundFileHex];

                reminderSoundFilePropertyId = reminderSoundFilePropertyId + stringTypeMask;

                Property reminderSoundFileProperty = propertyTable.ContainsKey(reminderSoundFilePropertyId) ? propertyTable[reminderSoundFilePropertyId] : null;

                if (reminderSoundFileProperty != null && reminderSoundFileProperty.Value != null)
                {
                    reminderSoundFile = encoding.GetString(reminderSoundFileProperty.Value, 0, reminderSoundFileProperty.Value.Length);
                }
            }

            uint isPrivateNamedPropertyId = 0x8506;
            byte[] isPrivateNamedPropertyGuid = StandardPropertySet.Common;

            string isPrivateHex = Util.ConvertNamedPropertyToHex(isPrivateNamedPropertyId, isPrivateNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isPrivateHex) && namedPropertyIdTable[isPrivateHex] != null)
            {
                string isPrivatePropertyId = namedPropertyIdTable[isPrivateHex];

                isPrivatePropertyId = isPrivatePropertyId + "000B";

                Property isPrivateProperty = propertyTable.ContainsKey(isPrivatePropertyId) ? propertyTable[isPrivatePropertyId] : null;

                if (isPrivateProperty != null && isPrivateProperty.Value != null)
                {
                    ushort isPrivateValue = BitConverter.ToUInt16(isPrivateProperty.Value, 0);

                    if (isPrivateValue > 0)
                    {
                        isPrivate = true;
                    }
                }
            }

            uint reminderOverrideDefaultNamedPropertyId = 0x851C;
            byte[] reminderOverrideDefaultNamedPropertyGuid = StandardPropertySet.Common;

            string reminderOverrideDefaultHex = Util.ConvertNamedPropertyToHex(reminderOverrideDefaultNamedPropertyId, reminderOverrideDefaultNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(reminderOverrideDefaultHex) && namedPropertyIdTable[reminderOverrideDefaultHex] != null)
            {
                string reminderOverrideDefaultPropertyId = namedPropertyIdTable[reminderOverrideDefaultHex];

                reminderOverrideDefaultPropertyId = reminderOverrideDefaultPropertyId + "000B";

                Property reminderOverrideDefaultProperty = propertyTable.ContainsKey(reminderOverrideDefaultPropertyId) ? propertyTable[reminderOverrideDefaultPropertyId] : null;

                if (reminderOverrideDefaultProperty != null && reminderOverrideDefaultProperty.Value != null)
                {
                    ushort reminderOverrideDefaultValue = BitConverter.ToUInt16(reminderOverrideDefaultProperty.Value, 0);

                    if (reminderOverrideDefaultValue > 0)
                    {
                        reminderOverrideDefault = true;
                    }
                }
            }

            uint reminderPlaySoundNamedPropertyId = 0x851E;
            byte[] reminderPlaySoundNamedPropertyGuid = StandardPropertySet.Common;

            string reminderPlaySoundHex = Util.ConvertNamedPropertyToHex(reminderPlaySoundNamedPropertyId, reminderPlaySoundNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(reminderPlaySoundHex) && namedPropertyIdTable[reminderPlaySoundHex] != null)
            {
                string reminderPlaySoundPropertyId = namedPropertyIdTable[reminderPlaySoundHex];

                reminderPlaySoundPropertyId = reminderPlaySoundPropertyId + "000B";

                Property reminderPlaySoundProperty = propertyTable.ContainsKey(reminderPlaySoundPropertyId) ? propertyTable[reminderPlaySoundPropertyId] : null;

                if (reminderPlaySoundProperty != null && reminderPlaySoundProperty.Value != null)
                {
                    ushort reminderPlaySoundValue = BitConverter.ToUInt16(reminderPlaySoundProperty.Value, 0);

                    if (reminderPlaySoundValue > 0)
                    {
                        reminderPlaySound = true;
                    }
                }
            }

            uint appointmentStartTimeNamedPropertyId = 0x820D;
            byte[] appointmentStartTimeNamedPropertyGuid = StandardPropertySet.Appointment;

            string appointmentStartTimeHex = Util.ConvertNamedPropertyToHex(appointmentStartTimeNamedPropertyId, appointmentStartTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(appointmentStartTimeHex) && namedPropertyIdTable[appointmentStartTimeHex] != null)
            {
                string appointmentStartTimePropertyId = namedPropertyIdTable[appointmentStartTimeHex];

                appointmentStartTimePropertyId = appointmentStartTimePropertyId + "0040";

                Property appointmentStartTimeProperty = propertyTable.ContainsKey(appointmentStartTimePropertyId) ? propertyTable[appointmentStartTimePropertyId] : null;

                if (appointmentStartTimeProperty != null && appointmentStartTimeProperty.Value != null)
                {
                    uint appointmentStartTimeLow = BitConverter.ToUInt32(appointmentStartTimeProperty.Value, 0);
                    ulong appointmentStartTimeHigh = BitConverter.ToUInt32(appointmentStartTimeProperty.Value, 4);

                    if (appointmentStartTimeHigh > 0)
                    {
                        long ticks = appointmentStartTimeLow + (long)(appointmentStartTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            appointmentStartTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint appointmentEndTimeNamedPropertyId = 0x820E;
            byte[] appointmentEndTimeNamedPropertyGuid = StandardPropertySet.Appointment;

            string appointmentEndTimeHex = Util.ConvertNamedPropertyToHex(appointmentEndTimeNamedPropertyId, appointmentEndTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(appointmentEndTimeHex) && namedPropertyIdTable[appointmentEndTimeHex] != null)
            {
                string appointmentEndTimePropertyId = namedPropertyIdTable[appointmentEndTimeHex];

                appointmentEndTimePropertyId = appointmentEndTimePropertyId + "0040";

                Property appointmentEndTimeProperty = propertyTable.ContainsKey(appointmentEndTimePropertyId) ? propertyTable[appointmentEndTimePropertyId] : null;

                if (appointmentEndTimeProperty != null && appointmentEndTimeProperty.Value != null)
                {
                    uint appointmentEndTimeLow = BitConverter.ToUInt32(appointmentEndTimeProperty.Value, 0);
                    ulong appointmentEndTimeHigh = BitConverter.ToUInt32(appointmentEndTimeProperty.Value, 4);

                    if (appointmentEndTimeHigh > 0)
                    {
                        long ticks = appointmentEndTimeLow + (long)(appointmentEndTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            appointmentEndTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint locationNamedPropertyId = 0x8208;
            byte[] locationNamedPropertyGuid = StandardPropertySet.Appointment;

            string locationHex = Util.ConvertNamedPropertyToHex(locationNamedPropertyId, locationNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(locationHex) && namedPropertyIdTable[locationHex] != null)
            {
                string locationPropertyId = namedPropertyIdTable[locationHex];

                locationPropertyId = locationPropertyId + stringTypeMask;

                Property locationProperty = propertyTable.ContainsKey(locationPropertyId) ? propertyTable[locationPropertyId] : null;

                if (locationProperty != null && locationProperty.Value != null)
                {
                    location = encoding.GetString(locationProperty.Value, 0, locationProperty.Value.Length);
                }
            }

            uint appointmentMessageClassNamedPropertyId = 0x24;
            byte[] appointmentMessageClassNamedPropertyGuid = new byte[] { 144, 218, 216, 110, 11, 69, 27, 16, 152, 218, 0, 170, 0, 63, 19, 5 };

            string appointmentMessageClassHex = Util.ConvertNamedPropertyToHex(appointmentMessageClassNamedPropertyId, appointmentMessageClassNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(appointmentMessageClassHex) && namedPropertyIdTable[appointmentMessageClassHex] != null)
            {
                string appointmentMessageClassPropertyId = namedPropertyIdTable[appointmentMessageClassHex];

                appointmentMessageClassPropertyId = appointmentMessageClassPropertyId + stringTypeMask;

                Property appointmentMessageClassProperty = propertyTable.ContainsKey(appointmentMessageClassPropertyId) ? propertyTable[appointmentMessageClassPropertyId] : null;

                if (appointmentMessageClassProperty != null && appointmentMessageClassProperty.Value != null)
                {
                    appointmentMessageClass = encoding.GetString(appointmentMessageClassProperty.Value, 0, appointmentMessageClassProperty.Value.Length);
                }
            }

            uint timeZoneNamedPropertyId = 0x8234;
            byte[] timeZoneNamedPropertyGuid = StandardPropertySet.Appointment;

            string timeZoneHex = Util.ConvertNamedPropertyToHex(timeZoneNamedPropertyId, timeZoneNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(timeZoneHex) && namedPropertyIdTable[timeZoneHex] != null)
            {
                string timeZonePropertyId = namedPropertyIdTable[timeZoneHex];

                timeZonePropertyId = timeZonePropertyId + stringTypeMask;

                Property timeZoneProperty = propertyTable.ContainsKey(timeZonePropertyId) ? propertyTable[timeZonePropertyId] : null;

                if (timeZoneProperty != null && timeZoneProperty.Value != null)
                {
                    timeZone = encoding.GetString(timeZoneProperty.Value, 0, timeZoneProperty.Value.Length);
                }
            }

            uint recurrencePatternDescriptionNamedPropertyId = 0x8232;
            byte[] recurrencePatternDescriptionNamedPropertyGuid = StandardPropertySet.Appointment;

            string recurrencePatternDescriptionHex = Util.ConvertNamedPropertyToHex(recurrencePatternDescriptionNamedPropertyId, recurrencePatternDescriptionNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(recurrencePatternDescriptionHex) && namedPropertyIdTable[recurrencePatternDescriptionHex] != null)
            {
                string recurrencePatternDescriptionPropertyId = namedPropertyIdTable[recurrencePatternDescriptionHex];

                recurrencePatternDescriptionPropertyId = recurrencePatternDescriptionPropertyId + stringTypeMask;

                Property recurrencePatternDescriptionProperty = propertyTable.ContainsKey(recurrencePatternDescriptionPropertyId) ? propertyTable[recurrencePatternDescriptionPropertyId] : null;

                if (recurrencePatternDescriptionProperty != null && recurrencePatternDescriptionProperty.Value != null)
                {
                    recurrencePatternDescription = encoding.GetString(recurrencePatternDescriptionProperty.Value, 0, recurrencePatternDescriptionProperty.Value.Length);
                }
            }

            //appointment recurrence pattern
            uint recurrencePatternNamedPropertyId = 0x8216;
            byte[] recurrencePatternNamedPropertyGuid = StandardPropertySet.Appointment;
            
            string recurrencePatternHex = Util.ConvertNamedPropertyToHex(recurrencePatternNamedPropertyId, recurrencePatternNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(recurrencePatternHex) && namedPropertyIdTable[recurrencePatternHex] != null)
            {
                string recurrencePatternPropertyId = namedPropertyIdTable[recurrencePatternHex];

                recurrencePatternPropertyId = recurrencePatternPropertyId + "0102";

                Property recurrencePatternProperty = propertyTable.ContainsKey(recurrencePatternPropertyId) ? propertyTable[recurrencePatternPropertyId] : null;

                if (recurrencePatternProperty != null && recurrencePatternProperty.Value != null)
                {
                    recurrencePattern = new RecurrencePattern(recurrencePatternProperty.Value);
                }
            }

            //task recurrence pattern
            uint taskRecurrencePatternNamedPropertyId = 0x8116;
            byte[] taskRecurrencePatternNamedPropertyGuid = StandardPropertySet.Task;

            string taskRecurrencePatternHex = Util.ConvertNamedPropertyToHex(taskRecurrencePatternNamedPropertyId, taskRecurrencePatternNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskRecurrencePatternHex) && namedPropertyIdTable[taskRecurrencePatternHex] != null)
            {
                string recurrencePatternPropertyId = namedPropertyIdTable[taskRecurrencePatternHex];

                recurrencePatternPropertyId = recurrencePatternPropertyId + "0102";

                Property recurrencePatternProperty = propertyTable.ContainsKey(recurrencePatternPropertyId) ? propertyTable[recurrencePatternPropertyId] : null;

                if (recurrencePatternProperty != null && recurrencePatternProperty.Value != null)
                {
                    recurrencePattern = new RecurrencePattern(recurrencePatternProperty.Value);
                }
            }

            uint busyStatusNamedPropertyId = 0x8205;
            byte[] busyStatusNamedPropertyGuid = StandardPropertySet.Appointment;

            string busyStatusHex = Util.ConvertNamedPropertyToHex(busyStatusNamedPropertyId, busyStatusNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(busyStatusHex) && namedPropertyIdTable[busyStatusHex] != null)
            {
                string busyStatusPropertyId = namedPropertyIdTable[busyStatusHex];

                busyStatusPropertyId = busyStatusPropertyId + "0003";

                Property busyStatusProperty = propertyTable.ContainsKey(busyStatusPropertyId) ? propertyTable[busyStatusPropertyId] : null;

                if (busyStatusProperty != null && busyStatusProperty.Value != null)
                {
                    uint busyStatusValue = BitConverter.ToUInt32(busyStatusProperty.Value, 0);

                    busyStatus = EnumUtil.ParseBusyStatus(busyStatusValue);
                }
            }

            uint meetingStatusNamedPropertyId = 0x8217;
            byte[] meetingStatusNamedPropertyGuid = StandardPropertySet.Appointment;

            string meetingStatusHex = Util.ConvertNamedPropertyToHex(meetingStatusNamedPropertyId, meetingStatusNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(meetingStatusHex) && namedPropertyIdTable[meetingStatusHex] != null)
            {
                string meetingStatusPropertyId = namedPropertyIdTable[meetingStatusHex];

                meetingStatusPropertyId = meetingStatusPropertyId + "0003";

                Property meetingStatusProperty = propertyTable.ContainsKey(meetingStatusPropertyId) ? propertyTable[meetingStatusPropertyId] : null;

                if (meetingStatusProperty != null && meetingStatusProperty.Value != null)
                {
                    uint meetingStatusValue = BitConverter.ToUInt32(meetingStatusProperty.Value, 0);

                    meetingStatus = EnumUtil.ParseMeetingStatus(meetingStatusValue);
                }
            }

            uint responseStatusNamedPropertyId = 0x8218;
            byte[] responseStatusNamedPropertyGuid = StandardPropertySet.Appointment;

            string responseStatusHex = Util.ConvertNamedPropertyToHex(responseStatusNamedPropertyId, responseStatusNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(responseStatusHex) && namedPropertyIdTable[responseStatusHex] != null)
            {
                string responseStatusPropertyId = namedPropertyIdTable[responseStatusHex];

                responseStatusPropertyId = responseStatusPropertyId + "0003";

                Property responseStatusProperty = propertyTable.ContainsKey(responseStatusPropertyId) ? propertyTable[responseStatusPropertyId] : null;

                if (responseStatusProperty != null && responseStatusProperty.Value != null)
                {
                    uint responseStatusValue = BitConverter.ToUInt32(responseStatusProperty.Value, 0);

                    responseStatus = EnumUtil.ParseResponseStatus(responseStatusValue);
                }
            }

            uint recurrenceTypeNamedPropertyId = 0x8231;
            byte[] recurrenceTypeNamedPropertyGuid = StandardPropertySet.Appointment;

            string recurrenceTypeHex = Util.ConvertNamedPropertyToHex(recurrenceTypeNamedPropertyId, recurrenceTypeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(recurrenceTypeHex) && namedPropertyIdTable[recurrenceTypeHex] != null)
            {
                string recurrenceTypePropertyId = namedPropertyIdTable[recurrenceTypeHex];

                recurrenceTypePropertyId = recurrenceTypePropertyId + "0003";

                Property recurrenceTypeProperty = propertyTable.ContainsKey(recurrenceTypePropertyId) ? propertyTable[recurrenceTypePropertyId] : null;

                if (recurrenceTypeProperty != null && recurrenceTypeProperty.Value != null)
                {
                    uint recurrenceTypeValue = BitConverter.ToUInt32(recurrenceTypeProperty.Value, 0);

                    recurrenceType = EnumUtil.ParseRecurrenceType(recurrenceTypeValue);
                }
            }

            uint guidNamedPropertyId = 0x3;
            byte[] guidNamedPropertyGuid = new byte[] { 144, 218, 216, 110, 11, 69, 27, 16, 152, 218, 0, 170, 0, 63, 19, 5 };

            string guidHex = Util.ConvertNamedPropertyToHex(guidNamedPropertyId, guidNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(guidHex) && namedPropertyIdTable[guidHex] != null)
            {
                string guidPropertyId = namedPropertyIdTable[guidHex];

                guidPropertyId = guidPropertyId + "0102";

                Property guidProperty = propertyTable.ContainsKey(guidPropertyId) ? propertyTable[guidPropertyId] : null;

                if (guidProperty != null && guidProperty.Value != null)
                {
                    guid = guidProperty.Value;
                }
            }

            uint labelNamedPropertyId = 0x8214;
            byte[] labelNamedPropertyGuid = StandardPropertySet.Appointment;

            string labelHex = Util.ConvertNamedPropertyToHex(labelNamedPropertyId, labelNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(labelHex) && namedPropertyIdTable[labelHex] != null)
            {
                string labelPropertyId = namedPropertyIdTable[labelHex];

                labelPropertyId = labelPropertyId + "0003";

                Property labelProperty = propertyTable.ContainsKey(labelPropertyId) ? propertyTable[labelPropertyId] : null;

                if (labelProperty != null && labelProperty.Value != null)
                {
                    label = BitConverter.ToInt32(labelProperty.Value, 0);
                }
            }

            uint durationNamedPropertyId = 0x8213;
            byte[] durationNamedPropertyGuid = StandardPropertySet.Appointment;

            string durationHex = Util.ConvertNamedPropertyToHex(durationNamedPropertyId, durationNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(durationHex) && namedPropertyIdTable[durationHex] != null)
            {
                string durationPropertyId = namedPropertyIdTable[durationHex];

                durationPropertyId = durationPropertyId + "0003";

                Property durationProperty = propertyTable.ContainsKey(durationPropertyId) ? propertyTable[durationPropertyId] : null;

                if (durationProperty != null && durationProperty.Value != null)
                {
                    duration = BitConverter.ToUInt32(durationProperty.Value, 0);
                }
            }

            uint ownerNamedPropertyId = 0x811F;
            byte[] ownerNamedPropertyGuid = StandardPropertySet.Task;

            string ownerHex = Util.ConvertNamedPropertyToHex(ownerNamedPropertyId, ownerNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(ownerHex) && namedPropertyIdTable[ownerHex] != null)
            {
                string ownerPropertyId = namedPropertyIdTable[ownerHex];

                ownerPropertyId = ownerPropertyId + stringTypeMask;

                Property ownerProperty = propertyTable.ContainsKey(ownerPropertyId) ? propertyTable[ownerPropertyId] : null;

                if (ownerProperty != null && ownerProperty.Value != null)
                {
                    owner = encoding.GetString(ownerProperty.Value, 0, ownerProperty.Value.Length);
                }
            }

            uint delegatorNamedPropertyId = 0x8121;
            byte[] delegatorNamedPropertyGuid = StandardPropertySet.Task;

            string delegatorHex = Util.ConvertNamedPropertyToHex(delegatorNamedPropertyId, delegatorNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(delegatorHex) && namedPropertyIdTable[delegatorHex] != null)
            {
                string delegatorPropertyId = namedPropertyIdTable[delegatorHex];

                delegatorPropertyId = delegatorPropertyId + stringTypeMask;

                Property delegatorProperty = propertyTable.ContainsKey(delegatorPropertyId) ? propertyTable[delegatorPropertyId] : null;

                if (delegatorProperty != null && delegatorProperty.Value != null)
                {
                    delegator = encoding.GetString(delegatorProperty.Value, 0, delegatorProperty.Value.Length);
                }
            }

            uint percentCompleteNamedPropertyId = 0x8102;
            byte[] percentCompleteNamedPropertyGuid = StandardPropertySet.Task;

            string percentCompleteHex = Util.ConvertNamedPropertyToHex(percentCompleteNamedPropertyId, percentCompleteNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(percentCompleteHex) && namedPropertyIdTable[percentCompleteHex] != null)
            {
                string percentCompletePropertyId = namedPropertyIdTable[percentCompleteHex];

                percentCompletePropertyId = percentCompletePropertyId + "0005";

                Property percentCompleteProperty = propertyTable.ContainsKey(percentCompletePropertyId) ? propertyTable[percentCompletePropertyId] : null;

                if (percentCompleteProperty != null && percentCompleteProperty.Value != null)
                {
                    percentComplete = BitConverter.ToDouble(percentCompleteProperty.Value, 0);
                    percentComplete = percentComplete * 100;
                }
            }

            uint actualWorkNamedPropertyId = 0x8110;
            byte[] actualWorkNamedPropertyGuid = StandardPropertySet.Task;

            string actualWorkHex = Util.ConvertNamedPropertyToHex(actualWorkNamedPropertyId, actualWorkNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(actualWorkHex) && namedPropertyIdTable[actualWorkHex] != null)
            {
                string actualWorkPropertyId = namedPropertyIdTable[actualWorkHex];

                actualWorkPropertyId = actualWorkPropertyId + "0003";

                Property actualWorkProperty = propertyTable.ContainsKey(actualWorkPropertyId) ? propertyTable[actualWorkPropertyId] : null;

                if (actualWorkProperty != null && actualWorkProperty.Value != null)
                {
                    actualWork = BitConverter.ToUInt32(actualWorkProperty.Value, 0);
                }
            }

            uint totalWorkNamedPropertyId = 0x8111;
            byte[] totalWorkNamedPropertyGuid = StandardPropertySet.Task;

            string totalWorkHex = Util.ConvertNamedPropertyToHex(totalWorkNamedPropertyId, totalWorkNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(totalWorkHex) && namedPropertyIdTable[totalWorkHex] != null)
            {
                string totalWorkPropertyId = namedPropertyIdTable[totalWorkHex];

                totalWorkPropertyId = totalWorkPropertyId + "0003";

                Property totalWorkProperty = propertyTable.ContainsKey(totalWorkPropertyId) ? propertyTable[totalWorkPropertyId] : null;

                if (totalWorkProperty != null && totalWorkProperty.Value != null)
                {
                    totalWork = BitConverter.ToUInt32(totalWorkProperty.Value, 0);
                }
            }

            uint isTeamTaskNamedPropertyId = 0x8103;
            byte[] isTeamTaskNamedPropertyGuid = StandardPropertySet.Task;

            string isTeamTaskHex = Util.ConvertNamedPropertyToHex(isTeamTaskNamedPropertyId, isTeamTaskNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isTeamTaskHex) && namedPropertyIdTable[isTeamTaskHex] != null)
            {
                string isTeamTaskPropertyId = namedPropertyIdTable[isTeamTaskHex];

                isTeamTaskPropertyId = isTeamTaskPropertyId + "000B";

                Property isTeamTaskProperty = propertyTable.ContainsKey(isTeamTaskPropertyId) ? propertyTable[isTeamTaskPropertyId] : null;

                if (isTeamTaskProperty != null && isTeamTaskProperty.Value != null)
                {
                    ushort isTeamTaskValue = BitConverter.ToUInt16(isTeamTaskProperty.Value, 0);

                    if (isTeamTaskValue > 0)
                    {
                        isTeamTask = true;
                    }
                }
            }

            uint isCompleteNamedPropertyId = 0x811C;
            byte[] isCompleteNamedPropertyGuid = StandardPropertySet.Task;

            string isCompleteHex = Util.ConvertNamedPropertyToHex(isCompleteNamedPropertyId, isCompleteNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isCompleteHex) && namedPropertyIdTable[isCompleteHex] != null)
            {
                string isCompletePropertyId = namedPropertyIdTable[isCompleteHex];

                isCompletePropertyId = isCompletePropertyId + "000B";

                Property isCompleteProperty = propertyTable.ContainsKey(isCompletePropertyId) ? propertyTable[isCompletePropertyId] : null;

                if (isCompleteProperty != null && isCompleteProperty.Value != null)
                {
                    ushort isCompleteValue = BitConverter.ToUInt16(isCompleteProperty.Value, 0);

                    if (isCompleteValue > 0)
                    {
                        isComplete = true;
                    }
                }
            }

            uint isRecurringNamedPropertyId = 0x8223;
            byte[] isRecurringNamedPropertyGuid = StandardPropertySet.Appointment;

            string isRecurringHex = Util.ConvertNamedPropertyToHex(isRecurringNamedPropertyId, isRecurringNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isRecurringHex) && namedPropertyIdTable[isRecurringHex] != null)
            {
                string isRecurringPropertyId = namedPropertyIdTable[isRecurringHex];

                isRecurringPropertyId = isRecurringPropertyId + "000B";

                Property isRecurringProperty = propertyTable.ContainsKey(isRecurringPropertyId) ? propertyTable[isRecurringPropertyId] : null;

                if (isRecurringProperty != null && isRecurringProperty.Value != null)
                {
                    ushort isRecurringValue = BitConverter.ToUInt16(isRecurringProperty.Value, 0);

                    if (isRecurringValue > 0)
                    {
                        isRecurring = true;
                    }
                }
            }

            uint isAllDayEventNamedPropertyId = 0x8215;
            byte[] isAllDayEventNamedPropertyGuid = StandardPropertySet.Appointment;

            string isAllDayEventHex = Util.ConvertNamedPropertyToHex(isAllDayEventNamedPropertyId, isAllDayEventNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isAllDayEventHex) && namedPropertyIdTable[isAllDayEventHex] != null)
            {
                string isAllDayEventPropertyId = namedPropertyIdTable[isAllDayEventHex];

                isAllDayEventPropertyId = isAllDayEventPropertyId + "000B";

                Property isAllDayEventProperty = propertyTable.ContainsKey(isAllDayEventPropertyId) ? propertyTable[isAllDayEventPropertyId] : null;

                if (isAllDayEventProperty != null && isAllDayEventProperty.Value != null)
                {
                    ushort isAllDayEventValue = BitConverter.ToUInt16(isAllDayEventProperty.Value, 0);

                    if (isAllDayEventValue > 0)
                    {
                        isAllDayEvent = true;
                    }
                }
            }

            uint isReminderSetNamedPropertyId = 0x8503;
            byte[] isReminderSetNamedPropertyGuid = StandardPropertySet.Common;

            string isReminderSetHex = Util.ConvertNamedPropertyToHex(isReminderSetNamedPropertyId, isReminderSetNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(isReminderSetHex) && namedPropertyIdTable[isReminderSetHex] != null)
            {
                string isReminderSetPropertyId = namedPropertyIdTable[isReminderSetHex];

                isReminderSetPropertyId = isReminderSetPropertyId + "000B";

                Property isReminderSetProperty = propertyTable.ContainsKey(isReminderSetPropertyId) ? propertyTable[isReminderSetPropertyId] : null;

                if (isReminderSetProperty != null && isReminderSetProperty.Value != null)
                {
                    ushort isReminderSetValue = BitConverter.ToUInt16(isReminderSetProperty.Value, 0);

                    if (isReminderSetValue > 0)
                    {
                        isReminderSet = true;
                    }
                }
            }

            uint reminderTimeNamedPropertyId = 0x8502;
            byte[] reminderTimeNamedPropertyGuid = StandardPropertySet.Common;

            string reminderTimeHex = Util.ConvertNamedPropertyToHex(reminderTimeNamedPropertyId, reminderTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(reminderTimeHex) && namedPropertyIdTable[reminderTimeHex] != null)
            {
                string reminderTimePropertyId = namedPropertyIdTable[reminderTimeHex];

                reminderTimePropertyId = reminderTimePropertyId + "0040";

                Property reminderTimeProperty = propertyTable.ContainsKey(reminderTimePropertyId) ? propertyTable[reminderTimePropertyId] : null;

                if (reminderTimeProperty != null && reminderTimeProperty.Value != null)
                {
                    uint reminderTimeLow = BitConverter.ToUInt32(reminderTimeProperty.Value, 0);
                    ulong reminderTimeHigh = BitConverter.ToUInt32(reminderTimeProperty.Value, 4);

                    if (reminderTimeHigh > 0)
                    {
                        long ticks = reminderTimeLow + (long)(reminderTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            reminderTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint reminderMinutesBeforeStartNamedPropertyId = 0x8501;
            byte[] reminderMinutesBeforeStartNamedPropertyGuid = StandardPropertySet.Common;

            string reminderMinutesBeforeStartHex = Util.ConvertNamedPropertyToHex(reminderMinutesBeforeStartNamedPropertyId, reminderMinutesBeforeStartNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(reminderMinutesBeforeStartHex) && namedPropertyIdTable[reminderMinutesBeforeStartHex] != null)
            {
                string reminderMinutesBeforeStartPropertyId = namedPropertyIdTable[reminderMinutesBeforeStartHex];

                reminderMinutesBeforeStartPropertyId = reminderMinutesBeforeStartPropertyId + "0003";

                Property reminderMinutesBeforeStartProperty = propertyTable.ContainsKey(reminderMinutesBeforeStartPropertyId) ? propertyTable[reminderMinutesBeforeStartPropertyId] : null;

                if (reminderMinutesBeforeStartProperty != null && reminderMinutesBeforeStartProperty.Value != null)
                {
                    reminderMinutesBeforeStart = BitConverter.ToUInt32(reminderMinutesBeforeStartProperty.Value, 0);
                }
            }

            uint taskStartDateNamedPropertyId = 0x8104;
            byte[] taskStartDateNamedPropertyGuid = StandardPropertySet.Task;

            string taskStartDateHex = Util.ConvertNamedPropertyToHex(taskStartDateNamedPropertyId, taskStartDateNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskStartDateHex) && namedPropertyIdTable[taskStartDateHex] != null)
            {
                string taskStartDatePropertyId = namedPropertyIdTable[taskStartDateHex];

                taskStartDatePropertyId = taskStartDatePropertyId + "0040";

                Property taskStartDateProperty = propertyTable.ContainsKey(taskStartDatePropertyId) ? propertyTable[taskStartDatePropertyId] : null;

                if (taskStartDateProperty != null && taskStartDateProperty.Value != null)
                {
                    uint taskStartDateLow = BitConverter.ToUInt32(taskStartDateProperty.Value, 0);
                    ulong taskStartDateHigh = BitConverter.ToUInt32(taskStartDateProperty.Value, 4);

                    if (taskStartDateHigh > 0)
                    {
                        long ticks = taskStartDateLow + (long)(taskStartDateHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            taskStartDate = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint taskDueDateNamedPropertyId = 0x8105;
            byte[] taskDueDateNamedPropertyGuid = StandardPropertySet.Task;

            string taskDueDateHex = Util.ConvertNamedPropertyToHex(taskDueDateNamedPropertyId, taskDueDateNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskDueDateHex) && namedPropertyIdTable[taskDueDateHex] != null)
            {
                string taskDueDatePropertyId = namedPropertyIdTable[taskDueDateHex];

                taskDueDatePropertyId = taskDueDatePropertyId + "0040";

                Property taskDueDateProperty = propertyTable.ContainsKey(taskDueDatePropertyId) ? propertyTable[taskDueDatePropertyId] : null;

                if (taskDueDateProperty != null && taskDueDateProperty.Value != null)
                {
                    uint taskDueDateLow = BitConverter.ToUInt32(taskDueDateProperty.Value, 0);
                    ulong taskDueDateHigh = BitConverter.ToUInt32(taskDueDateProperty.Value, 4);

                    if (taskDueDateHigh > 0)
                    {
                        long ticks = taskDueDateLow + (long)(taskDueDateHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            taskDueDate = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint dateCompletedNamedPropertyId = 0x810F;
            byte[] dateCompletedNamedPropertyGuid = StandardPropertySet.Task;

            string dateCompletedHex = Util.ConvertNamedPropertyToHex(dateCompletedNamedPropertyId, dateCompletedNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(dateCompletedHex) && namedPropertyIdTable[dateCompletedHex] != null)
            {
                string dateCompletedPropertyId = namedPropertyIdTable[dateCompletedHex];

                dateCompletedPropertyId = dateCompletedPropertyId + "0040";

                Property dateCompletedProperty = propertyTable.ContainsKey(dateCompletedPropertyId) ? propertyTable[dateCompletedPropertyId] : null;

                if (dateCompletedProperty != null && dateCompletedProperty.Value != null)
                {
                    uint dateCompletedLow = BitConverter.ToUInt32(dateCompletedProperty.Value, 0);
                    ulong dateCompletedHigh = BitConverter.ToUInt32(dateCompletedProperty.Value, 4);

                    if (dateCompletedHigh > 0)
                    {
                        long ticks = dateCompletedLow + (long)(dateCompletedHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            dateCompleted = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint taskStatusNamedPropertyId = 0x8101;
            byte[] taskStatusNamedPropertyGuid = StandardPropertySet.Task;

            string taskStatusHex = Util.ConvertNamedPropertyToHex(taskStatusNamedPropertyId, taskStatusNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskStatusHex) && namedPropertyIdTable[taskStatusHex] != null)
            {
                string taskStatusPropertyId = namedPropertyIdTable[taskStatusHex];

                taskStatusPropertyId = taskStatusPropertyId + "0003";

                Property taskStatusProperty = propertyTable.ContainsKey(taskStatusPropertyId) ? propertyTable[taskStatusPropertyId] : null;

                if (taskStatusProperty != null && taskStatusProperty.Value != null)
                {
                    uint taskStatusValue = BitConverter.ToUInt32(taskStatusProperty.Value, 0);

                    taskStatus = EnumUtil.ParseTaskStatus(taskStatusValue);
                }
            }

            uint taskOwnershipNamedPropertyId = 0x8129;
            byte[] taskOwnershipNamedPropertyGuid = StandardPropertySet.Task;

            string taskOwnershipHex = Util.ConvertNamedPropertyToHex(taskOwnershipNamedPropertyId, taskOwnershipNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskOwnershipHex) && namedPropertyIdTable[taskOwnershipHex] != null)
            {
                string taskOwnershipPropertyId = namedPropertyIdTable[taskOwnershipHex];

                taskOwnershipPropertyId = taskOwnershipPropertyId + "0003";

                Property taskOwnershipProperty = propertyTable.ContainsKey(taskOwnershipPropertyId) ? propertyTable[taskOwnershipPropertyId] : null;

                if (taskOwnershipProperty != null && taskOwnershipProperty.Value != null)
                {
                    uint taskOwnershipValue = BitConverter.ToUInt32(taskOwnershipProperty.Value, 0);

                    taskOwnership = EnumUtil.ParseTaskOwnership(taskOwnershipValue);
                }
            }

            uint taskDelegationStateNamedPropertyId = 0x812A;
            byte[] taskDelegationStateNamedPropertyGuid = StandardPropertySet.Task;

            string taskDelegationStateHex = Util.ConvertNamedPropertyToHex(taskDelegationStateNamedPropertyId, taskDelegationStateNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(taskDelegationStateHex) && namedPropertyIdTable[taskDelegationStateHex] != null)
            {
                string taskDelegationStatePropertyId = namedPropertyIdTable[taskDelegationStateHex];

                taskDelegationStatePropertyId = taskDelegationStatePropertyId + "0003";

                Property taskDelegationStateProperty = propertyTable.ContainsKey(taskDelegationStatePropertyId) ? propertyTable[taskDelegationStatePropertyId] : null;

                if (taskDelegationStateProperty != null && taskDelegationStateProperty.Value != null)
                {
                    uint taskDelegationStateValue = BitConverter.ToUInt32(taskDelegationStateProperty.Value, 0);

                    taskDelegationState = EnumUtil.ParseTaskDelegationState(taskDelegationStateValue);
                }
            }

            uint noteTopNamedPropertyId = 0x8B05;
            byte[] noteTopNamedPropertyGuid = StandardPropertySet.Note;

            string noteTopHex = Util.ConvertNamedPropertyToHex(noteTopNamedPropertyId, noteTopNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(noteTopHex) && namedPropertyIdTable[noteTopHex] != null)
            {
                string noteTopPropertyId = namedPropertyIdTable[noteTopHex];

                noteTopPropertyId = noteTopPropertyId + "0003";

                Property noteTopProperty = propertyTable.ContainsKey(noteTopPropertyId) ? propertyTable[noteTopPropertyId] : null;

                if (noteTopProperty != null && noteTopProperty.Value != null)
                {
                    noteTop = BitConverter.ToUInt32(noteTopProperty.Value, 0);
                }
            }

            uint noteLeftNamedPropertyId = 0x8B04;
            byte[] noteLeftNamedPropertyGuid = StandardPropertySet.Note;

            string noteLeftHex = Util.ConvertNamedPropertyToHex(noteLeftNamedPropertyId, noteLeftNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(noteLeftHex) && namedPropertyIdTable[noteLeftHex] != null)
            {
                string noteLeftPropertyId = namedPropertyIdTable[noteLeftHex];

                noteLeftPropertyId = noteLeftPropertyId + "0003";

                Property noteLeftProperty = propertyTable.ContainsKey(noteLeftPropertyId) ? propertyTable[noteLeftPropertyId] : null;

                if (noteLeftProperty != null && noteLeftProperty.Value != null)
                {
                    noteLeft = BitConverter.ToUInt32(noteLeftProperty.Value, 0);
                }
            }

            uint noteHeightNamedPropertyId = 0x8B03;
            byte[] noteHeightNamedPropertyGuid = StandardPropertySet.Note;

            string noteHeightHex = Util.ConvertNamedPropertyToHex(noteHeightNamedPropertyId, noteHeightNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(noteHeightHex) && namedPropertyIdTable[noteHeightHex] != null)
            {
                string noteHeightPropertyId = namedPropertyIdTable[noteHeightHex];

                noteHeightPropertyId = noteHeightPropertyId + "0003";

                Property noteHeightProperty = propertyTable.ContainsKey(noteHeightPropertyId) ? propertyTable[noteHeightPropertyId] : null;

                if (noteHeightProperty != null && noteHeightProperty.Value != null)
                {
                    noteHeight = BitConverter.ToUInt32(noteHeightProperty.Value, 0);
                }
            }

            uint noteWidthNamedPropertyId = 0x8B02;
            byte[] noteWidthNamedPropertyGuid = StandardPropertySet.Note;

            string noteWidthHex = Util.ConvertNamedPropertyToHex(noteWidthNamedPropertyId, noteWidthNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(noteWidthHex) && namedPropertyIdTable[noteWidthHex] != null)
            {
                string noteWidthPropertyId = namedPropertyIdTable[noteWidthHex];

                noteWidthPropertyId = noteWidthPropertyId + "0003";

                Property noteWidthProperty = propertyTable.ContainsKey(noteWidthPropertyId) ? propertyTable[noteWidthPropertyId] : null;

                if (noteWidthProperty != null && noteWidthProperty.Value != null)
                {
                    noteWidth = BitConverter.ToUInt32(noteWidthProperty.Value, 0);
                }
            }

            uint noteColorNamedPropertyId = 0x8B00;
            byte[] noteColorNamedPropertyGuid = StandardPropertySet.Note;

            string noteColorHex = Util.ConvertNamedPropertyToHex(noteColorNamedPropertyId, noteColorNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(noteColorHex) && namedPropertyIdTable[noteColorHex] != null)
            {
                string noteColorPropertyId = namedPropertyIdTable[noteColorHex];

                noteColorPropertyId = noteColorPropertyId + "0003";

                Property noteColorProperty = propertyTable.ContainsKey(noteColorPropertyId) ? propertyTable[noteColorPropertyId] : null;

                if (noteColorProperty != null && noteColorProperty.Value != null)
                {
                    uint noteColorValue = BitConverter.ToUInt32(noteColorProperty.Value, 0);

                    noteColor = EnumUtil.ParseNoteColor(noteColorValue);
                }
            }

            uint journalStartTimeNamedPropertyId = 0x8706;
            byte[] journalStartTimeNamedPropertyGuid = StandardPropertySet.Journal;

            string journalStartTimeHex = Util.ConvertNamedPropertyToHex(journalStartTimeNamedPropertyId, journalStartTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(journalStartTimeHex) && namedPropertyIdTable[journalStartTimeHex] != null)
            {
                string journalStartTimePropertyId = namedPropertyIdTable[journalStartTimeHex];

                journalStartTimePropertyId = journalStartTimePropertyId + "0040";

                Property journalStartTimeProperty = propertyTable.ContainsKey(journalStartTimePropertyId) ? propertyTable[journalStartTimePropertyId] : null;

                if (journalStartTimeProperty != null && journalStartTimeProperty.Value != null)
                {
                    uint journalStartTimeLow = BitConverter.ToUInt32(journalStartTimeProperty.Value, 0);
                    ulong journalStartTimeHigh = BitConverter.ToUInt32(journalStartTimeProperty.Value, 4);

                    if (journalStartTimeHigh > 0)
                    {
                        long ticks = journalStartTimeLow + (long)(journalStartTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            journalStartTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint journalEndTimeNamedPropertyId = 0x8708;
            byte[] journalEndTimeNamedPropertyGuid = StandardPropertySet.Journal;

            string journalEndTimeHex = Util.ConvertNamedPropertyToHex(journalEndTimeNamedPropertyId, journalEndTimeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(journalEndTimeHex) && namedPropertyIdTable[journalEndTimeHex] != null)
            {
                string journalEndTimePropertyId = namedPropertyIdTable[journalEndTimeHex];

                journalEndTimePropertyId = journalEndTimePropertyId + "0040";

                Property journalEndTimeProperty = propertyTable.ContainsKey(journalEndTimePropertyId) ? propertyTable[journalEndTimePropertyId] : null;

                if (journalEndTimeProperty != null && journalEndTimeProperty.Value != null)
                {
                    uint journalEndTimeLow = BitConverter.ToUInt32(journalEndTimeProperty.Value, 0);
                    ulong journalEndTimeHigh = BitConverter.ToUInt32(journalEndTimeProperty.Value, 4);

                    if (journalEndTimeHigh > 0)
                    {
                        long ticks = journalEndTimeLow + (long)(journalEndTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            journalEndTime = year1601.AddTicks(ticks).ToLocalTime();
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }
                }
            }

            uint journalTypeNamedPropertyId = 0x8700;
            byte[] journalTypeNamedPropertyGuid = StandardPropertySet.Journal;

            string journalTypeHex = Util.ConvertNamedPropertyToHex(journalTypeNamedPropertyId, journalTypeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(journalTypeHex) && namedPropertyIdTable[journalTypeHex] != null)
            {
                string journalTypePropertyId = namedPropertyIdTable[journalTypeHex];

                journalTypePropertyId = journalTypePropertyId + stringTypeMask;

                Property journalTypeProperty = propertyTable.ContainsKey(journalTypePropertyId) ? propertyTable[journalTypePropertyId] : null;

                if (journalTypeProperty != null && journalTypeProperty.Value != null)
                {
                    journalType = encoding.GetString(journalTypeProperty.Value, 0, journalTypeProperty.Value.Length);
                }
            }

            uint journalTypeDescriptionNamedPropertyId = 0x8712;
            byte[] journalTypeDescriptionNamedPropertyGuid = StandardPropertySet.Journal;

            string journalTypeDescriptionHex = Util.ConvertNamedPropertyToHex(journalTypeDescriptionNamedPropertyId, journalTypeDescriptionNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(journalTypeDescriptionHex) && namedPropertyIdTable[journalTypeDescriptionHex] != null)
            {
                string journalTypeDescriptionPropertyId = namedPropertyIdTable[journalTypeDescriptionHex];

                journalTypeDescriptionPropertyId = journalTypeDescriptionPropertyId + stringTypeMask;

                Property journalTypeDescriptionProperty = propertyTable.ContainsKey(journalTypeDescriptionPropertyId) ? propertyTable[journalTypeDescriptionPropertyId] : null;

                if (journalTypeDescriptionProperty != null && journalTypeDescriptionProperty.Value != null)
                {
                    journalTypeDescription = encoding.GetString(journalTypeDescriptionProperty.Value, 0, journalTypeDescriptionProperty.Value.Length);
                }
            }

            uint journalDurationNamedPropertyId = 0x8707;
            byte[] journalDurationNamedPropertyGuid = StandardPropertySet.Journal;

            string journalDurationHex = Util.ConvertNamedPropertyToHex(journalDurationNamedPropertyId, journalDurationNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(journalDurationHex) && namedPropertyIdTable[journalDurationHex] != null)
            {
                string journalDurationPropertyId = namedPropertyIdTable[journalDurationHex];

                journalDurationPropertyId = journalDurationPropertyId + "0003";

                Property journalDurationProperty = propertyTable.ContainsKey(journalDurationPropertyId) ? propertyTable[journalDurationPropertyId] : null;

                if (journalDurationProperty != null && journalDurationProperty.Value != null)
                {
                    journalDuration = BitConverter.ToUInt32(journalDurationProperty.Value, 0);
                }
            }

            Property birthdayProperty = propertyTable.ContainsKey("3A420040") ? propertyTable["3A420040"] : null;

            if (birthdayProperty != null && birthdayProperty.Value != null)
            {
                uint birthdayLow = BitConverter.ToUInt32(birthdayProperty.Value, 0);
                ulong birthdayHigh = BitConverter.ToUInt32(birthdayProperty.Value, 4);

                if (birthdayHigh > 0)
                {
                    long ticks = birthdayLow + (long)(birthdayHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        birthday = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property childrenNamesProperty = propertyTable.ContainsKey("3A58" + multiValueStringTypeMask) ? propertyTable["3A58" + multiValueStringTypeMask] : null;

            if (childrenNamesProperty != null && childrenNamesProperty.Value != null)
            {
                uint childrenNamesValue = childrenNamesProperty.Size / 4;

                childrenNames = new List<String>();

                for (int i = 0; i < childrenNamesValue; i++)
                {
                    string valueStreamName = "__substg1.0_3A58" + multiValueStringTypeMask + "-" + String.Format("{0:X8}", i);

                    Independentsoft.IO.StructuredStorage.Stream valueStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[valueStreamName];

                    if (valueStream != null && valueStream.Buffer != null)
                    {
                        childrenNames.Add(encoding.GetString(valueStream.Buffer, 0, valueStream.Buffer.Length - encoding.GetBytes("\0").Length));
                    }
                }
            }

            Property weddingAnniversaryProperty = propertyTable.ContainsKey("3A410040") ? propertyTable["3A410040"] : null;

            if (weddingAnniversaryProperty != null && weddingAnniversaryProperty.Value != null)
            {
                uint weddingAnniversaryLow = BitConverter.ToUInt32(weddingAnniversaryProperty.Value, 0);
                ulong weddingAnniversaryHigh = BitConverter.ToUInt32(weddingAnniversaryProperty.Value, 4);

                if (weddingAnniversaryHigh > 0)
                {
                    long ticks = weddingAnniversaryLow + (long)(weddingAnniversaryHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                        weddingAnniversary = year1601.AddTicks(ticks).ToLocalTime();
                    }
                    catch (Exception) //ignore wrong dates
                    {
                    }
                }
            }

            Property genderProperty = propertyTable.ContainsKey("3A4D0002") ? propertyTable["3A4D0002"] : null;

            if (genderProperty != null && genderProperty.Value != null)
            {
                short genderValue = BitConverter.ToInt16(genderProperty.Value, 0);

                gender = EnumUtil.ParseGender(genderValue);
            }

            uint selectedMailingAddressNamedPropertyId = 0x8022;
            byte[] selectedMailingAddressNamedPropertyGuid = StandardPropertySet.Address;

            string selectedMailingAddressHex = Util.ConvertNamedPropertyToHex(selectedMailingAddressNamedPropertyId, selectedMailingAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(selectedMailingAddressHex) && namedPropertyIdTable[selectedMailingAddressHex] != null)
            {
                string selectedMailingAddressPropertyId = namedPropertyIdTable[selectedMailingAddressHex];

                selectedMailingAddressPropertyId = selectedMailingAddressPropertyId + "0003";

                Property selectedMailingAddressProperty = propertyTable.ContainsKey(selectedMailingAddressPropertyId) ? propertyTable[selectedMailingAddressPropertyId] : null;

                if (selectedMailingAddressProperty != null && selectedMailingAddressProperty.Value != null)
                {
                    uint selectedMailingAddressValue = BitConverter.ToUInt32(selectedMailingAddressProperty.Value, 0);

                    selectedMailingAddress = EnumUtil.ParseSelectedMailingAddress(selectedMailingAddressValue);
                }
            }

            uint contactHasPictureNamedPropertyId = 0x8015;
            byte[] contactHasPictureNamedPropertyGuid = StandardPropertySet.Address;

            string contactHasPictureHex = Util.ConvertNamedPropertyToHex(contactHasPictureNamedPropertyId, contactHasPictureNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(contactHasPictureHex) && namedPropertyIdTable[contactHasPictureHex] != null)
            {
                string contactHasPicturePropertyId = namedPropertyIdTable[contactHasPictureHex];

                contactHasPicturePropertyId = contactHasPicturePropertyId + "000B";

                Property contactHasPictureProperty = propertyTable.ContainsKey(contactHasPicturePropertyId) ? propertyTable[contactHasPicturePropertyId] : null;

                if (contactHasPictureProperty != null && contactHasPictureProperty.Value != null)
                {
                    ushort contactHasPictureValue = BitConverter.ToUInt16(contactHasPictureProperty.Value, 0);

                    if (contactHasPictureValue > 0)
                    {
                        contactHasPicture = true;
                    }
                }
            }

            uint fileAsNamedPropertyId = 0x8005;
            byte[] fileAsNamedPropertyGuid = StandardPropertySet.Address;

            string fileAsHex = Util.ConvertNamedPropertyToHex(fileAsNamedPropertyId, fileAsNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(fileAsHex) && namedPropertyIdTable[fileAsHex] != null)
            {
                string fileAsPropertyId = namedPropertyIdTable[fileAsHex];

                fileAsPropertyId = fileAsPropertyId + stringTypeMask;

                Property fileAsProperty = propertyTable.ContainsKey(fileAsPropertyId) ? propertyTable[fileAsPropertyId] : null;

                if (fileAsProperty != null && fileAsProperty.Value != null)
                {
                    fileAs = encoding.GetString(fileAsProperty.Value, 0, fileAsProperty.Value.Length);
                }
            }

            uint instantMessengerAddressNamedPropertyId = 0x8062;
            byte[] instantMessengerAddressNamedPropertyGuid = StandardPropertySet.Address;

            string instantMessengerAddressHex = Util.ConvertNamedPropertyToHex(instantMessengerAddressNamedPropertyId, instantMessengerAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(instantMessengerAddressHex) && namedPropertyIdTable[instantMessengerAddressHex] != null)
            {
                string instantMessengerAddressPropertyId = namedPropertyIdTable[instantMessengerAddressHex];

                instantMessengerAddressPropertyId = instantMessengerAddressPropertyId + stringTypeMask;

                Property instantMessengerAddressProperty = propertyTable.ContainsKey(instantMessengerAddressPropertyId) ? propertyTable[instantMessengerAddressPropertyId] : null;

                if (instantMessengerAddressProperty != null && instantMessengerAddressProperty.Value != null)
                {
                    instantMessengerAddress = encoding.GetString(instantMessengerAddressProperty.Value, 0, instantMessengerAddressProperty.Value.Length);
                }
            }

            uint internetFreeBusyAddressNamedPropertyId = 0x80D8;
            byte[] internetFreeBusyAddressNamedPropertyGuid = StandardPropertySet.Address;

            string internetFreeBusyAddressHex = Util.ConvertNamedPropertyToHex(internetFreeBusyAddressNamedPropertyId, internetFreeBusyAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(internetFreeBusyAddressHex) && namedPropertyIdTable[internetFreeBusyAddressHex] != null)
            {
                string internetFreeBusyAddressPropertyId = namedPropertyIdTable[internetFreeBusyAddressHex];

                internetFreeBusyAddressPropertyId = internetFreeBusyAddressPropertyId + stringTypeMask;

                Property internetFreeBusyAddressProperty = propertyTable.ContainsKey(internetFreeBusyAddressPropertyId) ? propertyTable[internetFreeBusyAddressPropertyId] : null;

                if (internetFreeBusyAddressProperty != null && internetFreeBusyAddressProperty.Value != null)
                {
                    internetFreeBusyAddress = encoding.GetString(internetFreeBusyAddressProperty.Value, 0, internetFreeBusyAddressProperty.Value.Length);
                }
            }

            uint businessAddressNamedPropertyId = 0x801B;
            byte[] businessAddressNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressHex = Util.ConvertNamedPropertyToHex(businessAddressNamedPropertyId, businessAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressHex) && namedPropertyIdTable[businessAddressHex] != null)
            {
                string businessAddressPropertyId = namedPropertyIdTable[businessAddressHex];

                businessAddressPropertyId = businessAddressPropertyId + stringTypeMask;

                Property businessAddressProperty = propertyTable.ContainsKey(businessAddressPropertyId) ? propertyTable[businessAddressPropertyId] : null;

                if (businessAddressProperty != null && businessAddressProperty.Value != null)
                {
                    businessAddress = encoding.GetString(businessAddressProperty.Value, 0, businessAddressProperty.Value.Length);
                }
            }

            uint businessAddressStreetNamedPropertyId = 0x8045;
            byte[] businessAddressStreetNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressStreetHex = Util.ConvertNamedPropertyToHex(businessAddressStreetNamedPropertyId, businessAddressStreetNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressStreetHex) && namedPropertyIdTable[businessAddressStreetHex] != null)
            {
                string businessAddressStreetPropertyId = namedPropertyIdTable[businessAddressStreetHex];

                businessAddressStreetPropertyId = businessAddressStreetPropertyId + stringTypeMask;

                Property businessAddressStreetProperty = propertyTable.ContainsKey(businessAddressStreetPropertyId) ? propertyTable[businessAddressStreetPropertyId] : null;

                if (businessAddressStreetProperty != null && businessAddressStreetProperty.Value != null)
                {
                    businessAddressStreet = encoding.GetString(businessAddressStreetProperty.Value, 0, businessAddressStreetProperty.Value.Length);
                }
            }

            uint businessAddressCityNamedPropertyId = 0x8046;
            byte[] businessAddressCityNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressCityHex = Util.ConvertNamedPropertyToHex(businessAddressCityNamedPropertyId, businessAddressCityNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressCityHex) && namedPropertyIdTable[businessAddressCityHex] != null)
            {
                string businessAddressCityPropertyId = namedPropertyIdTable[businessAddressCityHex];

                businessAddressCityPropertyId = businessAddressCityPropertyId + stringTypeMask;

                Property businessAddressCityProperty = propertyTable.ContainsKey(businessAddressCityPropertyId) ? propertyTable[businessAddressCityPropertyId] : null;

                if (businessAddressCityProperty != null && businessAddressCityProperty.Value != null)
                {
                    businessAddressCity = encoding.GetString(businessAddressCityProperty.Value, 0, businessAddressCityProperty.Value.Length);
                }
            }

            uint businessAddressStateNamedPropertyId = 0x8047;
            byte[] businessAddressStateNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressStateHex = Util.ConvertNamedPropertyToHex(businessAddressStateNamedPropertyId, businessAddressStateNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressStateHex) && namedPropertyIdTable[businessAddressStateHex] != null)
            {
                string businessAddressStatePropertyId = namedPropertyIdTable[businessAddressStateHex];

                businessAddressStatePropertyId = businessAddressStatePropertyId + stringTypeMask;

                Property businessAddressStateProperty = propertyTable.ContainsKey(businessAddressStatePropertyId) ? propertyTable[businessAddressStatePropertyId] : null;

                if (businessAddressStateProperty != null && businessAddressStateProperty.Value != null)
                {
                    businessAddressState = encoding.GetString(businessAddressStateProperty.Value, 0, businessAddressStateProperty.Value.Length);
                }
            }

            uint businessAddressPostalCodeNamedPropertyId = 0x8048;
            byte[] businessAddressPostalCodeNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressPostalCodeHex = Util.ConvertNamedPropertyToHex(businessAddressPostalCodeNamedPropertyId, businessAddressPostalCodeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressPostalCodeHex) && namedPropertyIdTable[businessAddressPostalCodeHex] != null)
            {
                string businessAddressPostalCodePropertyId = namedPropertyIdTable[businessAddressPostalCodeHex];

                businessAddressPostalCodePropertyId = businessAddressPostalCodePropertyId + stringTypeMask;

                Property businessAddressPostalCodeProperty = propertyTable.ContainsKey(businessAddressPostalCodePropertyId) ? propertyTable[businessAddressPostalCodePropertyId] : null;

                if (businessAddressPostalCodeProperty != null && businessAddressPostalCodeProperty.Value != null)
                {
                    businessAddressPostalCode = encoding.GetString(businessAddressPostalCodeProperty.Value, 0, businessAddressPostalCodeProperty.Value.Length);
                }
            }

            uint businessAddressCountryNamedPropertyId = 0x8049;
            byte[] businessAddressCountryNamedPropertyGuid = StandardPropertySet.Address;

            string businessAddressCountryHex = Util.ConvertNamedPropertyToHex(businessAddressCountryNamedPropertyId, businessAddressCountryNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(businessAddressCountryHex) && namedPropertyIdTable[businessAddressCountryHex] != null)
            {
                string businessAddressCountryPropertyId = namedPropertyIdTable[businessAddressCountryHex];

                businessAddressCountryPropertyId = businessAddressCountryPropertyId + stringTypeMask;

                Property businessAddressCountryProperty = propertyTable.ContainsKey(businessAddressCountryPropertyId) ? propertyTable[businessAddressCountryPropertyId] : null;

                if (businessAddressCountryProperty != null && businessAddressCountryProperty.Value != null)
                {
                    businessAddressCountry = encoding.GetString(businessAddressCountryProperty.Value, 0, businessAddressCountryProperty.Value.Length);
                }
            }

            uint homeAddressNamedPropertyId = 0x801A;
            byte[] homeAddressNamedPropertyGuid = StandardPropertySet.Address;

            string homeAddressHex = Util.ConvertNamedPropertyToHex(homeAddressNamedPropertyId, homeAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(homeAddressHex) && namedPropertyIdTable[homeAddressHex] != null)
            {
                string homeAddressPropertyId = namedPropertyIdTable[homeAddressHex];

                homeAddressPropertyId = homeAddressPropertyId + stringTypeMask;

                Property homeAddressProperty = propertyTable.ContainsKey(homeAddressPropertyId) ? propertyTable[homeAddressPropertyId] : null;

                if (homeAddressProperty != null && homeAddressProperty.Value != null)
                {
                    homeAddress = encoding.GetString(homeAddressProperty.Value, 0, homeAddressProperty.Value.Length);
                }
            }

            uint otherAddressNamedPropertyId = 0x801C;
            byte[] otherAddressNamedPropertyGuid = StandardPropertySet.Address;

            string otherAddressHex = Util.ConvertNamedPropertyToHex(otherAddressNamedPropertyId, otherAddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(otherAddressHex) && namedPropertyIdTable[otherAddressHex] != null)
            {
                string otherAddressPropertyId = namedPropertyIdTable[otherAddressHex];

                otherAddressPropertyId = otherAddressPropertyId + stringTypeMask;

                Property otherAddressProperty = propertyTable.ContainsKey(otherAddressPropertyId) ? propertyTable[otherAddressPropertyId] : null;

                if (otherAddressProperty != null && otherAddressProperty.Value != null)
                {
                    otherAddress = encoding.GetString(otherAddressProperty.Value, 0, otherAddressProperty.Value.Length);
                }
            }

            uint email1AddressNamedPropertyId = 0x8083;
            byte[] email1AddressNamedPropertyGuid = StandardPropertySet.Address;

            string email1AddressHex = Util.ConvertNamedPropertyToHex(email1AddressNamedPropertyId, email1AddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email1AddressHex) && namedPropertyIdTable[email1AddressHex] != null)
            {
                string email1AddressPropertyId = namedPropertyIdTable[email1AddressHex];

                email1AddressPropertyId = email1AddressPropertyId + stringTypeMask;

                Property email1AddressProperty = propertyTable.ContainsKey(email1AddressPropertyId) ? propertyTable[email1AddressPropertyId] : null;

                if (email1AddressProperty != null && email1AddressProperty.Value != null)
                {
                    email1Address = encoding.GetString(email1AddressProperty.Value, 0, email1AddressProperty.Value.Length);
                }
            }

            uint email2AddressNamedPropertyId = 0x8093;
            byte[] email2AddressNamedPropertyGuid = StandardPropertySet.Address;

            string email2AddressHex = Util.ConvertNamedPropertyToHex(email2AddressNamedPropertyId, email2AddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email2AddressHex) && namedPropertyIdTable[email2AddressHex] != null)
            {
                string email2AddressPropertyId = namedPropertyIdTable[email2AddressHex];

                email2AddressPropertyId = email2AddressPropertyId + stringTypeMask;

                Property email2AddressProperty = propertyTable.ContainsKey(email2AddressPropertyId) ? propertyTable[email2AddressPropertyId] : null;

                if (email2AddressProperty != null && email2AddressProperty.Value != null)
                {
                    email2Address = encoding.GetString(email2AddressProperty.Value, 0, email2AddressProperty.Value.Length);
                }
            }

            uint email3AddressNamedPropertyId = 0x80A3;
            byte[] email3AddressNamedPropertyGuid = StandardPropertySet.Address;

            string email3AddressHex = Util.ConvertNamedPropertyToHex(email3AddressNamedPropertyId, email3AddressNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email3AddressHex) && namedPropertyIdTable[email3AddressHex] != null)
            {
                string email3AddressPropertyId = namedPropertyIdTable[email3AddressHex];

                email3AddressPropertyId = email3AddressPropertyId + stringTypeMask;

                Property email3AddressProperty = propertyTable.ContainsKey(email3AddressPropertyId) ? propertyTable[email3AddressPropertyId] : null;

                if (email3AddressProperty != null && email3AddressProperty.Value != null)
                {
                    email3Address = encoding.GetString(email3AddressProperty.Value, 0, email3AddressProperty.Value.Length);
                }
            }

            uint email1DisplayNameNamedPropertyId = 0x8084;
            byte[] email1DisplayNameNamedPropertyGuid = StandardPropertySet.Address;

            string email1DisplayNameHex = Util.ConvertNamedPropertyToHex(email1DisplayNameNamedPropertyId, email1DisplayNameNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email1DisplayNameHex) && namedPropertyIdTable[email1DisplayNameHex] != null)
            {
                string email1DisplayNamePropertyId = namedPropertyIdTable[email1DisplayNameHex];

                email1DisplayNamePropertyId = email1DisplayNamePropertyId + stringTypeMask;

                Property email1DisplayNameProperty = propertyTable.ContainsKey(email1DisplayNamePropertyId) ? propertyTable[email1DisplayNamePropertyId] : null;

                if (email1DisplayNameProperty != null && email1DisplayNameProperty.Value != null)
                {
                    email1DisplayName = encoding.GetString(email1DisplayNameProperty.Value, 0, email1DisplayNameProperty.Value.Length);
                }
            }

            uint email2DisplayNameNamedPropertyId = 0x8094;
            byte[] email2DisplayNameNamedPropertyGuid = StandardPropertySet.Address;

            string email2DisplayNameHex = Util.ConvertNamedPropertyToHex(email2DisplayNameNamedPropertyId, email2DisplayNameNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email2DisplayNameHex) && namedPropertyIdTable[email2DisplayNameHex] != null)
            {
                string email2DisplayNamePropertyId = namedPropertyIdTable[email2DisplayNameHex];

                email2DisplayNamePropertyId = email2DisplayNamePropertyId + stringTypeMask;

                Property email2DisplayNameProperty = propertyTable.ContainsKey(email2DisplayNamePropertyId) ? propertyTable[email2DisplayNamePropertyId] : null;

                if (email2DisplayNameProperty != null && email2DisplayNameProperty.Value != null)
                {
                    email2DisplayName = encoding.GetString(email2DisplayNameProperty.Value, 0, email2DisplayNameProperty.Value.Length);
                }
            }

            uint email3DisplayNameNamedPropertyId = 0x80A4;
            byte[] email3DisplayNameNamedPropertyGuid = StandardPropertySet.Address;

            string email3DisplayNameHex = Util.ConvertNamedPropertyToHex(email3DisplayNameNamedPropertyId, email3DisplayNameNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email3DisplayNameHex) && namedPropertyIdTable[email3DisplayNameHex] != null)
            {
                string email3DisplayNamePropertyId = namedPropertyIdTable[email3DisplayNameHex];

                email3DisplayNamePropertyId = email3DisplayNamePropertyId + stringTypeMask;

                Property email3DisplayNameProperty = propertyTable.ContainsKey(email3DisplayNamePropertyId) ? propertyTable[email3DisplayNamePropertyId] : null;

                if (email3DisplayNameProperty != null && email3DisplayNameProperty.Value != null)
                {
                    email3DisplayName = encoding.GetString(email3DisplayNameProperty.Value, 0, email3DisplayNameProperty.Value.Length);
                }
            }

            uint email1DisplayAsNamedPropertyId = 0x8080;
            byte[] email1DisplayAsNamedPropertyGuid = StandardPropertySet.Address;

            string email1DisplayAsHex = Util.ConvertNamedPropertyToHex(email1DisplayAsNamedPropertyId, email1DisplayAsNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email1DisplayAsHex) && namedPropertyIdTable[email1DisplayAsHex] != null)
            {
                string email1DisplayAsPropertyId = namedPropertyIdTable[email1DisplayAsHex];

                email1DisplayAsPropertyId = email1DisplayAsPropertyId + stringTypeMask;

                Property email1DisplayAsProperty = propertyTable.ContainsKey(email1DisplayAsPropertyId) ? propertyTable[email1DisplayAsPropertyId] : null;

                if (email1DisplayAsProperty != null && email1DisplayAsProperty.Value != null)
                {
                    email1DisplayAs = encoding.GetString(email1DisplayAsProperty.Value, 0, email1DisplayAsProperty.Value.Length);
                }
            }

            uint email2DisplayAsNamedPropertyId = 0x8090;
            byte[] email2DisplayAsNamedPropertyGuid = StandardPropertySet.Address;

            string email2DisplayAsHex = Util.ConvertNamedPropertyToHex(email2DisplayAsNamedPropertyId, email2DisplayAsNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email2DisplayAsHex) && namedPropertyIdTable[email2DisplayAsHex] != null)
            {
                string email2DisplayAsPropertyId = namedPropertyIdTable[email2DisplayAsHex];

                email2DisplayAsPropertyId = email2DisplayAsPropertyId + stringTypeMask;

                Property email2DisplayAsProperty = propertyTable.ContainsKey(email2DisplayAsPropertyId) ? propertyTable[email2DisplayAsPropertyId] : null;

                if (email2DisplayAsProperty != null && email2DisplayAsProperty.Value != null)
                {
                    email2DisplayAs = encoding.GetString(email2DisplayAsProperty.Value, 0, email2DisplayAsProperty.Value.Length);
                }
            }

            uint email3DisplayAsNamedPropertyId = 0x80A0;
            byte[] email3DisplayAsNamedPropertyGuid = StandardPropertySet.Address;

            string email3DisplayAsHex = Util.ConvertNamedPropertyToHex(email3DisplayAsNamedPropertyId, email3DisplayAsNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email3DisplayAsHex) && namedPropertyIdTable[email3DisplayAsHex] != null)
            {
                string email3DisplayAsPropertyId = namedPropertyIdTable[email3DisplayAsHex];

                email3DisplayAsPropertyId = email3DisplayAsPropertyId + stringTypeMask;

                Property email3DisplayAsProperty = propertyTable.ContainsKey(email3DisplayAsPropertyId) ? propertyTable[email3DisplayAsPropertyId] : null;

                if (email3DisplayAsProperty != null && email3DisplayAsProperty.Value != null)
                {
                    email3DisplayAs = encoding.GetString(email3DisplayAsProperty.Value, 0, email3DisplayAsProperty.Value.Length);
                }
            }

            uint email1TypeNamedPropertyId = 0x8082;
            byte[] email1TypeNamedPropertyGuid = StandardPropertySet.Address;

            string email1TypeHex = Util.ConvertNamedPropertyToHex(email1TypeNamedPropertyId, email1TypeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email1TypeHex) && namedPropertyIdTable[email1TypeHex] != null)
            {
                string email1TypePropertyId = namedPropertyIdTable[email1TypeHex];

                email1TypePropertyId = email1TypePropertyId + stringTypeMask;

                Property email1TypeProperty = propertyTable.ContainsKey(email1TypePropertyId) ? propertyTable[email1TypePropertyId] : null;

                if (email1TypeProperty != null && email1TypeProperty.Value != null)
                {
                    email1Type = encoding.GetString(email1TypeProperty.Value, 0, email1TypeProperty.Value.Length);
                }
            }

            uint email2TypeNamedPropertyId = 0x8092;
            byte[] email2TypeNamedPropertyGuid = StandardPropertySet.Address;

            string email2TypeHex = Util.ConvertNamedPropertyToHex(email2TypeNamedPropertyId, email2TypeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email2TypeHex) && namedPropertyIdTable[email2TypeHex] != null)
            {
                string email2TypePropertyId = namedPropertyIdTable[email2TypeHex];

                email2TypePropertyId = email2TypePropertyId + stringTypeMask;

                Property email2TypeProperty = propertyTable.ContainsKey(email2TypePropertyId) ? propertyTable[email2TypePropertyId] : null;

                if (email2TypeProperty != null && email2TypeProperty.Value != null)
                {
                    email2Type = encoding.GetString(email2TypeProperty.Value, 0, email2TypeProperty.Value.Length);
                }
            }

            uint email3TypeNamedPropertyId = 0x80A2;
            byte[] email3TypeNamedPropertyGuid = StandardPropertySet.Address;

            string email3TypeHex = Util.ConvertNamedPropertyToHex(email3TypeNamedPropertyId, email3TypeNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email3TypeHex) && namedPropertyIdTable[email3TypeHex] != null)
            {
                string email3TypePropertyId = namedPropertyIdTable[email3TypeHex];

                email3TypePropertyId = email3TypePropertyId + stringTypeMask;

                Property email3TypeProperty = propertyTable.ContainsKey(email3TypePropertyId) ? propertyTable[email3TypePropertyId] : null;

                if (email3TypeProperty != null && email3TypeProperty.Value != null)
                {
                    email3Type = encoding.GetString(email3TypeProperty.Value, 0, email3TypeProperty.Value.Length);
                }
            }

            uint email1EntryIdNamedPropertyId = 0x8085;
            byte[] email1EntryIdNamedPropertyGuid = StandardPropertySet.Address;

            string email1EntryIdHex = Util.ConvertNamedPropertyToHex(email1EntryIdNamedPropertyId, email1EntryIdNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email1EntryIdHex) && namedPropertyIdTable[email1EntryIdHex] != null)
            {
                string email1EntryIdPropertyId = namedPropertyIdTable[email1EntryIdHex];

                email1EntryIdPropertyId = email1EntryIdPropertyId + "0102";

                Property email1EntryIdProperty = propertyTable.ContainsKey(email1EntryIdPropertyId) ? propertyTable[email1EntryIdPropertyId] : null;

                if (email1EntryIdProperty != null && email1EntryIdProperty.Value != null)
                {
                    email1EntryId = email1EntryIdProperty.Value;
                }
            }

            uint email2EntryIdNamedPropertyId = 0x8095;
            byte[] email2EntryIdNamedPropertyGuid = StandardPropertySet.Address;

            string email2EntryIdHex = Util.ConvertNamedPropertyToHex(email2EntryIdNamedPropertyId, email2EntryIdNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email2EntryIdHex) && namedPropertyIdTable[email2EntryIdHex] != null)
            {
                string email2EntryIdPropertyId = namedPropertyIdTable[email2EntryIdHex];

                email2EntryIdPropertyId = email2EntryIdPropertyId + "0102";

                Property email2EntryIdProperty = propertyTable.ContainsKey(email2EntryIdPropertyId) ? propertyTable[email2EntryIdPropertyId] : null;

                if (email2EntryIdProperty != null && email2EntryIdProperty.Value != null)
                {
                    email2EntryId = email2EntryIdProperty.Value;
                }
            }

            uint email3EntryIdNamedPropertyId = 0x80A5;
            byte[] email3EntryIdNamedPropertyGuid = StandardPropertySet.Address;

            string email3EntryIdHex = Util.ConvertNamedPropertyToHex(email3EntryIdNamedPropertyId, email3EntryIdNamedPropertyGuid);

            if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(email3EntryIdHex) && namedPropertyIdTable[email3EntryIdHex] != null)
            {
                string email3EntryIdPropertyId = namedPropertyIdTable[email3EntryIdHex];

                email3EntryIdPropertyId = email3EntryIdPropertyId + "0102";

                Property email3EntryIdProperty = propertyTable.ContainsKey(email3EntryIdPropertyId) ? propertyTable[email3EntryIdPropertyId] : null;

                if (email3EntryIdProperty != null && email3EntryIdProperty.Value != null)
                {
                    email3EntryId = email3EntryIdProperty.Value;
                }
            }

            //set extendedProperties value and type
            for (int e = 0; e < extendedProperties.Count; e++)
            {
                string extendedPropertyHex = null;

                if (extendedProperties[e].Tag is ExtendedPropertyId)
                {
                    ExtendedPropertyId tag = (ExtendedPropertyId)extendedProperties[e].Tag;

                    extendedPropertyHex = Util.ConvertNamedPropertyToHex((uint)tag.Id, tag.Guid);
                }
                else
                {
                    ExtendedPropertyName tag = (ExtendedPropertyName)extendedProperties[e].Tag;

                    extendedPropertyHex = Util.ConvertNamedPropertyToHex(tag.Name, tag.Guid);
                }

                if (namedPropertyIdTable != null && namedPropertyIdTable.ContainsKey(extendedPropertyHex) && namedPropertyIdTable[extendedPropertyHex] != null)
                {
                    string extendedPropertyId = namedPropertyIdTable[extendedPropertyHex];

                    foreach (string key in propertyTable.Keys)
                    {
                        if (key.StartsWith(extendedPropertyId))
                        {
                            Property property = propertyTable[key];

                            extendedProperties[e].Tag.Type = property.Type;

                            if (extendedProperties[e].Tag.Type == PropertyType.MultipleBinary)
                            {
                                extendedPropertyId = extendedPropertyId + "1102";

                                if (property != null && property.Value != null)
                                {
                                    uint propertyLenghtValue = property.Size / 8;

                                    IList<byte[]> valueList = new List<byte[]>();

                                    for (int i = 0; i < propertyLenghtValue; i++)
                                    {
                                        string valueStreamName = "__substg1.0_" + extendedPropertyId + "-" + String.Format("{0:X8}", i);

                                        Independentsoft.IO.StructuredStorage.Stream valueStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries[valueStreamName];

                                        if (valueStream != null && valueStream.Buffer != null)
                                        {
                                            valueList.Add(valueStream.Buffer);
                                        }
                                    }

                                    if (valueList.Count > 0)
                                    {
                                        MemoryStream memoryStream = new MemoryStream();

                                        byte[] countBuffer = BitConverter.GetBytes(valueList.Count);
                                        memoryStream.Write(countBuffer, 0, 4);

                                        int offset = 0;

                                        for (int i = 0; i < valueList.Count; i++)
                                        {
                                            byte[] buffer = valueList[i];
                                            byte[] startBuffer = BitConverter.GetBytes(4 + valueList.Count * 4 + offset);
                                            memoryStream.Write(startBuffer, 0, startBuffer.Length);

                                            offset += buffer.Length;
                                        }

                                        for (int i = 0; i < valueList.Count; i++)
                                        {
                                            byte[] buffer = valueList[i];
                                            memoryStream.Write(buffer, 0, buffer.Length);
                                        }

                                        extendedProperties[e].Value = memoryStream.ToArray();
                                    }
                                }
                            }
                            else
                            {
                                extendedProperties[e].Value = property.Value;
                            }
                        }
                    }
                }
            }

            Independentsoft.IO.StructuredStorage.Stream messageClassStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_001A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream subjectStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0037" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream subjectPrefixStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_003D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream conversationTopicStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0070" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream displayBccStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0E02" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream displayCcStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0E03" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream displayToStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0E04" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream originalDisplayToStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0074" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream replyToStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0050" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream normalizedSubjectStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0E1D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream bodyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1000" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream rtfCompressedStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_10090102"];
            Independentsoft.IO.StructuredStorage.Stream searchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_300B0102"];
            Independentsoft.IO.StructuredStorage.Stream changeKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_65E20102"];
            Independentsoft.IO.StructuredStorage.Stream entryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0FFF0102"];
            Independentsoft.IO.StructuredStorage.Stream readReceiptEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00460102"];
            Independentsoft.IO.StructuredStorage.Stream readReceiptSearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00530102"];
            Independentsoft.IO.StructuredStorage.Stream reportTextStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1001" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream creatorNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3FF8" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream lastModifierNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3FFA" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream internetMessageIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1035" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream inReplyToStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1042" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream internetReferencesStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1039" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream conversationIndexStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00710102"];
            Independentsoft.IO.StructuredStorage.Stream bodyHtmlStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_10130102"];
            Independentsoft.IO.StructuredStorage.Stream bodyHtmlTextStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_1013" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedRepresentingAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0077" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedRepresentingEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0078" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedRepresentingEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00430102"];
            Independentsoft.IO.StructuredStorage.Stream receivedRepresentingNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0044" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedRepresentingSearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00520102"];
            Independentsoft.IO.StructuredStorage.Stream receivedByAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0075" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedByEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0076" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedByEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_003F0102"];
            Independentsoft.IO.StructuredStorage.Stream receivedByNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0040" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream receivedBySearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00510102"];
            Independentsoft.IO.StructuredStorage.Stream senderAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0C1E" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream senderEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0C1F" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream senderSmtpAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_5D01" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream senderEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0C190102"];
            Independentsoft.IO.StructuredStorage.Stream senderNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0C1A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream senderSearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0C1D0102"];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0064" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0065" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingSmtpAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_5D02" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_00410102"];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_0042" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream sentRepresentingSearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_003B0102"];
            Independentsoft.IO.StructuredStorage.Stream transportMessageHeadersStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_007D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream assistentNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A30" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream assistentPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2E" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessPhone2Stream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1B" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessFaxStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A24" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessHomePageStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A51" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream callbackPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A02" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream carPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1E" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream cellularPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1C" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream companyMainPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A57" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream companyNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A16" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream computerNetworkNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A49" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream customerIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A4A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream departmentNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A18" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream displayNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3001" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream displayNamePrefixStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A45" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream ftpSiteStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A4C" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream generationStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A05" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream givenNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A06" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream governmentIdStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A07" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream hobbiesStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A43" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homePhone2Stream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2F" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressCityStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A59" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressCountryStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressPostalCodeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5B" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressPostOfficeBoxStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5E" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressStateStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5C" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeAddressStreetStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homeFaxStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A25" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream homePhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A09" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream initialsStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A0A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream isdnStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream managerNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A4E" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream middleNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A44" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream nicknameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A4F" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream officeLocationStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A19" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A08" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherAddressCityStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A5F" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherAddressCountryStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A60" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherAddressPostalCodeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A61" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherAddressStateStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A62" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherAddressStreetStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A63" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream otherPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1F" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream pagerStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A21" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream personalHomePageStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A50" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream postalAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A15" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressCityStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A27" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressCountryStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A26" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressPostalCodeStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressPostOfficeBoxStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2B" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressStateStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A28" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream businessAddressStreetStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A29" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream primaryFaxStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A23" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream primaryPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1A" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream professionStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A46" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream radioPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A1D" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream spouseNameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A48" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream surnameStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A11" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream telexStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A2C" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream titleStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A17" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream ttyTddPhoneStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_3A4B" + stringTypeMask];
            Independentsoft.IO.StructuredStorage.Stream extendedSenderEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)directoryEntries["__substg1.0_5D0A" + unicodeTypeMask];

            if (messageClassStream != null && messageClassStream.Buffer != null)
            {
                messageClass = encoding.GetString(messageClassStream.Buffer, 0, messageClassStream.Buffer.Length);
            }

            if (subjectStream != null && subjectStream.Buffer != null)
            {
                subject = encoding.GetString(subjectStream.Buffer, 0, subjectStream.Buffer.Length);
            }

            if (subjectPrefixStream != null && subjectPrefixStream.Buffer != null)
            {
                subjectPrefix = encoding.GetString(subjectPrefixStream.Buffer, 0, subjectPrefixStream.Buffer.Length);
            }

            if (conversationTopicStream != null && conversationTopicStream.Buffer != null)
            {
                conversationTopic = encoding.GetString(conversationTopicStream.Buffer, 0, conversationTopicStream.Buffer.Length);
            }

            if (displayBccStream != null && displayBccStream.Buffer != null)
            {
                displayBcc = encoding.GetString(displayBccStream.Buffer, 0, displayBccStream.Buffer.Length);
            }

            if (displayCcStream != null && displayCcStream.Buffer != null)
            {
                displayCc = encoding.GetString(displayCcStream.Buffer, 0, displayCcStream.Buffer.Length);
            }

            if (displayToStream != null && displayToStream.Buffer != null)
            {
                displayTo = encoding.GetString(displayToStream.Buffer, 0, displayToStream.Buffer.Length);
            }

            if (originalDisplayToStream != null && originalDisplayToStream.Buffer != null)
            {
                originalDisplayTo = encoding.GetString(originalDisplayToStream.Buffer, 0, originalDisplayToStream.Buffer.Length);
            }

            if (replyToStream != null && replyToStream.Buffer != null)
            {
                replyTo = encoding.GetString(replyToStream.Buffer, 0, replyToStream.Buffer.Length);
            }

            if (normalizedSubjectStream != null && normalizedSubjectStream.Buffer != null)
            {
                normalizedSubject = encoding.GetString(normalizedSubjectStream.Buffer, 0, normalizedSubjectStream.Buffer.Length);
            }

            if (bodyStream != null && bodyStream.Buffer != null)
            {
                body = encoding.GetString(bodyStream.Buffer, 0, bodyStream.Buffer.Length);
            }

            if (rtfCompressedStream != null && rtfCompressedStream.Buffer != null)
            {
                rtfCompressed = rtfCompressedStream.Buffer;
            }

            if (searchKeyStream != null && searchKeyStream.Buffer != null)
            {
                searchKey = searchKeyStream.Buffer;
            }

            if (changeKeyStream != null && changeKeyStream.Buffer != null)
            {
                changeKey = changeKeyStream.Buffer;
            }

            if (entryIdStream != null && entryIdStream.Buffer != null)
            {
                entryId = entryIdStream.Buffer;
            }

            if (readReceiptEntryIdStream != null && readReceiptEntryIdStream.Buffer != null)
            {
                readReceiptEntryId = readReceiptEntryIdStream.Buffer;
            }

            if (readReceiptSearchKeyStream != null && readReceiptSearchKeyStream.Buffer != null)
            {
                readReceiptSearchKey = readReceiptSearchKeyStream.Buffer;
            }

            if (reportTextStream != null && reportTextStream.Buffer != null)
            {
                reportText = encoding.GetString(reportTextStream.Buffer, 0, reportTextStream.Buffer.Length);
            }

            if (creatorNameStream != null && creatorNameStream.Buffer != null)
            {
                creatorName = encoding.GetString(creatorNameStream.Buffer, 0, creatorNameStream.Buffer.Length);
            }

            if (lastModifierNameStream != null && lastModifierNameStream.Buffer != null)
            {
                lastModifierName = encoding.GetString(lastModifierNameStream.Buffer, 0, lastModifierNameStream.Buffer.Length);
            }

            if (internetMessageIdStream != null && internetMessageIdStream.Buffer != null)
            {
                internetMessageId = encoding.GetString(internetMessageIdStream.Buffer, 0, internetMessageIdStream.Buffer.Length);
            }

            if (inReplyToStream != null && inReplyToStream.Buffer != null)
            {
                inReplyTo = encoding.GetString(inReplyToStream.Buffer, 0, inReplyToStream.Buffer.Length);
            }

            if (internetReferencesStream != null && internetReferencesStream.Buffer != null)
            {
                internetReferences = encoding.GetString(internetReferencesStream.Buffer, 0, internetReferencesStream.Buffer.Length);
            }

            if (conversationIndexStream != null && conversationIndexStream.Buffer != null)
            {
                conversationIndex = conversationIndexStream.Buffer;
            }

            if (bodyHtmlStream != null && bodyHtmlStream.Buffer != null)
            {
                bodyHtml = bodyHtmlStream.Buffer;
            }
            else if (bodyHtmlTextStream != null && bodyHtmlTextStream.Buffer != null)
            {
                System.Text.Encoding internetCodePageEncoding = null;

                try
                {
                    internetCodePageEncoding = System.Text.Encoding.GetEncoding((int)this.internetCodePage);
                }
                catch //ignore error like non-existing code page
                {
                }

                if (this.internetCodePage > 0 && internetCodePageEncoding != null)
                {
                    string bodyHtmlText = internetCodePageEncoding.GetString(bodyHtmlTextStream.Buffer, 0, bodyHtmlTextStream.Buffer.Length);

                    bodyHtml = internetCodePageEncoding.GetBytes(bodyHtmlText);
                }
                else if (Utf8Util.IsUtf8(bodyHtmlTextStream.Buffer, bodyHtmlTextStream.Buffer.Length))
                {
                    System.Text.Encoding htmlEncoding = System.Text.Encoding.UTF8;

                    string bodyHtmlText = htmlEncoding.GetString(bodyHtmlTextStream.Buffer, 0, bodyHtmlTextStream.Buffer.Length);

                    bodyHtml = htmlEncoding.GetBytes(bodyHtmlText);
                }
                else
                {
                    string bodyHtmlText = encoding.GetString(bodyHtmlTextStream.Buffer, 0, bodyHtmlTextStream.Buffer.Length);

                    System.Text.Encoding htmlEncoding = System.Text.Encoding.UTF7;

                    int htmlCharsetIndex = bodyHtmlText.IndexOf("charset=");

                    if (htmlCharsetIndex > 0)
                    {
                        int htmlCharsetEndIndex = bodyHtmlText.IndexOf("\"", htmlCharsetIndex);

                        if (htmlCharsetEndIndex != -1)
                        {
                            string htmlCharset = bodyHtmlText.Substring(htmlCharsetIndex + 8, htmlCharsetEndIndex - htmlCharsetIndex - 8);
                            htmlEncoding = System.Text.Encoding.GetEncoding(htmlCharset);
                        }
                    }

                    bodyHtmlText = htmlEncoding.GetString(bodyHtmlTextStream.Buffer, 0, bodyHtmlTextStream.Buffer.Length);

                    bodyHtml = htmlEncoding.GetBytes(bodyHtmlText);
                }
            }

            if (receivedRepresentingAddressTypeStream != null && receivedRepresentingAddressTypeStream.Buffer != null)
            {
                receivedRepresentingAddressType = encoding.GetString(receivedRepresentingAddressTypeStream.Buffer, 0, receivedRepresentingAddressTypeStream.Buffer.Length);
            }

            if (receivedRepresentingEmailAddressStream != null && receivedRepresentingEmailAddressStream.Buffer != null)
            {
                receivedRepresentingEmailAddress = encoding.GetString(receivedRepresentingEmailAddressStream.Buffer, 0, receivedRepresentingEmailAddressStream.Buffer.Length);
            }

            if (receivedRepresentingEntryIdStream != null && receivedRepresentingEntryIdStream.Buffer != null)
            {
                receivedRepresentingEntryId = receivedRepresentingEntryIdStream.Buffer;
            }

            if (receivedRepresentingNameStream != null && receivedRepresentingNameStream.Buffer != null)
            {
                receivedRepresentingName = encoding.GetString(receivedRepresentingNameStream.Buffer, 0, receivedRepresentingNameStream.Buffer.Length);
            }

            if (receivedRepresentingSearchKeyStream != null && receivedRepresentingSearchKeyStream.Buffer != null)
            {
                receivedRepresentingSearchKey = receivedRepresentingSearchKeyStream.Buffer;
            }

            if (receivedByAddressTypeStream != null && receivedByAddressTypeStream.Buffer != null)
            {
                receivedByAddressType = encoding.GetString(receivedByAddressTypeStream.Buffer, 0, receivedByAddressTypeStream.Buffer.Length);
            }

            if (receivedByEmailAddressStream != null && receivedByEmailAddressStream.Buffer != null)
            {
                receivedByEmailAddress = encoding.GetString(receivedByEmailAddressStream.Buffer, 0, receivedByEmailAddressStream.Buffer.Length);
            }

            if (receivedByEntryIdStream != null && receivedByEntryIdStream.Buffer != null)
            {
                receivedByEntryId = receivedByEntryIdStream.Buffer;
            }

            if (receivedByNameStream != null && receivedByNameStream.Buffer != null)
            {
                receivedByName = encoding.GetString(receivedByNameStream.Buffer, 0, receivedByNameStream.Buffer.Length);
            }

            if (receivedBySearchKeyStream != null && receivedBySearchKeyStream.Buffer != null)
            {
                receivedBySearchKey = receivedBySearchKeyStream.Buffer;
            }

            if (senderAddressTypeStream != null && senderAddressTypeStream.Buffer != null)
            {
                senderAddressType = encoding.GetString(senderAddressTypeStream.Buffer, 0, senderAddressTypeStream.Buffer.Length);
            }

            if (extendedSenderEmailAddressStream != null && extendedSenderEmailAddressStream.Buffer != null)
            {
                senderEmailAddress = encodingUnicode.GetString(extendedSenderEmailAddressStream.Buffer, 0, extendedSenderEmailAddressStream.Buffer.Length);
            }

            if (string.IsNullOrWhiteSpace(senderEmailAddress) && senderEmailAddressStream != null && senderEmailAddressStream.Buffer != null)
            {
                senderEmailAddress = encoding.GetString(senderEmailAddressStream.Buffer, 0, senderEmailAddressStream.Buffer.Length);
            }

            if (senderSmtpAddressStream != null && senderSmtpAddressStream.Buffer != null)
            {
                senderSmtpAddress = encoding.GetString(senderSmtpAddressStream.Buffer, 0, senderSmtpAddressStream.Buffer.Length);
            }

            if (senderEntryIdStream != null && senderEntryIdStream.Buffer != null)
            {
                senderEntryId = senderEntryIdStream.Buffer;
            }

            if (senderNameStream != null && senderNameStream.Buffer != null)
            {
                senderName = encoding.GetString(senderNameStream.Buffer, 0, senderNameStream.Buffer.Length);
            }

            if (senderSearchKeyStream != null && senderSearchKeyStream.Buffer != null)
            {
                senderSearchKey = senderSearchKeyStream.Buffer;
            }

            if (sentRepresentingAddressTypeStream != null && sentRepresentingAddressTypeStream.Buffer != null)
            {
                sentRepresentingAddressType = encoding.GetString(sentRepresentingAddressTypeStream.Buffer, 0, sentRepresentingAddressTypeStream.Buffer.Length);
            }

            if (sentRepresentingEmailAddressStream != null && sentRepresentingEmailAddressStream.Buffer != null)
            {
                sentRepresentingEmailAddress = encoding.GetString(sentRepresentingEmailAddressStream.Buffer, 0, sentRepresentingEmailAddressStream.Buffer.Length);
            }

            if (sentRepresentingSmtpAddressStream != null && sentRepresentingSmtpAddressStream.Buffer != null)
            {
                sentRepresentingSmtpAddress = encoding.GetString(sentRepresentingSmtpAddressStream.Buffer, 0, sentRepresentingSmtpAddressStream.Buffer.Length);
            }

            if (sentRepresentingEntryIdStream != null && sentRepresentingEntryIdStream.Buffer != null)
            {
                sentRepresentingEntryId = sentRepresentingEntryIdStream.Buffer;
            }

            if (sentRepresentingNameStream != null && sentRepresentingNameStream.Buffer != null)
            {
                sentRepresentingName = encoding.GetString(sentRepresentingNameStream.Buffer, 0, sentRepresentingNameStream.Buffer.Length);
            }

            if (sentRepresentingSearchKeyStream != null && sentRepresentingSearchKeyStream.Buffer != null)
            {
                sentRepresentingSearchKey = sentRepresentingSearchKeyStream.Buffer;
            }

            if (transportMessageHeadersStream != null && transportMessageHeadersStream.Buffer != null)
            {
                transportMessageHeaders = encoding.GetString(transportMessageHeadersStream.Buffer, 0, transportMessageHeadersStream.Buffer.Length);
            }

            if (assistentNameStream != null && assistentNameStream.Buffer != null)
            {
                assistentName = encoding.GetString(assistentNameStream.Buffer, 0, assistentNameStream.Buffer.Length);
            }

            if (assistentPhoneStream != null && assistentPhoneStream.Buffer != null)
            {
                assistentPhone = encoding.GetString(assistentPhoneStream.Buffer, 0, assistentPhoneStream.Buffer.Length);
            }

            if (businessPhone2Stream != null && businessPhone2Stream.Buffer != null)
            {
                businessPhone2 = encoding.GetString(businessPhone2Stream.Buffer, 0, businessPhone2Stream.Buffer.Length);
            }

            if (businessFaxStream != null && businessFaxStream.Buffer != null)
            {
                businessFax = encoding.GetString(businessFaxStream.Buffer, 0, businessFaxStream.Buffer.Length);
            }

            if (businessHomePageStream != null && businessHomePageStream.Buffer != null)
            {
                businessHomePage = encoding.GetString(businessHomePageStream.Buffer, 0, businessHomePageStream.Buffer.Length);
            }

            if (callbackPhoneStream != null && callbackPhoneStream.Buffer != null)
            {
                callbackPhone = encoding.GetString(callbackPhoneStream.Buffer, 0, callbackPhoneStream.Buffer.Length);
            }

            if (carPhoneStream != null && carPhoneStream.Buffer != null)
            {
                carPhone = encoding.GetString(carPhoneStream.Buffer, 0, carPhoneStream.Buffer.Length);
            }

            if (cellularPhoneStream != null && cellularPhoneStream.Buffer != null)
            {
                cellularPhone = encoding.GetString(cellularPhoneStream.Buffer, 0, cellularPhoneStream.Buffer.Length);
            }

            if (companyMainPhoneStream != null && companyMainPhoneStream.Buffer != null)
            {
                companyMainPhone = encoding.GetString(companyMainPhoneStream.Buffer, 0, companyMainPhoneStream.Buffer.Length);
            }

            if (companyNameStream != null && companyNameStream.Buffer != null)
            {
                companyName = encoding.GetString(companyNameStream.Buffer, 0, companyNameStream.Buffer.Length);
            }

            if (computerNetworkNameStream != null && computerNetworkNameStream.Buffer != null)
            {
                computerNetworkName = encoding.GetString(computerNetworkNameStream.Buffer, 0, computerNetworkNameStream.Buffer.Length);
            }

            if (businessAddressCountryStream != null && businessAddressCountryStream.Buffer != null)
            {
                businessAddressCountry = encoding.GetString(businessAddressCountryStream.Buffer, 0, businessAddressCountryStream.Buffer.Length);
            }

            if (customerIdStream != null && customerIdStream.Buffer != null)
            {
                customerId = encoding.GetString(customerIdStream.Buffer, 0, customerIdStream.Buffer.Length);
            }

            if (departmentNameStream != null && departmentNameStream.Buffer != null)
            {
                departmentName = encoding.GetString(departmentNameStream.Buffer, 0, departmentNameStream.Buffer.Length);
            }

            if (displayNameStream != null && displayNameStream.Buffer != null)
            {
                displayName = encoding.GetString(displayNameStream.Buffer, 0, displayNameStream.Buffer.Length);
            }

            if (displayNamePrefixStream != null && displayNamePrefixStream.Buffer != null)
            {
                displayNamePrefix = encoding.GetString(displayNamePrefixStream.Buffer, 0, displayNamePrefixStream.Buffer.Length);
            }

            if (ftpSiteStream != null && ftpSiteStream.Buffer != null)
            {
                ftpSite = encoding.GetString(ftpSiteStream.Buffer, 0, ftpSiteStream.Buffer.Length);
            }

            if (generationStream != null && generationStream.Buffer != null)
            {
                generation = encoding.GetString(generationStream.Buffer, 0, generationStream.Buffer.Length);
            }

            if (givenNameStream != null && givenNameStream.Buffer != null)
            {
                givenName = encoding.GetString(givenNameStream.Buffer, 0, givenNameStream.Buffer.Length);
            }

            if (governmentIdStream != null && governmentIdStream.Buffer != null)
            {
                governmentId = encoding.GetString(governmentIdStream.Buffer, 0, governmentIdStream.Buffer.Length);
            }

            if (hobbiesStream != null && hobbiesStream.Buffer != null)
            {
                hobbies = encoding.GetString(hobbiesStream.Buffer, 0, hobbiesStream.Buffer.Length);
            }

            if (homePhone2Stream != null && homePhone2Stream.Buffer != null)
            {
                homePhone2 = encoding.GetString(homePhone2Stream.Buffer, 0, homePhone2Stream.Buffer.Length);
            }

            if (homeAddressCityStream != null && homeAddressCityStream.Buffer != null)
            {
                homeAddressCity = encoding.GetString(homeAddressCityStream.Buffer, 0, homeAddressCityStream.Buffer.Length);
            }

            if (homeAddressCountryStream != null && homeAddressCountryStream.Buffer != null)
            {
                homeAddressCountry = encoding.GetString(homeAddressCountryStream.Buffer, 0, homeAddressCountryStream.Buffer.Length);
            }

            if (homeAddressPostalCodeStream != null && homeAddressPostalCodeStream.Buffer != null)
            {
                homeAddressPostalCode = encoding.GetString(homeAddressPostalCodeStream.Buffer, 0, homeAddressPostalCodeStream.Buffer.Length);
            }

            if (homeAddressPostOfficeBoxStream != null && homeAddressPostOfficeBoxStream.Buffer != null)
            {
                homeAddressPostOfficeBox = encoding.GetString(homeAddressPostOfficeBoxStream.Buffer, 0, homeAddressPostOfficeBoxStream.Buffer.Length);
            }

            if (homeAddressStateStream != null && homeAddressStateStream.Buffer != null)
            {
                homeAddressState = encoding.GetString(homeAddressStateStream.Buffer, 0, homeAddressStateStream.Buffer.Length);
            }

            if (homeAddressStreetStream != null && homeAddressStreetStream.Buffer != null)
            {
                homeAddressStreet = encoding.GetString(homeAddressStreetStream.Buffer, 0, homeAddressStreetStream.Buffer.Length);
            }

            if (homeFaxStream != null && homeFaxStream.Buffer != null)
            {
                homeFax = encoding.GetString(homeFaxStream.Buffer, 0, homeFaxStream.Buffer.Length);
            }

            if (homePhoneStream != null && homePhoneStream.Buffer != null)
            {
                homePhone = encoding.GetString(homePhoneStream.Buffer, 0, homePhoneStream.Buffer.Length);
            }

            if (initialsStream != null && initialsStream.Buffer != null)
            {
                initials = encoding.GetString(initialsStream.Buffer, 0, initialsStream.Buffer.Length);
            }

            if (isdnStream != null && isdnStream.Buffer != null)
            {
                isdn = encoding.GetString(isdnStream.Buffer, 0, isdnStream.Buffer.Length);
            }

            if (businessAddressCityStream != null && businessAddressCityStream.Buffer != null)
            {
                businessAddressCity = encoding.GetString(businessAddressCityStream.Buffer, 0, businessAddressCityStream.Buffer.Length);
            }

            if (managerNameStream != null && managerNameStream.Buffer != null)
            {
                managerName = encoding.GetString(managerNameStream.Buffer, 0, managerNameStream.Buffer.Length);
            }

            if (middleNameStream != null && middleNameStream.Buffer != null)
            {
                middleName = encoding.GetString(middleNameStream.Buffer, 0, middleNameStream.Buffer.Length);
            }

            if (nicknameStream != null && nicknameStream.Buffer != null)
            {
                nickname = encoding.GetString(nicknameStream.Buffer, 0, nicknameStream.Buffer.Length);
            }

            if (officeLocationStream != null && officeLocationStream.Buffer != null)
            {
                officeLocation = encoding.GetString(officeLocationStream.Buffer, 0, officeLocationStream.Buffer.Length);
            }

            if (businessPhoneStream != null && businessPhoneStream.Buffer != null)
            {
                businessPhone = encoding.GetString(businessPhoneStream.Buffer, 0, businessPhoneStream.Buffer.Length);
            }

            if (otherAddressCityStream != null && otherAddressCityStream.Buffer != null)
            {
                otherAddressCity = encoding.GetString(otherAddressCityStream.Buffer, 0, otherAddressCityStream.Buffer.Length);
            }

            if (otherAddressCountryStream != null && otherAddressCountryStream.Buffer != null)
            {
                otherAddressCountry = encoding.GetString(otherAddressCountryStream.Buffer, 0, otherAddressCountryStream.Buffer.Length);
            }

            if (otherAddressPostalCodeStream != null && otherAddressPostalCodeStream.Buffer != null)
            {
                otherAddressPostalCode = encoding.GetString(otherAddressPostalCodeStream.Buffer, 0, otherAddressPostalCodeStream.Buffer.Length);
            }

            if (otherAddressStateStream != null && otherAddressStateStream.Buffer != null)
            {
                otherAddressState = encoding.GetString(otherAddressStateStream.Buffer, 0, otherAddressStateStream.Buffer.Length);
            }

            if (otherAddressStreetStream != null && otherAddressStreetStream.Buffer != null)
            {
                otherAddressStreet = encoding.GetString(otherAddressStreetStream.Buffer, 0, otherAddressStreetStream.Buffer.Length);
            }

            if (otherPhoneStream != null && otherPhoneStream.Buffer != null)
            {
                otherPhone = encoding.GetString(otherPhoneStream.Buffer, 0, otherPhoneStream.Buffer.Length);
            }

            if (pagerStream != null && pagerStream.Buffer != null)
            {
                pager = encoding.GetString(pagerStream.Buffer, 0, pagerStream.Buffer.Length);
            }

            if (personalHomePageStream != null && personalHomePageStream.Buffer != null)
            {
                personalHomePage = encoding.GetString(personalHomePageStream.Buffer, 0, personalHomePageStream.Buffer.Length);
            }

            if (postalAddressStream != null && postalAddressStream.Buffer != null)
            {
                postalAddress = encoding.GetString(postalAddressStream.Buffer, 0, postalAddressStream.Buffer.Length);
            }

            if (businessAddressPostalCodeStream != null && businessAddressPostalCodeStream.Buffer != null)
            {
                businessAddressPostalCode = encoding.GetString(businessAddressPostalCodeStream.Buffer, 0, businessAddressPostalCodeStream.Buffer.Length);
            }

            if (businessAddressPostOfficeBoxStream != null && businessAddressPostOfficeBoxStream.Buffer != null)
            {
                businessAddressPostOfficeBox = encoding.GetString(businessAddressPostOfficeBoxStream.Buffer, 0, businessAddressPostOfficeBoxStream.Buffer.Length);
            }

            if (businessAddressStateStream != null && businessAddressStateStream.Buffer != null)
            {
                businessAddressState = encoding.GetString(businessAddressStateStream.Buffer, 0, businessAddressStateStream.Buffer.Length);
            }

            if (businessAddressStreetStream != null && businessAddressStreetStream.Buffer != null)
            {
                businessAddressStreet = encoding.GetString(businessAddressStreetStream.Buffer, 0, businessAddressStreetStream.Buffer.Length);
            }

            if (primaryFaxStream != null && primaryFaxStream.Buffer != null)
            {
                primaryFax = encoding.GetString(primaryFaxStream.Buffer, 0, primaryFaxStream.Buffer.Length);
            }

            if (primaryPhoneStream != null && primaryPhoneStream.Buffer != null)
            {
                primaryPhone = encoding.GetString(primaryPhoneStream.Buffer, 0, primaryPhoneStream.Buffer.Length);
            }

            if (professionStream != null && professionStream.Buffer != null)
            {
                profession = encoding.GetString(professionStream.Buffer, 0, professionStream.Buffer.Length);
            }

            if (radioPhoneStream != null && radioPhoneStream.Buffer != null)
            {
                radioPhone = encoding.GetString(radioPhoneStream.Buffer, 0, radioPhoneStream.Buffer.Length);
            }

            if (spouseNameStream != null && spouseNameStream.Buffer != null)
            {
                spouseName = encoding.GetString(spouseNameStream.Buffer, 0, spouseNameStream.Buffer.Length);
            }

            if (surnameStream != null && surnameStream.Buffer != null)
            {
                surname = encoding.GetString(surnameStream.Buffer, 0, surnameStream.Buffer.Length);
            }

            if (telexStream != null && telexStream.Buffer != null)
            {
                telex = encoding.GetString(telexStream.Buffer, 0, telexStream.Buffer.Length);
            }

            if (titleStream != null && titleStream.Buffer != null)
            {
                title = encoding.GetString(titleStream.Buffer, 0, titleStream.Buffer.Length);
            }

            if (ttyTddPhoneStream != null && ttyTddPhoneStream.Buffer != null)
            {
                ttyTddPhone = encoding.GetString(ttyTddPhoneStream.Buffer, 0, ttyTddPhoneStream.Buffer.Length);
            }

            //Recipients
            for (int i = 0; i < recipientCount; i++)
            {
                IDictionary<String, Property> recipientPropertyTable = new Dictionary<String, Property>();
                string recipientStorageName = String.Format("__recip_version1.0_#{0:X8}", i);

                Independentsoft.IO.StructuredStorage.Storage recipientStorage = (Independentsoft.IO.StructuredStorage.Storage)directoryEntries[recipientStorageName];

                if (recipientStorage == null)
                {
                    continue;
                }

                Independentsoft.IO.StructuredStorage.Stream recipientPropertiesStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__properties_version1.0"];

                if (recipientPropertiesStream != null && recipientPropertiesStream.Buffer != null)
                {
                    //first 8 bytes is header
                    for (int j = 8; j < recipientPropertiesStream.Buffer.Length; j += 16)
                    {
                        byte[] recipientPropertyBuffer = new byte[16];

                        System.Array.Copy(recipientPropertiesStream.Buffer, j, recipientPropertyBuffer, 0, 16);

                        Property recipientProperty = new Property(recipientPropertyBuffer);

                        if (recipientProperty.Size > 0)
                        {
                            string recipientPropertyStreamName = "__substg1.0_" + String.Format("{0:X8}", recipientProperty.Tag);

                            Independentsoft.IO.StructuredStorage.Stream recipientPropertyStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries[recipientPropertyStreamName];

                            if (recipientPropertyStream != null && recipientPropertyStream.Buffer != null && recipientPropertyStream.Buffer.Length > 0)
                            {
                                recipientProperty.Value = new byte[recipientPropertyStream.Buffer.Length];
                                System.Array.Copy(recipientPropertyStream.Buffer, 0, recipientProperty.Value, 0, recipientProperty.Value.Length);
                            }
                        }

                        string recipientPropertyTagHex = String.Format("{0:X8}", recipientProperty.Tag);
                        recipientPropertyTable.Add(recipientPropertyTagHex, recipientProperty);
                    }
                }

                Recipient recipient = new Recipient();

                Property recipientDisplayTypeProperty = recipientPropertyTable.ContainsKey("39000003") ? recipientPropertyTable["39000003"] : null;

                if (recipientDisplayTypeProperty != null && recipientDisplayTypeProperty.Value != null)
                {
                    recipient.DisplayType = EnumUtil.ParseDisplayType(BitConverter.ToUInt32(recipientDisplayTypeProperty.Value, 0));
                }

                Property recipientObjectTypeProperty = recipientPropertyTable.ContainsKey("0FFE0003") ? recipientPropertyTable["0FFE0003"] : null;

                if (recipientObjectTypeProperty != null && recipientObjectTypeProperty.Value != null)
                {
                    recipient.ObjectType = EnumUtil.ParseObjectType(BitConverter.ToUInt32(recipientObjectTypeProperty.Value, 0));
                }

                Property recipientRecipientTypeProperty = recipientPropertyTable.ContainsKey("0C150003") ? recipientPropertyTable["0C150003"] : null;

                if (recipientRecipientTypeProperty != null && recipientRecipientTypeProperty.Value != null)
                {
                    recipient.RecipientType = EnumUtil.ParseRecipientType(BitConverter.ToUInt32(recipientRecipientTypeProperty.Value, 0));
                }

                Property recipientResponsibilityProperty = recipientPropertyTable.ContainsKey("0E0F000B") ? recipientPropertyTable["0E0F000B"] : null;

                if (recipientResponsibilityProperty != null && recipientResponsibilityProperty.Value != null)
                {
                    ushort recipientResponsibilityValue = BitConverter.ToUInt16(recipientResponsibilityProperty.Value, 0);

                    if (recipientResponsibilityValue > 0)
                    {
                        recipient.Responsibility = true;
                    }
                }

                Property recipientSendRichInfoProperty = recipientPropertyTable.ContainsKey("3A40000B") ? recipientPropertyTable["3A40000B"] : null;

                if (recipientSendRichInfoProperty != null && recipientSendRichInfoProperty.Value != null)
                {
                    ushort recipientSendRichInfoValue = BitConverter.ToUInt16(recipientSendRichInfoProperty.Value, 0);

                    if (recipientSendRichInfoValue > 0)
                    {
                        recipient.SendRichInfo = true;
                    }
                }

                Property recipientSendInternetEncodingProperty = recipientPropertyTable.ContainsKey("3A710003") ? recipientPropertyTable["3A710003"] : null;

                if (recipientSendInternetEncodingProperty != null && recipientSendInternetEncodingProperty.Value != null)
                {
                    recipient.SendInternetEncoding = BitConverter.ToInt32(recipientSendInternetEncodingProperty.Value, 0);
                }

                Independentsoft.IO.StructuredStorage.Stream recipientDisplayNameStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_3001" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_3002" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_3003" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientSmtpAddressStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_39FE" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientDisplayName7BitStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_39FF" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientTransmitableDisplayNameStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_3A20" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientOriginatingAddressTypeStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_403D" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientOriginatingEmailAddressStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_403E" + stringTypeMask];
                Independentsoft.IO.StructuredStorage.Stream recipientEntryIdStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_0FFF0102"];
                Independentsoft.IO.StructuredStorage.Stream recipientSearchKeyStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_300B0102"];
                Independentsoft.IO.StructuredStorage.Stream recipientInstanceKeyStream = (Independentsoft.IO.StructuredStorage.Stream)recipientStorage.DirectoryEntries["__substg1.0_0FF60102"];

                if (recipientDisplayNameStream != null && recipientDisplayNameStream.Buffer != null)
                {
                    recipient.DisplayName = encoding.GetString(recipientDisplayNameStream.Buffer, 0, recipientDisplayNameStream.Buffer.Length);
                }

                if (recipientAddressTypeStream != null && recipientAddressTypeStream.Buffer != null)
                {
                    recipient.AddressType = encoding.GetString(recipientAddressTypeStream.Buffer, 0, recipientAddressTypeStream.Buffer.Length);
                }

                if (recipientEmailAddressStream != null && recipientEmailAddressStream.Buffer != null)
                {
                    recipient.EmailAddress = encoding.GetString(recipientEmailAddressStream.Buffer, 0, recipientEmailAddressStream.Buffer.Length);
                }

                if (recipientSmtpAddressStream != null && recipientSmtpAddressStream.Buffer != null)
                {
                    recipient.SmtpAddress = encoding.GetString(recipientSmtpAddressStream.Buffer, 0, recipientSmtpAddressStream.Buffer.Length);
                }

                if (recipientDisplayName7BitStream != null && recipientDisplayName7BitStream.Buffer != null)
                {
                    recipient.DisplayName7Bit = encoding.GetString(recipientDisplayName7BitStream.Buffer, 0, recipientDisplayName7BitStream.Buffer.Length);
                }

                if (recipientTransmitableDisplayNameStream != null && recipientTransmitableDisplayNameStream.Buffer != null)
                {
                    recipient.TransmitableDisplayName = encoding.GetString(recipientTransmitableDisplayNameStream.Buffer, 0, recipientTransmitableDisplayNameStream.Buffer.Length);
                }

                if (recipientOriginatingAddressTypeStream != null && recipientOriginatingAddressTypeStream.Buffer != null)
                {
                    recipient.OriginatingAddressType = encoding.GetString(recipientOriginatingAddressTypeStream.Buffer, 0, recipientOriginatingAddressTypeStream.Buffer.Length);
                }

                if (recipientOriginatingEmailAddressStream != null && recipientOriginatingEmailAddressStream.Buffer != null)
                {
                    recipient.OriginatingEmailAddress = encoding.GetString(recipientOriginatingEmailAddressStream.Buffer, 0, recipientOriginatingEmailAddressStream.Buffer.Length);
                }

                if (recipientEntryIdStream != null && recipientEntryIdStream.Buffer != null)
                {
                    recipient.EntryId = recipientEntryIdStream.Buffer;
                }

                if (recipientSearchKeyStream != null && recipientSearchKeyStream.Buffer != null)
                {
                    recipient.SearchKey = recipientSearchKeyStream.Buffer;
                }

                if (recipientInstanceKeyStream != null && recipientInstanceKeyStream.Buffer != null)
                {
                    recipient.InstanceKey = recipientInstanceKeyStream.Buffer;
                }

                recipients.Add(recipient);
            }

            int storageCount = 0;

            for (int s = 0; s < directoryEntries.Count; s++)
            {
                if (directoryEntries[s] is Independentsoft.IO.StructuredStorage.Storage)
                {
                    storageCount++;
                }
            }

            //Attachments
            for (int i = 0; i < attachmentCount; i++)
            {
                IDictionary<String, Property> attachmentPropertyTable = new Dictionary<String, Property>();
                string attachmentStorageName = String.Format("__attach_version1.0_#{0:X8}", i);

                Independentsoft.IO.StructuredStorage.Storage attachmentStorage = (Independentsoft.IO.StructuredStorage.Storage)directoryEntries[attachmentStorageName];

                if (attachmentStorage == null)
                {
                    attachmentCount++;

                    if (attachmentCount > storageCount)
                    {
                        break;
                    }
                }
                else
                {
                    Independentsoft.IO.StructuredStorage.Stream attachmentPropertiesStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__properties_version1.0"];

                    Attachment attachment = new Attachment();

                    if (attachmentPropertiesStream != null && attachmentPropertiesStream.Buffer != null)
                    {
                        //first 8 bytes is header
                        for (int j = 8; j < attachmentPropertiesStream.Buffer.Length; j += 16)
                        {
                            byte[] attachmentPropertyBuffer = new byte[16];

                            System.Array.Copy(attachmentPropertiesStream.Buffer, j, attachmentPropertyBuffer, 0, 16);

                            Property attachmentProperty = new Property(attachmentPropertyBuffer);

                            if (attachmentProperty.Size > 0)
                            {
                                string attachmentPropertyStreamName = "__substg1.0_" + String.Format("{0:X8}", attachmentProperty.Tag);

                                if (attachmentStorage.DirectoryEntries[attachmentPropertyStreamName] is Independentsoft.IO.StructuredStorage.Stream)
                                {
                                    Independentsoft.IO.StructuredStorage.Stream attachmentPropertyStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries[attachmentPropertyStreamName];

                                    if (attachmentPropertyStream != null && attachmentPropertyStream.Buffer != null && attachmentPropertyStream.Buffer.Length > 0)
                                    {
                                        attachmentProperty.Value = new byte[attachmentPropertyStream.Buffer.Length];
                                        System.Array.Copy(attachmentPropertyStream.Buffer, 0, attachmentProperty.Value, 0, attachmentProperty.Value.Length);
                                    }
                                }
                                else if (attachmentStorage.DirectoryEntries[attachmentPropertyStreamName] is Independentsoft.IO.StructuredStorage.Storage)
                                {
                                    Independentsoft.IO.StructuredStorage.Storage attachmentPropertyStorage = (Independentsoft.IO.StructuredStorage.Storage)attachmentStorage.DirectoryEntries[attachmentPropertyStreamName];

                                    if (attachmentPropertyStorage != null && attachmentPropertyStorage.DirectoryEntries["__properties_version1.0"] != null)
                                    {
                                        Message message = new Message(attachmentPropertyStorage.DirectoryEntries, true);

                                        attachment.EmbeddedMessage = message;
                                    }
                                }
                            }

                            string attachmentPropertyTagHex = String.Format("{0:X8}", attachmentProperty.Tag);

                            if (!attachmentPropertyTable.ContainsKey(attachmentPropertyTagHex))
                            {
                                attachmentPropertyTable.Add(attachmentPropertyTagHex, attachmentProperty);
                            }
                        }
                    }

                    Property attachmentFlagsProperty = attachmentPropertyTable.ContainsKey("37140003") ? attachmentPropertyTable["37140003"] : null;

                    if (attachmentFlagsProperty != null && attachmentFlagsProperty.Value != null)
                    {
                        attachment.Flags = EnumUtil.ParseAttachmentFlags(BitConverter.ToUInt32(attachmentFlagsProperty.Value, 0));
                    }

                    Property attachmentMethodProperty = attachmentPropertyTable.ContainsKey("37050003") ? attachmentPropertyTable["37050003"] : null;

                    if (attachmentMethodProperty != null && attachmentMethodProperty.Value != null)
                    {
                        attachment.Method = EnumUtil.ParseAttachmentMethod(BitConverter.ToUInt32(attachmentMethodProperty.Value, 0));
                    }

                    Property attachmentMimeSequenceProperty = attachmentPropertyTable.ContainsKey("37100003") ? attachmentPropertyTable["37100003"] : null;

                    if (attachmentMimeSequenceProperty != null && attachmentMimeSequenceProperty.Value != null)
                    {
                        attachment.MimeSequence = BitConverter.ToUInt32(attachmentMimeSequenceProperty.Value, 0);
                    }

                    Property attachmentRenderingPositionProperty = attachmentPropertyTable.ContainsKey("370B0003") ? attachmentPropertyTable["370B0003"] : null;

                    if (attachmentRenderingPositionProperty != null && attachmentRenderingPositionProperty.Value != null)
                    {
                        attachment.RenderingPosition = BitConverter.ToUInt32(attachmentRenderingPositionProperty.Value, 0);
                    }

                    Property attachmentSizeProperty = attachmentPropertyTable.ContainsKey("0E200003") ? attachmentPropertyTable["0E200003"] : null;

                    if (attachmentSizeProperty != null && attachmentSizeProperty.Value != null)
                    {
                        attachment.Size = BitConverter.ToUInt32(attachmentSizeProperty.Value, 0);
                    }

                    Property attachmentObjectTypeProperty = attachmentPropertyTable.ContainsKey("0FFE0003") ? attachmentPropertyTable["0FFE0003"] : null;

                    if (attachmentObjectTypeProperty != null && attachmentObjectTypeProperty.Value != null)
                    {
                        attachment.ObjectType = EnumUtil.ParseObjectType(BitConverter.ToUInt32(attachmentObjectTypeProperty.Value, 0));
                    }

                    Property attachmentIsHiddenProperty = attachmentPropertyTable.ContainsKey("7FFE000B") ? attachmentPropertyTable["7FFE000B"] : null;

                    if (attachmentIsHiddenProperty != null && attachmentIsHiddenProperty.Value != null)
                    {
                        ushort attachmentIsHiddenValue = BitConverter.ToUInt16(attachmentIsHiddenProperty.Value, 0);

                        if (attachmentIsHiddenValue > 0)
                        {
                            attachment.IsHidden = true;
                        }
                    }

                    Property attachmentIsContactPhotoProperty = attachmentPropertyTable.ContainsKey("7FFF000B") ? attachmentPropertyTable["7FFF000B"] : null;

                    if (attachmentIsContactPhotoProperty != null && attachmentIsContactPhotoProperty.Value != null)
                    {
                        ushort attachmentIsContactValue = BitConverter.ToUInt16(attachmentIsContactPhotoProperty.Value, 0);

                        if (attachmentIsContactValue > 0)
                        {
                            attachment.IsContactPhoto = true;
                        }
                    }

                    Property attachmentCreationTimeProperty = attachmentPropertyTable.ContainsKey("30070040") ? attachmentPropertyTable["30070040"] : null;

                    if (attachmentCreationTimeProperty != null && attachmentCreationTimeProperty.Value != null)
                    {
                        uint creationTimeLow = BitConverter.ToUInt32(attachmentCreationTimeProperty.Value, 0);
                        ulong creationTimeHigh = BitConverter.ToUInt32(attachmentCreationTimeProperty.Value, 4);

                        if (creationTimeHigh > 0)
                        {
                            long ticks = creationTimeLow + (long)(creationTimeHigh << 32);

                            DateTime year1601 = new DateTime(1601, 1, 1);

                            try
                            {
                                attachment.CreationTime = year1601.AddTicks(ticks).ToLocalTime();
                            }
                            catch (Exception) //ignore wrong dates
                            {
                            }
                        }
                    }

                    Property attachmentLastModificationTimeProperty = attachmentPropertyTable.ContainsKey("30080040") ? attachmentPropertyTable["30080040"] : null;

                    if (attachmentLastModificationTimeProperty != null && attachmentLastModificationTimeProperty.Value != null)
                    {
                        uint lastModificationTimeLow = BitConverter.ToUInt32(attachmentLastModificationTimeProperty.Value, 0);
                        ulong lastModificationTimeHigh = BitConverter.ToUInt32(attachmentLastModificationTimeProperty.Value, 4);

                        if (lastModificationTimeHigh > 0)
                        {
                            long ticks = lastModificationTimeLow + (long)(lastModificationTimeHigh << 32);

                            DateTime year1601 = new DateTime(1601, 1, 1);

                            try
                            {
                                attachment.LastModificationTime = year1601.AddTicks(ticks).ToLocalTime();
                            }
                            catch (Exception) //ignore wrong dates
                            {
                            }
                        }
                    }

                    Independentsoft.IO.StructuredStorage.Stream attachmentAdditionalInfoStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_370F0102"];
                    Independentsoft.IO.StructuredStorage.Stream attachmentContentBaseStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3711" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentContentIdStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3712" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentContentLocationStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3713" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentContentDispositionStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3716" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentDataStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_37010102"];
                    Independentsoft.IO.StructuredStorage.Stream attachmentEncodingStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_37020102"];
                    Independentsoft.IO.StructuredStorage.Stream attachmentExtensionStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3703" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentFileNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3704" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentLongFileNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3707" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentLongPathNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_370D" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentMimeTagStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_370E" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentPathNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3708" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentRenderingStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_37090102"];
                    Independentsoft.IO.StructuredStorage.Stream attachmentTagStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_370A0102"];
                    Independentsoft.IO.StructuredStorage.Stream attachmentTransportNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_370C" + stringTypeMask];
                    Independentsoft.IO.StructuredStorage.Stream attachmentDisplayNameStream = (Independentsoft.IO.StructuredStorage.Stream)attachmentStorage.DirectoryEntries["__substg1.0_3001" + stringTypeMask];                  
                    Independentsoft.IO.StructuredStorage.Storage attachmentDataObjectStorage = null;

                    if (attachmentStorage.DirectoryEntries["__substg1.0_3701000D"] != null && attachmentStorage.DirectoryEntries["__substg1.0_3701000D"] is Independentsoft.IO.StructuredStorage.Storage)
                    {
                        attachmentDataObjectStorage = (Independentsoft.IO.StructuredStorage.Storage)attachmentStorage.DirectoryEntries["__substg1.0_3701000D"];

                        attachment.DataObjectStorage = attachmentDataObjectStorage;
                        attachment.PropertiesStream = attachmentPropertiesStream;
                    }

                    if (attachmentAdditionalInfoStream != null && attachmentAdditionalInfoStream.Buffer != null)
                    {
                        attachment.AdditionalInfo = attachmentAdditionalInfoStream.Buffer;
                    }

                    if (attachmentContentBaseStream != null && attachmentContentBaseStream.Buffer != null)
                    {
                        attachment.ContentBase = encoding.GetString(attachmentContentBaseStream.Buffer, 0, attachmentContentBaseStream.Buffer.Length);
                    }

                    if (attachmentContentIdStream != null && attachmentContentIdStream.Buffer != null)
                    {
                        attachment.ContentId = encoding.GetString(attachmentContentIdStream.Buffer, 0, attachmentContentIdStream.Buffer.Length);
                    }

                    if (attachmentContentLocationStream != null && attachmentContentLocationStream.Buffer != null)
                    {
                        attachment.ContentLocation = encoding.GetString(attachmentContentLocationStream.Buffer, 0, attachmentContentLocationStream.Buffer.Length);
                    }

                    if (attachmentContentDispositionStream != null && attachmentContentDispositionStream.Buffer != null)
                    {
                        attachment.ContentDisposition = encoding.GetString(attachmentContentDispositionStream.Buffer, 0, attachmentContentDispositionStream.Buffer.Length);
                    }

                    if (attachmentDataStream != null && attachmentDataStream.Buffer != null)
                    {
                        attachment.Data = attachmentDataStream.Buffer;
                    }

                    if (attachmentEncodingStream != null && attachmentEncodingStream.Buffer != null)
                    {
                        attachment.Encoding = attachmentEncodingStream.Buffer;
                    }

                    if (attachmentExtensionStream != null && attachmentExtensionStream.Buffer != null)
                    {
                        attachment.Extension = encoding.GetString(attachmentExtensionStream.Buffer, 0, attachmentExtensionStream.Buffer.Length);
                    }

                    if (attachmentFileNameStream != null && attachmentFileNameStream.Buffer != null)
                    {
                        attachment.FileName = encoding.GetString(attachmentFileNameStream.Buffer, 0, attachmentFileNameStream.Buffer.Length);
                    }

                    if (attachmentLongFileNameStream != null && attachmentLongFileNameStream.Buffer != null)
                    {
                        attachment.LongFileName = encoding.GetString(attachmentLongFileNameStream.Buffer, 0, attachmentLongFileNameStream.Buffer.Length);
                    }

                    if (attachmentLongPathNameStream != null && attachmentLongPathNameStream.Buffer != null)
                    {
                        attachment.LongPathName = encoding.GetString(attachmentLongPathNameStream.Buffer, 0, attachmentLongPathNameStream.Buffer.Length);
                    }

                    if (attachmentMimeTagStream != null && attachmentMimeTagStream.Buffer != null)
                    {
                        attachment.MimeTag = encoding.GetString(attachmentMimeTagStream.Buffer, 0, attachmentMimeTagStream.Buffer.Length);
                    }

                    if (attachmentPathNameStream != null && attachmentPathNameStream.Buffer != null)
                    {
                        attachment.PathName = encoding.GetString(attachmentPathNameStream.Buffer, 0, attachmentPathNameStream.Buffer.Length);
                    }

                    if (attachmentRenderingStream != null && attachmentRenderingStream.Buffer != null)
                    {
                        attachment.Rendering = attachmentRenderingStream.Buffer;
                    }

                    if (attachmentTagStream != null && attachmentTagStream.Buffer != null)
                    {
                        attachment.Tag = attachmentTagStream.Buffer;
                    }

                    if (attachmentTransportNameStream != null && attachmentTransportNameStream.Buffer != null)
                    {
                        attachment.TransportName = encoding.GetString(attachmentTransportNameStream.Buffer, 0, attachmentTransportNameStream.Buffer.Length);
                    }

                    if (attachmentDisplayNameStream != null && attachmentDisplayNameStream.Buffer != null)
                    {
                        attachment.DisplayName = encoding.GetString(attachmentDisplayNameStream.Buffer, 0, attachmentDisplayNameStream.Buffer.Length);
                    }

                    if (attachment.Data != null || attachment.DataObject != null || attachment.DataObjectStorage != null || attachment.EmbeddedMessage != null)
                    {
                        attachments.Add(attachment);
                    }
                }
            }
        }

        /// <summary>
        /// Gets stream to read from this message.
        /// </summary>
        /// <returns>A stream.</returns>
        public System.IO.Stream GetStream()
        {
            byte[] buffer = CreateMessage();
            MemoryStream memoryStream = new MemoryStream(buffer);

            return memoryStream;
        }

        /// <summary>
        /// Gets bytes to read from this message.
        /// </summary>
        /// <returns>Attachment as a byte array.</returns>
        public byte[] GetBytes()
        {
            return CreateMessage();
        }

        /// <summary>
        /// Saves this message to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false);
        }

        /// <summary>
        /// Saves this message to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        /// <param name="overwrite">True to overwrite existing file, otherwise false.</param>
        public void Save(string filePath, bool overwrite)
        {
            FileMode mode = FileMode.CreateNew;

            if (overwrite)
            {
                mode = FileMode.Create;
            }

            using (FileStream fileStream = new FileStream(filePath, mode, FileAccess.Write))
            {
                Save(fileStream);
            }
        }

        /// <summary>
        /// Saves this message to the specified stream.
        /// </summary>
        /// <param name="stream">A stream.</param>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public void Save(System.IO.Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            byte[] buffer = CreateMessage();
            stream.Write(buffer, 0, buffer.Length);
        }

        private byte[] CreateMessage()
        {
            CompoundFile compoundFile = new CompoundFile();
            compoundFile.MajorVersion = 4;
            compoundFile.Root.ClassId = new byte[] { 11, 13, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };

            MemoryStream guidMemoryStream = new MemoryStream();
            MemoryStream entryMemoryStream = new MemoryStream();
            MemoryStream stringMemoryStream = new MemoryStream();

            IList<byte[]> guidList = new List<byte[]>();

            guidList.Add(new byte[16]); //null guid
            guidList.Add(StandardPropertySet.Mapi);
            guidList.Add(StandardPropertySet.PublicStrings);

            IList<String> stringNameList = new List<String>();
            IDictionary<String, MemoryStream> namedPropertyStreams = new Dictionary<String, MemoryStream>();

            namedProperties = new List<NamedProperty>();

            DirectoryEntryList messagePropertiesDirectoryEntires = CreateMessageProperties(ref namedProperties);

            int stringOffset = 0;

            for (int i = 0; i < namedProperties.Count; i++)
            {
                int guidIndex = GetGuidIndex(namedProperties[i].Guid, guidList);

                if (guidIndex == -1 && namedProperties[i].Guid != null)
                {
                    guidList.Add(namedProperties[i].Guid);
                    guidIndex = guidList.Count - 1;
                }

                uint nameOrStringOffset = 0;
                ushort propertyKind = 0;

                if (namedProperties[i].Name != null) //string named property
                {
                    stringNameList.Add(namedProperties[i].Name);
                    propertyKind = 1;

                    nameOrStringOffset = (uint)stringOffset;
                }
                else //numerical named property
                {
                    nameOrStringOffset = namedProperties[i].Id;
                }

                ushort propertyIndex = (ushort)i;

                uint indexAndKind = (uint)(propertyIndex << 16);
                ushort guidAndKind = (ushort)(guidIndex << 1);

                if (propertyKind == 1)
                {
                    guidAndKind = (ushort)(guidAndKind + 1);
                }

                indexAndKind = indexAndKind + guidAndKind;

                byte[] entryBuffer = new byte[8];

                byte[] nameOrStringOffsetBuffer = BitConverter.GetBytes(nameOrStringOffset);
                byte[] indexAndKindBuffer = BitConverter.GetBytes(indexAndKind);

                System.Array.Copy(nameOrStringOffsetBuffer, 0, entryBuffer, 0, 4);
                System.Array.Copy(indexAndKindBuffer, 0, entryBuffer, 4, 4);

                entryMemoryStream.Write(entryBuffer, 0, 8);

                //Create streams
                if (propertyKind == 0) //numerical named property
                {
                    uint streamId = (uint)(0x1000 + ((namedProperties[i].Id ^ guidIndex << 1) % 0x1F));

                    streamId = (streamId << 16) | 0x00000102;

                    string streamName = String.Format("{0:X8}", streamId);

                    streamName = "__substg1.0_" + streamName;

                    if (namedPropertyStreams.ContainsKey(streamName))
                    {
                        MemoryStream memoryStream = namedPropertyStreams[streamName];
                        memoryStream.Write(entryBuffer, 0, entryBuffer.Length);
                    }
                    else
                    {
                        MemoryStream memoryStream = new MemoryStream();
                        memoryStream.Write(entryBuffer, 0, entryBuffer.Length);

                        namedPropertyStreams.Add(streamName, memoryStream);
                    }
                }
                else //string named property
                {
                    Crc crc = new Crc();
                    crc.Update(encoding.GetBytes(namedProperties[i].Name));
                    uint crcCheckSum = crc.Value;

                    uint streamId = (uint)(0x1000 + ((crcCheckSum ^ ((guidIndex << 1) | 1)) % 0x1F));

                    streamId = (streamId << 16) | 0x00000102;

                    string streamName = String.Format("{0:X8}", streamId);

                    streamName = "__substg1.0_" + streamName;

                    if (namedPropertyStreams.ContainsKey(streamName))
                    {
                        MemoryStream memoryStream = (MemoryStream)namedPropertyStreams[streamName];
                        memoryStream.Write(entryBuffer, 0, entryBuffer.Length);
                    }
                    else
                    {
                        MemoryStream memoryStream = new MemoryStream();
                        memoryStream.Write(entryBuffer, 0, entryBuffer.Length);

                        namedPropertyStreams.Add(streamName, memoryStream);
                    }
                }

                if (namedProperties[i].Name != null)
                {
                    byte[] stringBuffer = System.Text.Encoding.Unicode.GetBytes(namedProperties[i].Name);
                    int addNullsCount = stringBuffer.Length % 4;
                    stringOffset += stringBuffer.Length + addNullsCount + 4;
                }
            }

            Storage namedPropertyMappingStorage = new Storage("__nameid_version1.0");

            Independentsoft.IO.StructuredStorage.Stream entryStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00030102", entryMemoryStream.ToArray());

            for (int i = 3; i < guidList.Count; i++)
            {
                byte[] guidBuffer = guidList[i];
                guidMemoryStream.Write(guidBuffer, 0, guidBuffer.Length);
            }

            Independentsoft.IO.StructuredStorage.Stream guidStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00020102", guidMemoryStream.ToArray());

            for (int i = 0; i < stringNameList.Count; i++)
            {
                byte[] stringBuffer = System.Text.Encoding.Unicode.GetBytes(stringNameList[i]);
                byte[] stringNameLengthBuffer = BitConverter.GetBytes(stringBuffer.Length);

                stringMemoryStream.Write(stringNameLengthBuffer, 0, 4);
                stringMemoryStream.Write(stringBuffer, 0, stringBuffer.Length);

                int addNullsCount = stringBuffer.Length % 4;

                if (addNullsCount > 0)
                {
                    byte[] addNullsBuffer = new byte[addNullsCount];
                    stringMemoryStream.Write(addNullsBuffer, 0, addNullsBuffer.Length);
                }
            }

            Independentsoft.IO.StructuredStorage.Stream stringStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00040102", stringMemoryStream.ToArray());

            namedPropertyMappingStorage.DirectoryEntries.Add(guidStream);
            namedPropertyMappingStorage.DirectoryEntries.Add(entryStream);
            namedPropertyMappingStorage.DirectoryEntries.Add(stringStream);

            foreach (string key in namedPropertyStreams.Keys)
            {
                MemoryStream memoryStream = (MemoryStream)namedPropertyStreams[key];
                Independentsoft.IO.StructuredStorage.Stream currentNamedPropertyStream = new Independentsoft.IO.StructuredStorage.Stream(key, memoryStream.ToArray());
                namedPropertyMappingStorage.DirectoryEntries.Add(currentNamedPropertyStream);
            }

            for (int i = 0; i < messagePropertiesDirectoryEntires.Count; i++)
            {
                compoundFile.Root.DirectoryEntries.Add(messagePropertiesDirectoryEntires[i]);
            }

            compoundFile.Root.DirectoryEntries.Add(namedPropertyMappingStorage);

            return compoundFile.GetBytes();
        }

        private DirectoryEntryList CreateMessageProperties(ref IList<NamedProperty> namedProperties)
        {
            DirectoryEntryList directoryEntries = new DirectoryEntryList();
            MemoryStream propertiesMemoryStream = new MemoryStream();

            //first 32 bytes is header
            uint headerZeroValue = 0;
            byte[] headerZeroBytes = BitConverter.GetBytes(headerZeroValue);
            byte[] recipientCountBytes = BitConverter.GetBytes(recipients.Count);
            byte[] attachmentCountBytes = BitConverter.GetBytes(attachments.Count);

            propertiesMemoryStream.Write(headerZeroBytes, 0, 4); //reserved1
            propertiesMemoryStream.Write(headerZeroBytes, 0, 4); //reserved2
            propertiesMemoryStream.Write(recipientCountBytes, 0, 4); //nextRecipientId
            propertiesMemoryStream.Write(attachmentCountBytes, 0, 4); //nextAttachmentId
            propertiesMemoryStream.Write(recipientCountBytes, 0, 4); //recipientCount
            propertiesMemoryStream.Write(attachmentCountBytes, 0, 4); //attachmentCount

            if (!isEmbedded)
            {
                propertiesMemoryStream.Write(headerZeroBytes, 0, 4); //reserved3
                propertiesMemoryStream.Write(headerZeroBytes, 0, 4); //reserved4
            }

            if (encoding == System.Text.Encoding.Unicode || encoding is System.Text.UnicodeEncoding)
            {
                stringTypeMask = "001F";
                stringTypeHexMask = 0x001F;
                multiValueStringTypeMask = "101F";
                multiValueStringTypeHexMask = 0x101F;

                if (!storeSupportMasks.Contains(StoreSupportMask.Unicode))
                {
                    storeSupportMasks.Add(StoreSupportMask.Unicode);
                }
            }
            else
            {
                storeSupportMasks.Remove(StoreSupportMask.Unicode);
            }

            if (storeSupportMasks != null)
            {
                Property storeSupportMaskProperty = new Property();
                storeSupportMaskProperty.Tag = 0x340D0003;
                storeSupportMaskProperty.Type = PropertyType.Integer32;
                storeSupportMaskProperty.Value = BitConverter.GetBytes(EnumUtil.ParseStoreSupportMask(storeSupportMasks));
                storeSupportMaskProperty.IsReadable = true;
                storeSupportMaskProperty.IsWriteable = true;

                propertiesMemoryStream.Write(storeSupportMaskProperty.ToBytes(), 0, 16);
            }

            if (messageClass != null)
            {
                byte[] messageClassBuffer = encoding.GetBytes(messageClass);
                Independentsoft.IO.StructuredStorage.Stream messageClassStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_001A" + stringTypeMask, messageClassBuffer);
                directoryEntries.Add(messageClassStream);

                Property messageClassProperty = new Property();
                messageClassProperty.Tag = 0x001A << 16 | stringTypeHexMask;
                messageClassProperty.Type = PropertyType.String8;
                messageClassProperty.Size = (uint)(messageClassBuffer.Length + encoding.GetBytes("\0").Length);
                messageClassProperty.IsReadable = true;
                messageClassProperty.IsWriteable = true;

                propertiesMemoryStream.Write(messageClassProperty.ToBytes(), 0, 16);
            }

            if (subject != null)
            {
                byte[] subjectBuffer = encoding.GetBytes(subject);
                Independentsoft.IO.StructuredStorage.Stream subjectStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0037" + stringTypeMask, subjectBuffer);
                directoryEntries.Add(subjectStream);

                Property subjectProperty = new Property();
                subjectProperty.Tag = 0x0037 << 16 | stringTypeHexMask;
                subjectProperty.Type = PropertyType.String8;
                subjectProperty.Size = (uint)(subjectBuffer.Length + encoding.GetBytes("\0").Length);
                subjectProperty.IsReadable = true;
                subjectProperty.IsWriteable = true;

                propertiesMemoryStream.Write(subjectProperty.ToBytes(), 0, 16);
            }

            if (subjectPrefix != null)
            {
                byte[] subjectPrefixBuffer = encoding.GetBytes(subjectPrefix);
                Independentsoft.IO.StructuredStorage.Stream subjectPrefixStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_003D" + stringTypeMask, subjectPrefixBuffer);
                directoryEntries.Add(subjectPrefixStream);

                Property subjectPrefixProperty = new Property();
                subjectPrefixProperty.Tag = 0x003D << 16 | stringTypeHexMask;
                subjectPrefixProperty.Type = PropertyType.String8;
                subjectPrefixProperty.Size = (uint)(subjectPrefixBuffer.Length + encoding.GetBytes("\0").Length);
                subjectPrefixProperty.IsReadable = true;
                subjectPrefixProperty.IsWriteable = true;

                propertiesMemoryStream.Write(subjectPrefixProperty.ToBytes(), 0, 16);
            }

            if (conversationTopic != null)
            {
                byte[] conversationTopicBuffer = encoding.GetBytes(conversationTopic);
                Independentsoft.IO.StructuredStorage.Stream conversationTopicStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0070" + stringTypeMask, conversationTopicBuffer);
                directoryEntries.Add(conversationTopicStream);

                Property conversationTopicProperty = new Property();
                conversationTopicProperty.Tag = 0x0070 << 16 | stringTypeHexMask;
                conversationTopicProperty.Type = PropertyType.String8;
                conversationTopicProperty.Size = (uint)(conversationTopicBuffer.Length + encoding.GetBytes("\0").Length);
                conversationTopicProperty.IsReadable = true;
                conversationTopicProperty.IsWriteable = true;

                propertiesMemoryStream.Write(conversationTopicProperty.ToBytes(), 0, 16);
            }

            if (displayBcc != null)
            {
                byte[] displayBccBuffer = encoding.GetBytes(displayBcc);
                Independentsoft.IO.StructuredStorage.Stream displayBccStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0E02" + stringTypeMask, displayBccBuffer);
                directoryEntries.Add(displayBccStream);

                Property displayBccProperty = new Property();
                displayBccProperty.Tag = 0x0E02 << 16 | stringTypeHexMask;
                displayBccProperty.Type = PropertyType.String8;
                displayBccProperty.Size = (uint)(displayBccBuffer.Length + encoding.GetBytes("\0").Length);
                displayBccProperty.IsReadable = true;
                displayBccProperty.IsWriteable = true;

                propertiesMemoryStream.Write(displayBccProperty.ToBytes(), 0, 16);
            }

            if (displayCc != null)
            {
                byte[] displayCcBuffer = encoding.GetBytes(displayCc);
                Independentsoft.IO.StructuredStorage.Stream displayCcStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0E03" + stringTypeMask, displayCcBuffer);
                directoryEntries.Add(displayCcStream);

                Property displayCcProperty = new Property();
                displayCcProperty.Tag = 0x0E03 << 16 | stringTypeHexMask;
                displayCcProperty.Type = PropertyType.String8;
                displayCcProperty.Size = (uint)(displayCcBuffer.Length + encoding.GetBytes("\0").Length);
                displayCcProperty.IsReadable = true;
                displayCcProperty.IsWriteable = true;

                propertiesMemoryStream.Write(displayCcProperty.ToBytes(), 0, 16);
            }

            if (displayTo != null)
            {
                byte[] displayToBuffer = encoding.GetBytes(displayTo);
                Independentsoft.IO.StructuredStorage.Stream displayToStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0E04" + stringTypeMask, displayToBuffer);
                directoryEntries.Add(displayToStream);

                Property displayToProperty = new Property();
                displayToProperty.Tag = 0x0E04 << 16 | stringTypeHexMask;
                displayToProperty.Type = PropertyType.String8;
                displayToProperty.Size = (uint)(displayToBuffer.Length + encoding.GetBytes("\0").Length);
                displayToProperty.IsReadable = true;
                displayToProperty.IsWriteable = true;

                propertiesMemoryStream.Write(displayToProperty.ToBytes(), 0, 16);
            }

            if (originalDisplayTo != null)
            {
                byte[] originalDisplayToBuffer = encoding.GetBytes(originalDisplayTo);
                Independentsoft.IO.StructuredStorage.Stream originalDisplayToStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0074" + stringTypeMask, originalDisplayToBuffer);
                directoryEntries.Add(originalDisplayToStream);

                Property originalDisplayToProperty = new Property();
                originalDisplayToProperty.Tag = 0x0074 << 16 | stringTypeHexMask;
                originalDisplayToProperty.Type = PropertyType.String8;
                originalDisplayToProperty.Size = (uint)(originalDisplayToBuffer.Length + encoding.GetBytes("\0").Length);
                originalDisplayToProperty.IsReadable = true;
                originalDisplayToProperty.IsWriteable = true;

                propertiesMemoryStream.Write(originalDisplayToProperty.ToBytes(), 0, 16);
            }

            if (replyTo != null)  //have to set PR_REPLY_RECIPIENT_NAMES and PR_REPLY_RECIPIENT_ENTRIES
            {
                byte[] replyToBuffer = encoding.GetBytes(replyTo);
                Independentsoft.IO.StructuredStorage.Stream replyToStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0050" + stringTypeMask, replyToBuffer);
                directoryEntries.Add(replyToStream);

                Property replyToProperty = new Property();
                replyToProperty.Tag = 0x0050 << 16 | stringTypeHexMask;
                replyToProperty.Type = PropertyType.String8;
                replyToProperty.Size = (uint)(replyToBuffer.Length + encoding.GetBytes("\0").Length);
                replyToProperty.IsReadable = true;
                replyToProperty.IsWriteable = true;

                propertiesMemoryStream.Write(replyToProperty.ToBytes(), 0, 16);

                //set PR_REPLY_RECIPIENT_ENTRIES

                byte[] replyToEntries = Util.CreateReplyToEntries(replyTo);

                Independentsoft.IO.StructuredStorage.Stream replyToEntriesStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_004F0102", replyToEntries);
                directoryEntries.Add(replyToEntriesStream);

                Property replyToEntriesProperty = new Property();
                replyToEntriesProperty.Tag = 0x004F0102;
                replyToEntriesProperty.Type = PropertyType.Binary;
                replyToEntriesProperty.Size = (uint)replyToEntries.Length;
                replyToEntriesProperty.IsReadable = true;
                replyToEntriesProperty.IsWriteable = true;

                propertiesMemoryStream.Write(replyToEntriesProperty.ToBytes(), 0, 16);
            }

            if (normalizedSubject != null)
            {
                byte[] normalizedSubjectBuffer = encoding.GetBytes(normalizedSubject);
                Independentsoft.IO.StructuredStorage.Stream normalizedSubjectStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0E1D" + stringTypeMask, normalizedSubjectBuffer);
                directoryEntries.Add(normalizedSubjectStream);

                Property normalizedSubjectProperty = new Property();
                normalizedSubjectProperty.Tag = 0x0E1D << 16 | stringTypeHexMask;
                normalizedSubjectProperty.Type = PropertyType.String8;
                normalizedSubjectProperty.Size = (uint)(normalizedSubjectBuffer.Length + encoding.GetBytes("\0").Length);
                normalizedSubjectProperty.IsReadable = true;
                normalizedSubjectProperty.IsWriteable = true;

                propertiesMemoryStream.Write(normalizedSubjectProperty.ToBytes(), 0, 16);
            }

            if (body != null)
            {
                byte[] bodyBuffer = encoding.GetBytes(body);
                Independentsoft.IO.StructuredStorage.Stream bodyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_1000" + stringTypeMask, bodyBuffer);
                directoryEntries.Add(bodyStream);

                Property bodyProperty = new Property();
                bodyProperty.Tag = 0x1000 << 16 | stringTypeHexMask;
                bodyProperty.Type = PropertyType.String8;
                bodyProperty.Size = (uint)(bodyBuffer.Length + encoding.GetBytes("\0").Length);
                bodyProperty.IsReadable = true;
                bodyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(bodyProperty.ToBytes(), 0, 16);
            }

            if (rtfCompressed != null)
            {
                Independentsoft.IO.StructuredStorage.Stream rtfCompressedStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_10090102", rtfCompressed);
                directoryEntries.Add(rtfCompressedStream);

                Property rtfCompressedProperty = new Property();
                rtfCompressedProperty.Tag = 0x10090102;
                rtfCompressedProperty.Type = PropertyType.Binary;
                rtfCompressedProperty.Size = (uint)rtfCompressed.Length;
                rtfCompressedProperty.IsReadable = true;
                rtfCompressedProperty.IsWriteable = true;

                propertiesMemoryStream.Write(rtfCompressedProperty.ToBytes(), 0, 16);
            }

            if (searchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream searchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_300B0102", searchKey);
                directoryEntries.Add(searchKeyStream);

                Property searchKeyProperty = new Property();
                searchKeyProperty.Tag = 0x300B0102;
                searchKeyProperty.Type = PropertyType.Binary;
                searchKeyProperty.Size = (uint)searchKey.Length;
                searchKeyProperty.IsReadable = true;
                searchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(searchKeyProperty.ToBytes(), 0, 16);
            }

            if (changeKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream changeKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_65E20102", changeKey);
                directoryEntries.Add(changeKeyStream);

                Property changeKeyProperty = new Property();
                changeKeyProperty.Tag = 0x65E20102;
                changeKeyProperty.Type = PropertyType.Binary;
                changeKeyProperty.Size = (uint)changeKey.Length;
                changeKeyProperty.IsReadable = true;
                changeKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(changeKeyProperty.ToBytes(), 0, 16);
            }

            if (entryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream entryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0FFF0102", entryId);
                directoryEntries.Add(entryIdStream);

                Property entryIdProperty = new Property();
                entryIdProperty.Tag = 0x0FFF0102;
                entryIdProperty.Type = PropertyType.Binary;
                entryIdProperty.Size = (uint)entryId.Length;
                entryIdProperty.IsReadable = true;
                entryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(entryIdProperty.ToBytes(), 0, 16);
            }

            if (readReceiptEntryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream readReceiptEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00460102", readReceiptEntryId);
                directoryEntries.Add(readReceiptEntryIdStream);

                Property readReceiptEntryIdProperty = new Property();
                readReceiptEntryIdProperty.Tag = 0x00460102;
                readReceiptEntryIdProperty.Type = PropertyType.Binary;
                readReceiptEntryIdProperty.Size = (uint)readReceiptEntryId.Length;
                readReceiptEntryIdProperty.IsReadable = true;
                readReceiptEntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(readReceiptEntryIdProperty.ToBytes(), 0, 16);
            }

            if (readReceiptSearchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream readReceiptSearchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00530102", readReceiptSearchKey);
                directoryEntries.Add(readReceiptSearchKeyStream);

                Property readReceiptSearchKeyProperty = new Property();
                readReceiptSearchKeyProperty.Tag = 0x00530102;
                readReceiptSearchKeyProperty.Type = PropertyType.Binary;
                readReceiptSearchKeyProperty.Size = (uint)readReceiptSearchKey.Length;
                readReceiptSearchKeyProperty.IsReadable = true;
                readReceiptSearchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(readReceiptSearchKeyProperty.ToBytes(), 0, 16);
            }

            if (creationTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = creationTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property creationTimeProperty = new Property();
                creationTimeProperty.Tag = 0x30070040;
                creationTimeProperty.Type = PropertyType.Time;
                creationTimeProperty.Value = ticksBytes;
                creationTimeProperty.IsReadable = true;
                creationTimeProperty.IsWriteable = false;

                propertiesMemoryStream.Write(creationTimeProperty.ToBytes(), 0, 16);
            }

            if (lastModificationTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = lastModificationTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property lastModificationTimeProperty = new Property();
                lastModificationTimeProperty.Tag = 0x30080040;
                lastModificationTimeProperty.Type = PropertyType.Time;
                lastModificationTimeProperty.Value = ticksBytes;
                lastModificationTimeProperty.IsReadable = true;
                lastModificationTimeProperty.IsWriteable = false;

                propertiesMemoryStream.Write(lastModificationTimeProperty.ToBytes(), 0, 16);
            }

            if (messageDeliveryTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = messageDeliveryTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property messageDeliveryTimeProperty = new Property();
                messageDeliveryTimeProperty.Tag = 0x0E060040;
                messageDeliveryTimeProperty.Type = PropertyType.Time;
                messageDeliveryTimeProperty.Value = ticksBytes;
                messageDeliveryTimeProperty.IsReadable = true;
                messageDeliveryTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(messageDeliveryTimeProperty.ToBytes(), 0, 16);
            }

            if (clientSubmitTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = clientSubmitTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property clientSubmitTimeProperty = new Property();
                clientSubmitTimeProperty.Tag = 0x00390040;
                clientSubmitTimeProperty.Type = PropertyType.Time;
                clientSubmitTimeProperty.Value = ticksBytes;
                clientSubmitTimeProperty.IsReadable = true;
                clientSubmitTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(clientSubmitTimeProperty.ToBytes(), 0, 16);
            }

            if (deferredDeliveryTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = deferredDeliveryTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property deferredDeliveryTimeProperty = new Property();
                deferredDeliveryTimeProperty.Tag = 0x000F0040;
                deferredDeliveryTimeProperty.Type = PropertyType.Time;
                deferredDeliveryTimeProperty.Value = ticksBytes;
                deferredDeliveryTimeProperty.IsReadable = true;
                deferredDeliveryTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(deferredDeliveryTimeProperty.ToBytes(), 0, 16);
            }

            if (providerSubmitTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = providerSubmitTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property providerSubmitTimeProperty = new Property();
                providerSubmitTimeProperty.Tag = 0x00480040;
                providerSubmitTimeProperty.Type = PropertyType.Time;
                providerSubmitTimeProperty.Value = ticksBytes;
                providerSubmitTimeProperty.IsReadable = true;
                providerSubmitTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(providerSubmitTimeProperty.ToBytes(), 0, 16);
            }

            if (reportTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = reportTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property reportTimeProperty = new Property();
                reportTimeProperty.Tag = 0x00320040;
                reportTimeProperty.Type = PropertyType.Time;
                reportTimeProperty.Value = ticksBytes;
                reportTimeProperty.IsReadable = true;
                reportTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reportTimeProperty.ToBytes(), 0, 16);
            }

            if (lastVerbExecutionTime.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = lastVerbExecutionTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property lastVerbExecutionTimeProperty = new Property();
                lastVerbExecutionTimeProperty.Tag = 0x10820040;
                lastVerbExecutionTimeProperty.Type = PropertyType.Time;
                lastVerbExecutionTimeProperty.Value = ticksBytes;
                lastVerbExecutionTimeProperty.IsReadable = true;
                lastVerbExecutionTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(lastVerbExecutionTimeProperty.ToBytes(), 0, 16);
            }

            if (reportText != null)
            {
                byte[] reportTextBuffer = encoding.GetBytes(reportText);
                Independentsoft.IO.StructuredStorage.Stream reportTextStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_1001" + stringTypeMask, reportTextBuffer);
                directoryEntries.Add(reportTextStream);

                Property reportTextProperty = new Property();
                reportTextProperty.Tag = (uint)0x1001 << 16 | stringTypeHexMask;
                reportTextProperty.Type = PropertyType.String8;
                reportTextProperty.Size = (uint)(reportTextBuffer.Length + encoding.GetBytes("\0").Length);
                reportTextProperty.IsReadable = true;
                reportTextProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reportTextProperty.ToBytes(), 0, 16);
            }

            if (creatorName != null)
            {
                byte[] creatorNameBuffer = encoding.GetBytes(creatorName);
                Independentsoft.IO.StructuredStorage.Stream creatorNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3FF8" + stringTypeMask, creatorNameBuffer);
                directoryEntries.Add(creatorNameStream);

                Property creatorNameProperty = new Property();
                creatorNameProperty.Tag = (uint)0x3FF8 << 16 | stringTypeHexMask;
                creatorNameProperty.Type = PropertyType.String8;
                creatorNameProperty.Size = (uint)(creatorNameBuffer.Length + encoding.GetBytes("\0").Length);
                creatorNameProperty.IsReadable = true;
                creatorNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(creatorNameProperty.ToBytes(), 0, 16);
            }

            if (lastModifierName != null)
            {
                byte[] lastModifierNameBuffer = encoding.GetBytes(lastModifierName);
                Independentsoft.IO.StructuredStorage.Stream lastModifierNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3FFA" + stringTypeMask, lastModifierNameBuffer);
                directoryEntries.Add(lastModifierNameStream);

                Property lastModifierNameProperty = new Property();
                lastModifierNameProperty.Tag = (uint)0x3FFA << 16 | stringTypeHexMask;
                lastModifierNameProperty.Type = PropertyType.String8;
                lastModifierNameProperty.Size = (uint)(lastModifierNameBuffer.Length + encoding.GetBytes("\0").Length);
                lastModifierNameProperty.IsReadable = true;
                lastModifierNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(lastModifierNameProperty.ToBytes(), 0, 16);
            }

            if (internetMessageId != null)
            {
                byte[] internetMessageIdBuffer = encoding.GetBytes(internetMessageId);
                Independentsoft.IO.StructuredStorage.Stream internetMessageIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_1035" + stringTypeMask, internetMessageIdBuffer);
                directoryEntries.Add(internetMessageIdStream);

                Property internetMessageIdProperty = new Property();
                internetMessageIdProperty.Tag = (uint)0x1035 << 16 | stringTypeHexMask;
                internetMessageIdProperty.Type = PropertyType.String8;
                internetMessageIdProperty.Size = (uint)(internetMessageIdBuffer.Length + encoding.GetBytes("\0").Length);
                internetMessageIdProperty.IsReadable = true;
                internetMessageIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(internetMessageIdProperty.ToBytes(), 0, 16);
            }

            if (inReplyTo != null)
            {
                byte[] inReplyToBuffer = encoding.GetBytes(inReplyTo);
                Independentsoft.IO.StructuredStorage.Stream inReplyToStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_1042" + stringTypeMask, inReplyToBuffer);
                directoryEntries.Add(inReplyToStream);

                Property inReplyToProperty = new Property();
                inReplyToProperty.Tag = (uint)0x1042 << 16 | stringTypeHexMask;
                inReplyToProperty.Type = PropertyType.String8;
                inReplyToProperty.Size = (uint)(inReplyToBuffer.Length + encoding.GetBytes("\0").Length);
                inReplyToProperty.IsReadable = true;
                inReplyToProperty.IsWriteable = true;

                propertiesMemoryStream.Write(inReplyToProperty.ToBytes(), 0, 16);
            }

            if (internetReferences != null)
            {
                byte[] internetReferencesBuffer = encoding.GetBytes(internetReferences);
                Independentsoft.IO.StructuredStorage.Stream internetReferencesStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_1039" + stringTypeMask, internetReferencesBuffer);
                directoryEntries.Add(internetReferencesStream);

                Property internetReferencesProperty = new Property();
                internetReferencesProperty.Tag = (uint)0x1039 << 16 | stringTypeHexMask;
                internetReferencesProperty.Type = PropertyType.String8;
                internetReferencesProperty.Size = (uint)(internetReferencesBuffer.Length + encoding.GetBytes("\0").Length);
                internetReferencesProperty.IsReadable = true;
                internetReferencesProperty.IsWriteable = true;

                propertiesMemoryStream.Write(internetReferencesProperty.ToBytes(), 0, 16);
            }

            if (messageCodePage > 0)
            {
                Property messageCodePageProperty = new Property();
                messageCodePageProperty.Tag = 0x3FFD0003;
                messageCodePageProperty.Type = PropertyType.Integer32;
                messageCodePageProperty.Value = BitConverter.GetBytes(messageCodePage);
                messageCodePageProperty.IsReadable = true;
                messageCodePageProperty.IsWriteable = true;

                propertiesMemoryStream.Write(messageCodePageProperty.ToBytes(), 0, 16);
            }

            if (iconIndex > 0)
            {
                Property iconIndexProperty = new Property();
                iconIndexProperty.Tag = 0x10800003;
                iconIndexProperty.Type = PropertyType.Integer32;
                iconIndexProperty.Value = BitConverter.GetBytes(iconIndex);
                iconIndexProperty.IsReadable = true;
                iconIndexProperty.IsWriteable = true;

                propertiesMemoryStream.Write(iconIndexProperty.ToBytes(), 0, 16);
            }

            if (messageSize > 0)
            {
                Property messageSizeProperty = new Property();
                messageSizeProperty.Tag = 0x0E080003;
                messageSizeProperty.Type = PropertyType.Integer32;
                messageSizeProperty.Value = BitConverter.GetBytes(messageSize);
                messageSizeProperty.IsReadable = true;
                messageSizeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(messageSizeProperty.ToBytes(), 0, 16);
            }

            if (messageFlags != null && messageFlags.Count > 0)
            {
                Property messageFlagsProperty = new Property();
                messageFlagsProperty.Tag = 0x0E070003;
                messageFlagsProperty.Type = PropertyType.Integer32;
                messageFlagsProperty.Value = BitConverter.GetBytes(EnumUtil.ParseMessageFlag(messageFlags));
                messageFlagsProperty.IsReadable = true;
                messageFlagsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(messageFlagsProperty.ToBytes(), 0, 16);
            }

            if (internetCodePage > 0)
            {
                Property internetCodePageProperty = new Property();
                internetCodePageProperty.Tag = 0x3FDE0003;
                internetCodePageProperty.Type = PropertyType.Integer32;
                internetCodePageProperty.Value = BitConverter.GetBytes(internetCodePage);
                internetCodePageProperty.IsReadable = true;
                internetCodePageProperty.IsWriteable = true;

                propertiesMemoryStream.Write(internetCodePageProperty.ToBytes(), 0, 16);
            }

            if (conversationIndex != null)
            {
                Independentsoft.IO.StructuredStorage.Stream conversationIndexStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00710102", conversationIndex);
                directoryEntries.Add(conversationIndexStream);

                Property conversationIndexProperty = new Property();
                conversationIndexProperty.Tag = 0x00710102;
                conversationIndexProperty.Type = PropertyType.Binary;
                conversationIndexProperty.Size = (uint)conversationIndex.Length;
                conversationIndexProperty.IsReadable = true;
                conversationIndexProperty.IsWriteable = true;

                propertiesMemoryStream.Write(conversationIndexProperty.ToBytes(), 0, 16);
            }

            if (isHidden)
            {
                Property isHiddenProperty = new Property();
                isHiddenProperty.Tag = 0x10F4000B;
                isHiddenProperty.Type = PropertyType.Boolean;
                isHiddenProperty.Value = BitConverter.GetBytes(1);
                isHiddenProperty.IsReadable = true;
                isHiddenProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isHiddenProperty.ToBytes(), 0, 16);
            }

            if (isReadOnly)
            {
                Property isReadOnlyProperty = new Property();
                isReadOnlyProperty.Tag = 0x10F6000B;
                isReadOnlyProperty.Type = PropertyType.Boolean;
                isReadOnlyProperty.Value = BitConverter.GetBytes(1);
                isReadOnlyProperty.IsReadable = true;
                isReadOnlyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isReadOnlyProperty.ToBytes(), 0, 16);
            }

            if (isSystem)
            {
                Property isSystemProperty = new Property();
                isSystemProperty.Tag = 0x10F5000B;
                isSystemProperty.Type = PropertyType.Boolean;
                isSystemProperty.Value = BitConverter.GetBytes(1);
                isSystemProperty.IsReadable = true;
                isSystemProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isSystemProperty.ToBytes(), 0, 16);
            }

            if (disableFullFidelity)
            {
                Property disableFullFidelityProperty = new Property();
                disableFullFidelityProperty.Tag = 0x10F2000B;
                disableFullFidelityProperty.Type = PropertyType.Boolean;
                disableFullFidelityProperty.Value = BitConverter.GetBytes(1);
                disableFullFidelityProperty.IsReadable = true;
                disableFullFidelityProperty.IsWriteable = true;

                propertiesMemoryStream.Write(disableFullFidelityProperty.ToBytes(), 0, 16);
            }

            if (attachments.Count > 0)
            {
                Property hasAttachmentProperty = new Property();
                hasAttachmentProperty.Tag = 0x0E1B000B;
                hasAttachmentProperty.Type = PropertyType.Boolean;
                hasAttachmentProperty.Value = BitConverter.GetBytes(1);
                hasAttachmentProperty.IsReadable = true;
                hasAttachmentProperty.IsWriteable = true;

                propertiesMemoryStream.Write(hasAttachmentProperty.ToBytes(), 0, 16);
            }

            if (rtfInSync)
            {
                Property rtfInSyncProperty = new Property();
                rtfInSyncProperty.Tag = 0x0E1F000B;
                rtfInSyncProperty.Type = PropertyType.Boolean;
                rtfInSyncProperty.Value = BitConverter.GetBytes(1);
                rtfInSyncProperty.IsReadable = true;
                rtfInSyncProperty.IsWriteable = true;

                propertiesMemoryStream.Write(rtfInSyncProperty.ToBytes(), 0, 16);
            }

            if (readReceiptRequested)
            {
                Property readReceiptRequestedProperty = new Property();
                readReceiptRequestedProperty.Tag = 0x0029000B;
                readReceiptRequestedProperty.Type = PropertyType.Boolean;
                readReceiptRequestedProperty.Value = BitConverter.GetBytes(1);
                readReceiptRequestedProperty.IsReadable = true;
                readReceiptRequestedProperty.IsWriteable = true;

                propertiesMemoryStream.Write(readReceiptRequestedProperty.ToBytes(), 0, 16);
            }

            if (deliveryReportRequested)
            {
                Property deliveryReportRequestedProperty = new Property();
                deliveryReportRequestedProperty.Tag = 0x0023000B;
                deliveryReportRequestedProperty.Type = PropertyType.Boolean;
                deliveryReportRequestedProperty.Value = BitConverter.GetBytes(1);
                deliveryReportRequestedProperty.IsReadable = true;
                deliveryReportRequestedProperty.IsWriteable = true;

                propertiesMemoryStream.Write(deliveryReportRequestedProperty.ToBytes(), 0, 16);
            }

            if (bodyHtml != null)
            {
                Independentsoft.IO.StructuredStorage.Stream bodyHtmlStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_10130102", bodyHtml);
                directoryEntries.Add(bodyHtmlStream);

                Property bodyHtmlProperty = new Property();
                bodyHtmlProperty.Tag = 0x10130102;
                bodyHtmlProperty.Type = PropertyType.Binary;
                bodyHtmlProperty.Size = (uint)bodyHtml.Length;
                bodyHtmlProperty.IsReadable = true;
                bodyHtmlProperty.IsWriteable = true;

                propertiesMemoryStream.Write(bodyHtmlProperty.ToBytes(), 0, 16);
            }

            if (sensitivity != Sensitivity.None)
            {
                Property sensitivityProperty = new Property();
                sensitivityProperty.Tag = 0x00360003;
                sensitivityProperty.Type = PropertyType.Integer32;
                sensitivityProperty.Value = BitConverter.GetBytes(EnumUtil.ParseSensitivity(sensitivity));
                sensitivityProperty.IsReadable = true;
                sensitivityProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sensitivityProperty.ToBytes(), 0, 16);
            }

            if (lastVerbExecuted != LastVerbExecuted.None)
            {
                Property lastVerbExecutedProperty = new Property();
                lastVerbExecutedProperty.Tag = 0x10810003;
                lastVerbExecutedProperty.Type = PropertyType.Integer32;
                lastVerbExecutedProperty.Value = BitConverter.GetBytes(EnumUtil.ParseLastVerbExecuted(lastVerbExecuted));
                lastVerbExecutedProperty.IsReadable = true;
                lastVerbExecutedProperty.IsWriteable = true;

                propertiesMemoryStream.Write(lastVerbExecutedProperty.ToBytes(), 0, 16);
            }

            if (importance != Importance.None)
            {
                Property importanceProperty = new Property();
                importanceProperty.Tag = 0x00170003;
                importanceProperty.Type = PropertyType.Integer32;
                importanceProperty.Value = BitConverter.GetBytes(EnumUtil.ParseImportance(importance));
                importanceProperty.IsReadable = true;
                importanceProperty.IsWriteable = true;

                propertiesMemoryStream.Write(importanceProperty.ToBytes(), 0, 16);
            }

            if (priority != Priority.None)
            {
                Property priorityProperty = new Property();
                priorityProperty.Tag = 0x00260003;
                priorityProperty.Type = PropertyType.Integer32;
                priorityProperty.Value = BitConverter.GetBytes(EnumUtil.ParsePriority(priority));
                priorityProperty.IsReadable = true;
                priorityProperty.IsWriteable = true;

                propertiesMemoryStream.Write(priorityProperty.ToBytes(), 0, 16);
            }

            if (flagIcon != FlagIcon.None)
            {
                Property flagIconProperty = new Property();
                flagIconProperty.Tag = 0x10950003;
                flagIconProperty.Type = PropertyType.Integer32;
                flagIconProperty.Value = BitConverter.GetBytes(EnumUtil.ParseFlagIcon(flagIcon));
                flagIconProperty.IsReadable = true;
                flagIconProperty.IsWriteable = true;

                propertiesMemoryStream.Write(flagIconProperty.ToBytes(), 0, 16);
            }

            if (flagStatus != FlagStatus.None)
            {
                Property flagStatusProperty = new Property();
                flagStatusProperty.Tag = 0x10900003;
                flagStatusProperty.Type = PropertyType.Integer32;
                flagStatusProperty.Value = BitConverter.GetBytes(EnumUtil.ParseFlagStatus(flagStatus));
                flagStatusProperty.IsReadable = true;
                flagStatusProperty.IsWriteable = true;

                propertiesMemoryStream.Write(flagStatusProperty.ToBytes(), 0, 16);
            }

            if (objectType != ObjectType.None)
            {
                Property objectTypeProperty = new Property();
                objectTypeProperty.Tag = 0x0FFE0003;
                objectTypeProperty.Type = PropertyType.Integer32;
                objectTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseObjectType(objectType));
                objectTypeProperty.IsReadable = true;
                objectTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(objectTypeProperty.ToBytes(), 0, 16);
            }

            if (receivedRepresentingAddressType != null)
            {
                byte[] receivedRepresentingAddressTypeBuffer = encoding.GetBytes(receivedRepresentingAddressType);
                Independentsoft.IO.StructuredStorage.Stream receivedRepresentingAddressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0077" + stringTypeMask, receivedRepresentingAddressTypeBuffer);
                directoryEntries.Add(receivedRepresentingAddressTypeStream);

                Property receivedRepresentingAddressTypeProperty = new Property();
                receivedRepresentingAddressTypeProperty.Tag = (uint)0x0077 << 16 | stringTypeHexMask;
                receivedRepresentingAddressTypeProperty.Type = PropertyType.String8;
                receivedRepresentingAddressTypeProperty.Size = (uint)(receivedRepresentingAddressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                receivedRepresentingAddressTypeProperty.IsReadable = true;
                receivedRepresentingAddressTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedRepresentingAddressTypeProperty.ToBytes(), 0, 16);
            }

            if (receivedRepresentingEmailAddress != null)
            {
                byte[] receivedRepresentingEmailAddressBuffer = encoding.GetBytes(receivedRepresentingEmailAddress);
                Independentsoft.IO.StructuredStorage.Stream receivedRepresentingEmailAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0078" + stringTypeMask, receivedRepresentingEmailAddressBuffer);
                directoryEntries.Add(receivedRepresentingEmailAddressStream);

                Property receivedRepresentingEmailAddressProperty = new Property();
                receivedRepresentingEmailAddressProperty.Tag = (uint)0x0078 << 16 | stringTypeHexMask;
                receivedRepresentingEmailAddressProperty.Type = PropertyType.String8;
                receivedRepresentingEmailAddressProperty.Size = (uint)(receivedRepresentingEmailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                receivedRepresentingEmailAddressProperty.IsReadable = true;
                receivedRepresentingEmailAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedRepresentingEmailAddressProperty.ToBytes(), 0, 16);
            }

            if (receivedRepresentingEntryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream receivedRepresentingEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00430102", receivedRepresentingEntryId);
                directoryEntries.Add(receivedRepresentingEntryIdStream);

                Property receivedRepresentingEntryIdProperty = new Property();
                receivedRepresentingEntryIdProperty.Tag = 0x00430102;
                receivedRepresentingEntryIdProperty.Type = PropertyType.Binary;
                receivedRepresentingEntryIdProperty.Size = (uint)receivedRepresentingEntryId.Length;
                receivedRepresentingEntryIdProperty.IsReadable = true;
                receivedRepresentingEntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedRepresentingEntryIdProperty.ToBytes(), 0, 16);
            }

            if (receivedRepresentingName != null)
            {
                byte[] receivedRepresentingNameBuffer = encoding.GetBytes(receivedRepresentingName);
                Independentsoft.IO.StructuredStorage.Stream receivedRepresentingNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0044" + stringTypeMask, receivedRepresentingNameBuffer);
                directoryEntries.Add(receivedRepresentingNameStream);

                Property receivedRepresentingNameProperty = new Property();
                receivedRepresentingNameProperty.Tag = (uint)0x0044 << 16 | stringTypeHexMask;
                receivedRepresentingNameProperty.Type = PropertyType.String8;
                receivedRepresentingNameProperty.Size = (uint)(receivedRepresentingNameBuffer.Length + encoding.GetBytes("\0").Length);
                receivedRepresentingNameProperty.IsReadable = true;
                receivedRepresentingNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedRepresentingNameProperty.ToBytes(), 0, 16);
            }

            if (receivedRepresentingSearchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream receivedRepresentingSearchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00520102", receivedRepresentingSearchKey);
                directoryEntries.Add(receivedRepresentingSearchKeyStream);

                Property receivedRepresentingSearchKeyProperty = new Property();
                receivedRepresentingSearchKeyProperty.Tag = 0x00520102;
                receivedRepresentingSearchKeyProperty.Type = PropertyType.Binary;
                receivedRepresentingSearchKeyProperty.Size = (uint)receivedRepresentingSearchKey.Length;
                receivedRepresentingSearchKeyProperty.IsReadable = true;
                receivedRepresentingSearchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedRepresentingSearchKeyProperty.ToBytes(), 0, 16);
            }

            if (receivedByAddressType != null)
            {
                byte[] receivedByAddressTypeBuffer = encoding.GetBytes(receivedByAddressType);
                Independentsoft.IO.StructuredStorage.Stream receivedByAddressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0075" + stringTypeMask, receivedByAddressTypeBuffer);
                directoryEntries.Add(receivedByAddressTypeStream);

                Property receivedByAddressTypeProperty = new Property();
                receivedByAddressTypeProperty.Tag = (uint)0x0075 << 16 | stringTypeHexMask;
                receivedByAddressTypeProperty.Type = PropertyType.String8;
                receivedByAddressTypeProperty.Size = (uint)(receivedByAddressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                receivedByAddressTypeProperty.IsReadable = true;
                receivedByAddressTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedByAddressTypeProperty.ToBytes(), 0, 16);
            }

            if (receivedByEmailAddress != null)
            {
                byte[] receivedByEmailAddressBuffer = encoding.GetBytes(receivedByEmailAddress);
                Independentsoft.IO.StructuredStorage.Stream receivedByEmailAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0076" + stringTypeMask, receivedByEmailAddressBuffer);
                directoryEntries.Add(receivedByEmailAddressStream);

                Property receivedByEmailAddressProperty = new Property();
                receivedByEmailAddressProperty.Tag = (uint)0x0076 << 16 | stringTypeHexMask;
                receivedByEmailAddressProperty.Type = PropertyType.String8;
                receivedByEmailAddressProperty.Size = (uint)(receivedByEmailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                receivedByEmailAddressProperty.IsReadable = true;
                receivedByEmailAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedByEmailAddressProperty.ToBytes(), 0, 16);
            }

            if (receivedByEntryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream receivedByEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_003F0102", receivedByEntryId);
                directoryEntries.Add(receivedByEntryIdStream);

                Property receivedByEntryIdProperty = new Property();
                receivedByEntryIdProperty.Tag = 0x003F0102;
                receivedByEntryIdProperty.Type = PropertyType.Binary;
                receivedByEntryIdProperty.Size = (uint)receivedByEntryId.Length;
                receivedByEntryIdProperty.IsReadable = true;
                receivedByEntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedByEntryIdProperty.ToBytes(), 0, 16);
            }

            if (receivedByName != null)
            {
                byte[] receivedByNameBuffer = encoding.GetBytes(receivedByName);
                Independentsoft.IO.StructuredStorage.Stream receivedByNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0040" + stringTypeMask, receivedByNameBuffer);
                directoryEntries.Add(receivedByNameStream);

                Property receivedByNameProperty = new Property();
                receivedByNameProperty.Tag = (uint)0x0040 << 16 | stringTypeHexMask;
                receivedByNameProperty.Type = PropertyType.String8;
                receivedByNameProperty.Size = (uint)(receivedByNameBuffer.Length + encoding.GetBytes("\0").Length);
                receivedByNameProperty.IsReadable = true;
                receivedByNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedByNameProperty.ToBytes(), 0, 16);
            }

            if (receivedBySearchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream receivedBySearchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00510102", receivedBySearchKey);
                directoryEntries.Add(receivedBySearchKeyStream);

                Property receivedBySearchKeyProperty = new Property();
                receivedBySearchKeyProperty.Tag = 0x00510102;
                receivedBySearchKeyProperty.Type = PropertyType.Binary;
                receivedBySearchKeyProperty.Size = (uint)receivedBySearchKey.Length;
                receivedBySearchKeyProperty.IsReadable = true;
                receivedBySearchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(receivedBySearchKeyProperty.ToBytes(), 0, 16);
            }

            if (senderAddressType != null)
            {
                byte[] senderAddressTypeBuffer = encoding.GetBytes(senderAddressType);
                Independentsoft.IO.StructuredStorage.Stream senderAddressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0C1E" + stringTypeMask, senderAddressTypeBuffer);
                directoryEntries.Add(senderAddressTypeStream);

                Property senderAddressTypeProperty = new Property();
                senderAddressTypeProperty.Tag = (uint)0x0C1E << 16 | stringTypeHexMask;
                senderAddressTypeProperty.Type = PropertyType.String8;
                senderAddressTypeProperty.Size = (uint)(senderAddressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                senderAddressTypeProperty.IsReadable = true;
                senderAddressTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderAddressTypeProperty.ToBytes(), 0, 16);
            }

            if (senderEmailAddress != null)
            {
                byte[] senderEmailAddressBuffer = encoding.GetBytes(senderEmailAddress);
                Independentsoft.IO.StructuredStorage.Stream senderEmailAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0C1F" + stringTypeMask, senderEmailAddressBuffer);
                directoryEntries.Add(senderEmailAddressStream);

                Property senderEmailAddressProperty = new Property();
                senderEmailAddressProperty.Tag = (uint)0x0C1F << 16 | stringTypeHexMask;
                senderEmailAddressProperty.Type = PropertyType.String8;
                senderEmailAddressProperty.Size = (uint)(senderEmailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                senderEmailAddressProperty.IsReadable = true;
                senderEmailAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderEmailAddressProperty.ToBytes(), 0, 16);
            }

            if (senderSmtpAddress != null)
            {
                byte[] senderSmtpAddressBuffer = encoding.GetBytes(senderSmtpAddress);
                Independentsoft.IO.StructuredStorage.Stream senderSmtpAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_5D01" + stringTypeMask, senderSmtpAddressBuffer);
                directoryEntries.Add(senderSmtpAddressStream);

                Property senderSmtpAddressProperty = new Property();
                senderSmtpAddressProperty.Tag = (uint)0x5D01 << 16 | stringTypeHexMask;
                senderSmtpAddressProperty.Type = PropertyType.String8;
                senderSmtpAddressProperty.Size = (uint)(senderSmtpAddressBuffer.Length + encoding.GetBytes("\0").Length);
                senderSmtpAddressProperty.IsReadable = true;
                senderSmtpAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderSmtpAddressProperty.ToBytes(), 0, 16);
            }

            if (senderEntryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream senderEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0C190102", senderEntryId);
                directoryEntries.Add(senderEntryIdStream);

                Property senderEntryIdProperty = new Property();
                senderEntryIdProperty.Tag = 0x0C190102;
                senderEntryIdProperty.Type = PropertyType.Binary;
                senderEntryIdProperty.Size = (uint)senderEntryId.Length;
                senderEntryIdProperty.IsReadable = true;
                senderEntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderEntryIdProperty.ToBytes(), 0, 16);
            }

            if (senderName != null)
            {
                byte[] senderNameBuffer = encoding.GetBytes(senderName);
                Independentsoft.IO.StructuredStorage.Stream senderNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0C1A" + stringTypeMask, senderNameBuffer);
                directoryEntries.Add(senderNameStream);

                Property senderNameProperty = new Property();
                senderNameProperty.Tag = (uint)0x0C1A << 16 | stringTypeHexMask;
                senderNameProperty.Type = PropertyType.String8;
                senderNameProperty.Size = (uint)(senderNameBuffer.Length + encoding.GetBytes("\0").Length);
                senderNameProperty.IsReadable = true;
                senderNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderNameProperty.ToBytes(), 0, 16);
            }

            if (senderSearchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream senderSearchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0C1D0102", senderSearchKey);
                directoryEntries.Add(senderSearchKeyStream);

                Property senderSearchKeyProperty = new Property();
                senderSearchKeyProperty.Tag = 0x0C1D0102;
                senderSearchKeyProperty.Type = PropertyType.Binary;
                senderSearchKeyProperty.Size = (uint)senderSearchKey.Length;
                senderSearchKeyProperty.IsReadable = true;
                senderSearchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(senderSearchKeyProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingAddressType != null)
            {
                byte[] sentRepresentingAddressTypeBuffer = encoding.GetBytes(sentRepresentingAddressType);
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingAddressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0064" + stringTypeMask, sentRepresentingAddressTypeBuffer);
                directoryEntries.Add(sentRepresentingAddressTypeStream);

                Property sentRepresentingAddressTypeProperty = new Property();
                sentRepresentingAddressTypeProperty.Tag = (uint)0x0064 << 16 | stringTypeHexMask;
                sentRepresentingAddressTypeProperty.Type = PropertyType.String8;
                sentRepresentingAddressTypeProperty.Size = (uint)(sentRepresentingAddressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                sentRepresentingAddressTypeProperty.IsReadable = true;
                sentRepresentingAddressTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingAddressTypeProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingEmailAddress != null)
            {
                byte[] sentRepresentingEmailAddressBuffer = encoding.GetBytes(sentRepresentingEmailAddress);
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingEmailAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0065" + stringTypeMask, sentRepresentingEmailAddressBuffer);
                directoryEntries.Add(sentRepresentingEmailAddressStream);

                Property sentRepresentingEmailAddressProperty = new Property();
                sentRepresentingEmailAddressProperty.Tag = (uint)0x0065 << 16 | stringTypeHexMask;
                sentRepresentingEmailAddressProperty.Type = PropertyType.String8;
                sentRepresentingEmailAddressProperty.Size = (uint)(sentRepresentingEmailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                sentRepresentingEmailAddressProperty.IsReadable = true;
                sentRepresentingEmailAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingEmailAddressProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingSmtpAddress != null)
            {
                byte[] sentRepresentingSmtpAddressBuffer = encoding.GetBytes(sentRepresentingSmtpAddress);
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingSmtpStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_5D02" + stringTypeMask, sentRepresentingSmtpAddressBuffer);
                directoryEntries.Add(sentRepresentingSmtpStream);

                Property sentRepresentingSmtpAddressProperty = new Property();
                sentRepresentingSmtpAddressProperty.Tag = (uint)0x5D02 << 16 | stringTypeHexMask;
                sentRepresentingSmtpAddressProperty.Type = PropertyType.String8;
                sentRepresentingSmtpAddressProperty.Size = (uint)(sentRepresentingSmtpAddressBuffer.Length + encoding.GetBytes("\0").Length);
                sentRepresentingSmtpAddressProperty.IsReadable = true;
                sentRepresentingSmtpAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingSmtpAddressProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingEntryId != null)
            {
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_00410102", sentRepresentingEntryId);
                directoryEntries.Add(sentRepresentingEntryIdStream);

                Property sentRepresentingEntryIdProperty = new Property();
                sentRepresentingEntryIdProperty.Tag = 0x00410102;
                sentRepresentingEntryIdProperty.Type = PropertyType.Binary;
                sentRepresentingEntryIdProperty.Size = (uint)sentRepresentingEntryId.Length;
                sentRepresentingEntryIdProperty.IsReadable = true;
                sentRepresentingEntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingEntryIdProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingName != null)
            {
                byte[] sentRepresentingNameBuffer = encoding.GetBytes(sentRepresentingName);
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0042" + stringTypeMask, sentRepresentingNameBuffer);
                directoryEntries.Add(sentRepresentingNameStream);

                Property sentRepresentingNameProperty = new Property();
                sentRepresentingNameProperty.Tag = (uint)0x0042 << 16 | stringTypeHexMask;
                sentRepresentingNameProperty.Type = PropertyType.String8;
                sentRepresentingNameProperty.Size = (uint)(sentRepresentingNameBuffer.Length + encoding.GetBytes("\0").Length);
                sentRepresentingNameProperty.IsReadable = true;
                sentRepresentingNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingNameProperty.ToBytes(), 0, 16);
            }

            if (sentRepresentingSearchKey != null)
            {
                Independentsoft.IO.StructuredStorage.Stream sentRepresentingSearchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_003B0102", sentRepresentingSearchKey);
                directoryEntries.Add(sentRepresentingSearchKeyStream);

                Property sentRepresentingSearchKeyProperty = new Property();
                sentRepresentingSearchKeyProperty.Tag = 0x003B0102;
                sentRepresentingSearchKeyProperty.Type = PropertyType.Binary;
                sentRepresentingSearchKeyProperty.Size = (uint)sentRepresentingSearchKey.Length;
                sentRepresentingSearchKeyProperty.IsReadable = true;
                sentRepresentingSearchKeyProperty.IsWriteable = true;

                propertiesMemoryStream.Write(sentRepresentingSearchKeyProperty.ToBytes(), 0, 16);
            }

            if (transportMessageHeaders != null)
            {
                byte[] transportMessageHeadersBuffer = encoding.GetBytes(transportMessageHeaders);
                Independentsoft.IO.StructuredStorage.Stream transportMessageHeadersStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_007D" + stringTypeMask, transportMessageHeadersBuffer);
                directoryEntries.Add(transportMessageHeadersStream);

                Property transportMessageHeadersProperty = new Property();
                transportMessageHeadersProperty.Tag = (uint)0x007D << 16 | stringTypeHexMask;
                transportMessageHeadersProperty.Type = PropertyType.String8;
                transportMessageHeadersProperty.Size = (uint)(transportMessageHeadersBuffer.Length + encoding.GetBytes("\0").Length);
                transportMessageHeadersProperty.IsReadable = true;
                transportMessageHeadersProperty.IsWriteable = true;

                propertiesMemoryStream.Write(transportMessageHeadersProperty.ToBytes(), 0, 16);
            }

            if (outlookVersion != null)
            {
                NamedProperty outlookVersionNamedProperty = new NamedProperty();
                outlookVersionNamedProperty.Id = 0x8554;
                outlookVersionNamedProperty.Guid = StandardPropertySet.Common;
                outlookVersionNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, outlookVersionNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(outlookVersionNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] outlookVersionBuffer = encoding.GetBytes(outlookVersion);
                Independentsoft.IO.StructuredStorage.Stream outlookVersionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, outlookVersionBuffer);
                directoryEntries.Add(outlookVersionStream);

                Property outlookVersionProperty = new Property();
                outlookVersionProperty.Tag = propertyTag;
                outlookVersionProperty.Type = PropertyType.String8;
                outlookVersionProperty.Size = (uint)(outlookVersionBuffer.Length + encoding.GetBytes("\0").Length);
                outlookVersionProperty.IsReadable = true;
                outlookVersionProperty.IsWriteable = true;

                propertiesMemoryStream.Write(outlookVersionProperty.ToBytes(), 0, 16);
            }

            if (outlookInternalVersion > 0)
            {
                NamedProperty outlookInternalVersionNamedProperty = new NamedProperty();
                outlookInternalVersionNamedProperty.Id = 0x8552;
                outlookInternalVersionNamedProperty.Guid = StandardPropertySet.Common;
                outlookInternalVersionNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, outlookInternalVersionNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(outlookInternalVersionNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property outlookInternalVersionProperty = new Property();
                outlookInternalVersionProperty.Tag = propertyTag;
                outlookInternalVersionProperty.Type = PropertyType.Integer32;
                outlookInternalVersionProperty.Value = BitConverter.GetBytes(outlookInternalVersion);
                outlookInternalVersionProperty.IsReadable = true;
                outlookInternalVersionProperty.IsWriteable = true;

                propertiesMemoryStream.Write(outlookInternalVersionProperty.ToBytes(), 0, 16);
            }

            if (commonStartTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty commonStartTimeNamedProperty = new NamedProperty();
                commonStartTimeNamedProperty.Id = 0x8516;
                commonStartTimeNamedProperty.Guid = StandardPropertySet.Common;
                commonStartTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, commonStartTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(commonStartTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = commonStartTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property commonStartTimeProperty = new Property();
                commonStartTimeProperty.Tag = propertyTag;
                commonStartTimeProperty.Type = PropertyType.Time;
                commonStartTimeProperty.Value = ticksBytes;
                commonStartTimeProperty.IsReadable = true;
                commonStartTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(commonStartTimeProperty.ToBytes(), 0, 16);
            }

            if (commonEndTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty commonEndTimeNamedProperty = new NamedProperty();
                commonEndTimeNamedProperty.Id = 0x8517;
                commonEndTimeNamedProperty.Guid = StandardPropertySet.Common;
                commonEndTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, commonEndTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(commonEndTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = commonEndTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property commonEndTimeProperty = new Property();
                commonEndTimeProperty.Tag = propertyTag;
                commonEndTimeProperty.Type = PropertyType.Time;
                commonEndTimeProperty.Value = ticksBytes;
                commonEndTimeProperty.IsReadable = true;
                commonEndTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(commonEndTimeProperty.ToBytes(), 0, 16);
            }

            if (flagDueBy.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty flagDueByNamedProperty = new NamedProperty();
                flagDueByNamedProperty.Id = 0x8560;
                flagDueByNamedProperty.Guid = StandardPropertySet.Common;
                flagDueByNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, flagDueByNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(flagDueByNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = flagDueBy.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property flagDueByProperty = new Property();
                flagDueByProperty.Tag = propertyTag;
                flagDueByProperty.Type = PropertyType.Time;
                flagDueByProperty.Value = ticksBytes;
                flagDueByProperty.IsReadable = true;
                flagDueByProperty.IsWriteable = true;

                propertiesMemoryStream.Write(flagDueByProperty.ToBytes(), 0, 16);
            }

            if (companies.Count > 0)
            {
                NamedProperty companiesNamedProperty = new NamedProperty();
                companiesNamedProperty.Id = 0x8539;
                companiesNamedProperty.Guid = StandardPropertySet.Common;
                companiesNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, companiesNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(companiesNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | multiValueStringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                MemoryStream companiesLengthMemoryStream = new MemoryStream();

                for (int i = 0; i < companies.Count; i++)
                {
                    byte[] companiesBuffer = encoding.GetBytes(companies[i] + "\0");
                    uint size = (uint)companiesBuffer.Length;
                    byte[] sizeBuffer = BitConverter.GetBytes(size);

                    companiesLengthMemoryStream.Write(sizeBuffer, 0, 4);

                    string streamName = "__substg1.0_" + propertyTagString + "-" + String.Format("{0:X8}", i);

                    Independentsoft.IO.StructuredStorage.Stream companiesStream = new Independentsoft.IO.StructuredStorage.Stream(streamName, companiesBuffer);
                    directoryEntries.Add(companiesStream);
                }

                byte[] companiesLengthStreamBuffer = companiesLengthMemoryStream.ToArray();

                Independentsoft.IO.StructuredStorage.Stream companiesLengthStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, companiesLengthStreamBuffer);
                directoryEntries.Add(companiesLengthStream);

                Property companiesLengthProperty = new Property();
                companiesLengthProperty.Tag = propertyTag;
                companiesLengthProperty.Type = PropertyType.MultipleString8;
                companiesLengthProperty.Size = (uint)companiesLengthStreamBuffer.Length;
                companiesLengthProperty.IsReadable = true;
                companiesLengthProperty.IsWriteable = true;

                propertiesMemoryStream.Write(companiesLengthProperty.ToBytes(), 0, 16);
            }

            if (contactNames.Count > 0)
            {
                NamedProperty contactNamesNamedProperty = new NamedProperty();
                contactNamesNamedProperty.Id = 0x853A;
                contactNamesNamedProperty.Guid = StandardPropertySet.Common;
                contactNamesNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, contactNamesNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(contactNamesNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | multiValueStringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                MemoryStream contactNamesLengthMemoryStream = new MemoryStream();

                for (int i = 0; i < contactNames.Count; i++)
                {
                    byte[] contactNamesBuffer = encoding.GetBytes(contactNames[i] + "\0");
                    uint size = (uint)contactNamesBuffer.Length;
                    byte[] sizeBuffer = BitConverter.GetBytes(size);

                    contactNamesLengthMemoryStream.Write(sizeBuffer, 0, 4);

                    string streamName = "__substg1.0_" + propertyTagString + "-" + String.Format("{0:X8}", i);

                    Independentsoft.IO.StructuredStorage.Stream contactNamesStream = new Independentsoft.IO.StructuredStorage.Stream(streamName, contactNamesBuffer);
                    directoryEntries.Add(contactNamesStream);
                }

                byte[] contactNamesLengthStreamBuffer = contactNamesLengthMemoryStream.ToArray();

                Independentsoft.IO.StructuredStorage.Stream contactNamesLengthStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, contactNamesLengthStreamBuffer);
                directoryEntries.Add(contactNamesLengthStream);

                Property contactNamesLengthProperty = new Property();
                contactNamesLengthProperty.Tag = propertyTag;
                contactNamesLengthProperty.Type = PropertyType.MultipleString8;
                contactNamesLengthProperty.Size = (uint)contactNamesLengthStreamBuffer.Length;
                contactNamesLengthProperty.IsReadable = true;
                contactNamesLengthProperty.IsWriteable = true;

                propertiesMemoryStream.Write(contactNamesLengthProperty.ToBytes(), 0, 16);
            }

            if (keywords.Count > 0)
            {
                NamedProperty keywordsNamedProperty = new NamedProperty();
                keywordsNamedProperty.Name = "Keywords";
                keywordsNamedProperty.Guid = StandardPropertySet.PublicStrings;
                keywordsNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, keywordsNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(keywordsNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | multiValueStringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                MemoryStream keywordsLengthMemoryStream = new MemoryStream();

                for (int i = 0; i < keywords.Count; i++)
                {
                    byte[] keywordsBuffer = encoding.GetBytes(keywords[i] + "\0");
                    uint size = (uint)keywordsBuffer.Length;
                    byte[] sizeBuffer = BitConverter.GetBytes(size);

                    keywordsLengthMemoryStream.Write(sizeBuffer, 0, 4);

                    string streamName = "__substg1.0_" + propertyTagString + "-" + String.Format("{0:X8}", i);

                    Independentsoft.IO.StructuredStorage.Stream keywordsStream = new Independentsoft.IO.StructuredStorage.Stream(streamName, keywordsBuffer);
                    directoryEntries.Add(keywordsStream);
                }

                byte[] keywordsLengthStreamBuffer = keywordsLengthMemoryStream.ToArray();

                Independentsoft.IO.StructuredStorage.Stream keywordsLengthStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, keywordsLengthStreamBuffer);
                directoryEntries.Add(keywordsLengthStream);

                Property keywordsLengthProperty = new Property();
                keywordsLengthProperty.Tag = propertyTag;
                keywordsLengthProperty.Type = PropertyType.MultipleString8;
                keywordsLengthProperty.Size = (uint)keywordsLengthStreamBuffer.Length;
                keywordsLengthProperty.IsReadable = true;
                keywordsLengthProperty.IsWriteable = true;

                propertiesMemoryStream.Write(keywordsLengthProperty.ToBytes(), 0, 16);
            }

            if (billingInformation != null)
            {
                NamedProperty billingInformationNamedProperty = new NamedProperty();
                billingInformationNamedProperty.Id = 0x8535;
                billingInformationNamedProperty.Guid = StandardPropertySet.Common;
                billingInformationNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, billingInformationNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(billingInformationNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] billingInformationBuffer = encoding.GetBytes(billingInformation);
                Independentsoft.IO.StructuredStorage.Stream billingInformationStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, billingInformationBuffer);
                directoryEntries.Add(billingInformationStream);

                Property billingInformationProperty = new Property();
                billingInformationProperty.Tag = propertyTag;
                billingInformationProperty.Type = PropertyType.String8;
                billingInformationProperty.Size = (uint)(billingInformationBuffer.Length + encoding.GetBytes("\0").Length);
                billingInformationProperty.IsReadable = true;
                billingInformationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(billingInformationProperty.ToBytes(), 0, 16);
            }

            if (mileage != null)
            {
                NamedProperty mileageNamedProperty = new NamedProperty();
                mileageNamedProperty.Id = 0x8534;
                mileageNamedProperty.Guid = StandardPropertySet.Common;
                mileageNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, mileageNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(mileageNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] mileageBuffer = encoding.GetBytes(mileage);
                Independentsoft.IO.StructuredStorage.Stream mileageStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, mileageBuffer);
                directoryEntries.Add(mileageStream);

                Property mileageProperty = new Property();
                mileageProperty.Tag = propertyTag;
                mileageProperty.Type = PropertyType.String8;
                mileageProperty.Size = (uint)(mileageBuffer.Length + encoding.GetBytes("\0").Length);
                mileageProperty.IsReadable = true;
                mileageProperty.IsWriteable = true;

                propertiesMemoryStream.Write(mileageProperty.ToBytes(), 0, 16);
            }

            if (internetAccountName != null)
            {
                NamedProperty internetAccountNameNamedProperty = new NamedProperty();
                internetAccountNameNamedProperty.Id = 0x8580;
                internetAccountNameNamedProperty.Guid = StandardPropertySet.Common;
                internetAccountNameNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, internetAccountNameNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(internetAccountNameNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] internetAccountNameBuffer = encoding.GetBytes(internetAccountName);
                Independentsoft.IO.StructuredStorage.Stream internetAccountNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, internetAccountNameBuffer);
                directoryEntries.Add(internetAccountNameStream);

                Property internetAccountNameProperty = new Property();
                internetAccountNameProperty.Tag = propertyTag;
                internetAccountNameProperty.Type = PropertyType.String8;
                internetAccountNameProperty.Size = (uint)(internetAccountNameBuffer.Length + encoding.GetBytes("\0").Length);
                internetAccountNameProperty.IsReadable = true;
                internetAccountNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(internetAccountNameProperty.ToBytes(), 0, 16);
            }

            if (reminderSoundFile != null)
            {
                NamedProperty reminderSoundFileNamedProperty = new NamedProperty();
                reminderSoundFileNamedProperty.Id = 0x851F;
                reminderSoundFileNamedProperty.Guid = StandardPropertySet.Common;
                reminderSoundFileNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, reminderSoundFileNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(reminderSoundFileNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] reminderSoundFileBuffer = encoding.GetBytes(reminderSoundFile);
                Independentsoft.IO.StructuredStorage.Stream reminderSoundFileStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, reminderSoundFileBuffer);
                directoryEntries.Add(reminderSoundFileStream);

                Property reminderSoundFileProperty = new Property();
                reminderSoundFileProperty.Tag = propertyTag;
                reminderSoundFileProperty.Type = PropertyType.String8;
                reminderSoundFileProperty.Size = (uint)(reminderSoundFileBuffer.Length + encoding.GetBytes("\0").Length);
                reminderSoundFileProperty.IsReadable = true;
                reminderSoundFileProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reminderSoundFileProperty.ToBytes(), 0, 16);
            }

            if (isPrivate)
            {
                NamedProperty isPrivateNamedProperty = new NamedProperty();
                isPrivateNamedProperty.Id = 0x8506;
                isPrivateNamedProperty.Guid = StandardPropertySet.Common;
                isPrivateNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isPrivateNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isPrivateNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isPrivateProperty = new Property();
                isPrivateProperty.Tag = propertyTag;
                isPrivateProperty.Type = PropertyType.Boolean;
                isPrivateProperty.Value = BitConverter.GetBytes(1);
                isPrivateProperty.IsReadable = true;
                isPrivateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isPrivateProperty.ToBytes(), 0, 16);
            }

            if (reminderOverrideDefault)
            {
                NamedProperty reminderOverrideDefaultNamedProperty = new NamedProperty();
                reminderOverrideDefaultNamedProperty.Id = 0x851C;
                reminderOverrideDefaultNamedProperty.Guid = StandardPropertySet.Common;
                reminderOverrideDefaultNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, reminderOverrideDefaultNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(reminderOverrideDefaultNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property reminderOverrideDefaultProperty = new Property();
                reminderOverrideDefaultProperty.Tag = propertyTag;
                reminderOverrideDefaultProperty.Type = PropertyType.Boolean;
                reminderOverrideDefaultProperty.Value = BitConverter.GetBytes(1);
                reminderOverrideDefaultProperty.IsReadable = true;
                reminderOverrideDefaultProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reminderOverrideDefaultProperty.ToBytes(), 0, 16);
            }

            if (reminderPlaySound)
            {
                NamedProperty reminderPlaySoundNamedProperty = new NamedProperty();
                reminderPlaySoundNamedProperty.Id = 0x851E;
                reminderPlaySoundNamedProperty.Guid = StandardPropertySet.Common;
                reminderPlaySoundNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, reminderPlaySoundNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(reminderPlaySoundNamedProperty); 
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property reminderPlaySoundProperty = new Property();
                reminderPlaySoundProperty.Tag = propertyTag;
                reminderPlaySoundProperty.Type = PropertyType.Boolean;
                reminderPlaySoundProperty.Value = BitConverter.GetBytes(1);
                reminderPlaySoundProperty.IsReadable = true;
                reminderPlaySoundProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reminderPlaySoundProperty.ToBytes(), 0, 16);
            }

            if (appointmentStartTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty appointmentStartTimeNamedProperty = new NamedProperty();
                appointmentStartTimeNamedProperty.Id = 0x820D;
                appointmentStartTimeNamedProperty.Guid = StandardPropertySet.Appointment;
                appointmentStartTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, appointmentStartTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(appointmentStartTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = appointmentStartTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property appointmentStartTimeProperty = new Property();
                appointmentStartTimeProperty.Tag = propertyTag;
                appointmentStartTimeProperty.Type = PropertyType.Time;
                appointmentStartTimeProperty.Value = ticksBytes;
                appointmentStartTimeProperty.IsReadable = true;
                appointmentStartTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(appointmentStartTimeProperty.ToBytes(), 0, 16);
            }

            if (appointmentEndTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty appointmentEndTimeNamedProperty = new NamedProperty();
                appointmentEndTimeNamedProperty.Id = 0x820E;
                appointmentEndTimeNamedProperty.Guid = StandardPropertySet.Appointment;
                appointmentEndTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, appointmentEndTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(appointmentEndTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = appointmentEndTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property appointmentEndTimeProperty = new Property();
                appointmentEndTimeProperty.Tag = propertyTag;
                appointmentEndTimeProperty.Type = PropertyType.Time;
                appointmentEndTimeProperty.Value = ticksBytes;
                appointmentEndTimeProperty.IsReadable = true;
                appointmentEndTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(appointmentEndTimeProperty.ToBytes(), 0, 16);
            }

            if (location != null)
            {
                NamedProperty locationNamedProperty = new NamedProperty();
                locationNamedProperty.Id = 0x8208;
                locationNamedProperty.Guid = StandardPropertySet.Appointment;
                locationNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, locationNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(locationNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] locationBuffer = encoding.GetBytes(location);
                Independentsoft.IO.StructuredStorage.Stream locationStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, locationBuffer);
                directoryEntries.Add(locationStream);

                Property locationProperty = new Property();
                locationProperty.Tag = propertyTag;
                locationProperty.Type = PropertyType.String8;
                locationProperty.Size = (uint)(locationBuffer.Length + encoding.GetBytes("\0").Length);
                locationProperty.IsReadable = true;
                locationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(locationProperty.ToBytes(), 0, 16);
            }

            if (appointmentMessageClass != null)
            {
                NamedProperty appointmentMessageClassNamedProperty = new NamedProperty();
                appointmentMessageClassNamedProperty.Id = 0x24;
                appointmentMessageClassNamedProperty.Guid = new byte[16] { 144, 218, 216, 110, 11, 69, 27, 16, 152, 218, 0, 170, 0, 63, 19, 5 };
                appointmentMessageClassNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, appointmentMessageClassNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(appointmentMessageClassNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] appointmentMessageClassBuffer = encoding.GetBytes(appointmentMessageClass);
                Independentsoft.IO.StructuredStorage.Stream appointmentMessageClassStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, appointmentMessageClassBuffer);
                directoryEntries.Add(appointmentMessageClassStream);

                Property appointmentMessageClassProperty = new Property();
                appointmentMessageClassProperty.Tag = propertyTag;
                appointmentMessageClassProperty.Type = PropertyType.String8;
                appointmentMessageClassProperty.Size = (uint)(appointmentMessageClassBuffer.Length + encoding.GetBytes("\0").Length);
                appointmentMessageClassProperty.IsReadable = true;
                appointmentMessageClassProperty.IsWriteable = true;

                propertiesMemoryStream.Write(appointmentMessageClassProperty.ToBytes(), 0, 16);
            }

            if (timeZone != null)
            {
                NamedProperty timeZoneNamedProperty = new NamedProperty();
                timeZoneNamedProperty.Id = 0x8234;
                timeZoneNamedProperty.Guid = StandardPropertySet.Appointment;
                timeZoneNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, timeZoneNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(timeZoneNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] timeZoneBuffer = encoding.GetBytes(timeZone);
                Independentsoft.IO.StructuredStorage.Stream timeZoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, timeZoneBuffer);
                directoryEntries.Add(timeZoneStream);

                Property timeZoneProperty = new Property();
                timeZoneProperty.Tag = propertyTag;
                timeZoneProperty.Type = PropertyType.String8;
                timeZoneProperty.Size = (uint)(timeZoneBuffer.Length + encoding.GetBytes("\0").Length);
                timeZoneProperty.IsReadable = true;
                timeZoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(timeZoneProperty.ToBytes(), 0, 16);
            }

            if (recurrencePatternDescription != null)
            {
                NamedProperty recurrencePatternDescriptionNamedProperty = new NamedProperty();
                recurrencePatternDescriptionNamedProperty.Id = 0x8232;
                recurrencePatternDescriptionNamedProperty.Guid = StandardPropertySet.Appointment;
                recurrencePatternDescriptionNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, recurrencePatternDescriptionNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(recurrencePatternDescriptionNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] recurrencePatternDescriptionBuffer = encoding.GetBytes(recurrencePatternDescription);
                Independentsoft.IO.StructuredStorage.Stream recurrencePatternDescriptionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, recurrencePatternDescriptionBuffer);
                directoryEntries.Add(recurrencePatternDescriptionStream);

                Property recurrencePatternDescriptionProperty = new Property();
                recurrencePatternDescriptionProperty.Tag = propertyTag;
                recurrencePatternDescriptionProperty.Type = PropertyType.String8;
                recurrencePatternDescriptionProperty.Size = (uint)(recurrencePatternDescriptionBuffer.Length + encoding.GetBytes("\0").Length);
                recurrencePatternDescriptionProperty.IsReadable = true;
                recurrencePatternDescriptionProperty.IsWriteable = true;

                propertiesMemoryStream.Write(recurrencePatternDescriptionProperty.ToBytes(), 0, 16);
            }

            if (busyStatus != BusyStatus.None)
            {
                NamedProperty busyStatusNamedProperty = new NamedProperty();
                busyStatusNamedProperty.Id = 0x8205;
                busyStatusNamedProperty.Guid = StandardPropertySet.Appointment;
                busyStatusNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, busyStatusNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(busyStatusNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property busyStatusProperty = new Property();
                busyStatusProperty.Tag = propertyTag;
                busyStatusProperty.Type = PropertyType.Integer32;
                busyStatusProperty.Value = BitConverter.GetBytes(EnumUtil.ParseBusyStatus(busyStatus));
                busyStatusProperty.IsReadable = true;
                busyStatusProperty.IsWriteable = true;

                propertiesMemoryStream.Write(busyStatusProperty.ToBytes(), 0, 16);
            }

            if (meetingStatus != MeetingStatus.None)
            {
                NamedProperty meetingStatusNamedProperty = new NamedProperty();
                meetingStatusNamedProperty.Id = 0x8217;
                meetingStatusNamedProperty.Guid = StandardPropertySet.Appointment;
                meetingStatusNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, meetingStatusNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(meetingStatusNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property meetingStatusProperty = new Property();
                meetingStatusProperty.Tag = propertyTag;
                meetingStatusProperty.Type = PropertyType.Integer32;
                meetingStatusProperty.Value = BitConverter.GetBytes(EnumUtil.ParseMeetingStatus(meetingStatus));
                meetingStatusProperty.IsReadable = true;
                meetingStatusProperty.IsWriteable = true;

                propertiesMemoryStream.Write(meetingStatusProperty.ToBytes(), 0, 16);
            }

            if (responseStatus != ResponseStatus.None)
            {
                NamedProperty responseStatusNamedProperty = new NamedProperty();
                responseStatusNamedProperty.Id = 0x8218;
                responseStatusNamedProperty.Guid = StandardPropertySet.Appointment;
                responseStatusNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, responseStatusNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(responseStatusNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property responseStatusProperty = new Property();
                responseStatusProperty.Tag = propertyTag;
                responseStatusProperty.Type = PropertyType.Integer32;
                responseStatusProperty.Value = BitConverter.GetBytes(EnumUtil.ParseResponseStatus(responseStatus));
                responseStatusProperty.IsReadable = true;
                responseStatusProperty.IsWriteable = true;

                propertiesMemoryStream.Write(responseStatusProperty.ToBytes(), 0, 16);
            }

            if (recurrenceType != RecurrenceType.None)
            {
                NamedProperty recurrenceTypeNamedProperty = new NamedProperty();
                recurrenceTypeNamedProperty.Id = 0x8231;
                recurrenceTypeNamedProperty.Guid = StandardPropertySet.Appointment;
                recurrenceTypeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, recurrenceTypeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(recurrenceTypeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property recurrenceTypeProperty = new Property();
                recurrenceTypeProperty.Tag = propertyTag;
                recurrenceTypeProperty.Type = PropertyType.Integer32;
                recurrenceTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseRecurrenceType(recurrenceType));
                recurrenceTypeProperty.IsReadable = true;
                recurrenceTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(recurrenceTypeProperty.ToBytes(), 0, 16);
            }

            if (guid != null)
            {
                NamedProperty guidNamedProperty = new NamedProperty();
                guidNamedProperty.Id = 0x3;
                guidNamedProperty.Guid = new byte[16] { 144, 218, 216, 110, 11, 69, 27, 16, 152, 218, 0, 170, 0, 63, 19, 5 };
                guidNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, guidNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(guidNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0102;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                Independentsoft.IO.StructuredStorage.Stream guidStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, guid);
                directoryEntries.Add(guidStream);

                Property guidProperty = new Property();
                guidProperty.Tag = propertyTag;
                guidProperty.Type = PropertyType.Integer32;
                guidProperty.Size = (uint)guid.Length;
                guidProperty.IsReadable = true;
                guidProperty.IsWriteable = true;

                propertiesMemoryStream.Write(guidProperty.ToBytes(), 0, 16);
            }

            if (label > -1)
            {
                NamedProperty labelNamedProperty = new NamedProperty();
                labelNamedProperty.Id = 0x8214;
                labelNamedProperty.Guid = StandardPropertySet.Appointment;
                labelNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, labelNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(labelNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property labelProperty = new Property();
                labelProperty.Tag = propertyTag;
                labelProperty.Type = PropertyType.Integer32;
                labelProperty.Value = BitConverter.GetBytes(label);
                labelProperty.IsReadable = true;
                labelProperty.IsWriteable = true;

                propertiesMemoryStream.Write(labelProperty.ToBytes(), 0, 16);
            }

            if (duration > 0)
            {
                NamedProperty durationNamedProperty = new NamedProperty();
                durationNamedProperty.Id = 0x8213;
                durationNamedProperty.Guid = StandardPropertySet.Appointment;
                durationNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, durationNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(durationNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property durationProperty = new Property();
                durationProperty.Tag = propertyTag;
                durationProperty.Type = PropertyType.Integer32;
                durationProperty.Value = BitConverter.GetBytes(duration);
                durationProperty.IsReadable = true;
                durationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(durationProperty.ToBytes(), 0, 16);
            }


            if (owner != null)
            {
                NamedProperty ownerNamedProperty = new NamedProperty();
                ownerNamedProperty.Id = 0x811F;
                ownerNamedProperty.Guid = StandardPropertySet.Task;
                ownerNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, ownerNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(ownerNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] ownerBuffer = encoding.GetBytes(owner);
                Independentsoft.IO.StructuredStorage.Stream ownerStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, ownerBuffer);
                directoryEntries.Add(ownerStream);

                Property ownerProperty = new Property();
                ownerProperty.Tag = propertyTag;
                ownerProperty.Type = PropertyType.String8;
                ownerProperty.Size = (uint)(ownerBuffer.Length + encoding.GetBytes("\0").Length);
                ownerProperty.IsReadable = true;
                ownerProperty.IsWriteable = true;

                propertiesMemoryStream.Write(ownerProperty.ToBytes(), 0, 16);
            }

            if (delegator != null)
            {
                NamedProperty delegatorNamedProperty = new NamedProperty();
                delegatorNamedProperty.Id = 0x8121;
                delegatorNamedProperty.Guid = StandardPropertySet.Task;
                delegatorNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, delegatorNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(delegatorNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] delegatorBuffer = encoding.GetBytes(delegator);
                Independentsoft.IO.StructuredStorage.Stream delegatorStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, delegatorBuffer);
                directoryEntries.Add(delegatorStream);

                Property delegatorProperty = new Property();
                delegatorProperty.Tag = propertyTag;
                delegatorProperty.Type = PropertyType.String8;
                delegatorProperty.Size = (uint)(delegatorBuffer.Length + encoding.GetBytes("\0").Length);
                delegatorProperty.IsReadable = true;
                delegatorProperty.IsWriteable = true;

                propertiesMemoryStream.Write(delegatorProperty.ToBytes(), 0, 16);
            }

            if (percentComplete > 0)
            {
                NamedProperty percentCompleteNamedProperty = new NamedProperty();
                percentCompleteNamedProperty.Id = 0x8102;
                percentCompleteNamedProperty.Guid = StandardPropertySet.Task;
                percentCompleteNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, percentCompleteNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(percentCompleteNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0005;

                Property percentCompleteProperty = new Property();
                percentCompleteProperty.Tag = propertyTag;
                percentCompleteProperty.Type = PropertyType.Floating64;
                percentCompleteProperty.Value = BitConverter.GetBytes(percentComplete / 100);
                percentCompleteProperty.IsReadable = true;
                percentCompleteProperty.IsWriteable = true;

                propertiesMemoryStream.Write(percentCompleteProperty.ToBytes(), 0, 16);
            }

            if (actualWork > 0)
            {
                NamedProperty actualWorkNamedProperty = new NamedProperty();
                actualWorkNamedProperty.Id = 0x8110;
                actualWorkNamedProperty.Guid = StandardPropertySet.Task;
                actualWorkNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, actualWorkNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(actualWorkNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property actualWorkProperty = new Property();
                actualWorkProperty.Tag = propertyTag;
                actualWorkProperty.Type = PropertyType.Integer32;
                actualWorkProperty.Value = BitConverter.GetBytes(actualWork);
                actualWorkProperty.IsReadable = true;
                actualWorkProperty.IsWriteable = true;

                propertiesMemoryStream.Write(actualWorkProperty.ToBytes(), 0, 16);
            }

            if (totalWork > 0)
            {
                NamedProperty totalWorkNamedProperty = new NamedProperty();
                totalWorkNamedProperty.Id = 0x8111;
                totalWorkNamedProperty.Guid = StandardPropertySet.Task;
                totalWorkNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, totalWorkNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(totalWorkNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property totalWorkProperty = new Property();
                totalWorkProperty.Tag = propertyTag;
                totalWorkProperty.Type = PropertyType.Integer32;
                totalWorkProperty.Value = BitConverter.GetBytes(totalWork);
                totalWorkProperty.IsReadable = true;
                totalWorkProperty.IsWriteable = true;

                propertiesMemoryStream.Write(totalWorkProperty.ToBytes(), 0, 16);
            }

            if (isTeamTask)
            {
                NamedProperty isTeamTaskNamedProperty = new NamedProperty();
                isTeamTaskNamedProperty.Id = 0x8103;
                isTeamTaskNamedProperty.Guid = StandardPropertySet.Task;
                isTeamTaskNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isTeamTaskNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isTeamTaskNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isTeamTaskProperty = new Property();
                isTeamTaskProperty.Tag = propertyTag;
                isTeamTaskProperty.Type = PropertyType.Boolean;
                isTeamTaskProperty.Value = BitConverter.GetBytes(1);
                isTeamTaskProperty.IsReadable = true;
                isTeamTaskProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isTeamTaskProperty.ToBytes(), 0, 16);
            }

            if (isComplete)
            {
                NamedProperty isCompleteNamedProperty = new NamedProperty();
                isCompleteNamedProperty.Id = 0x811C;
                isCompleteNamedProperty.Guid = StandardPropertySet.Task;
                isCompleteNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isCompleteNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isCompleteNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isCompleteProperty = new Property();
                isCompleteProperty.Tag = propertyTag;
                isCompleteProperty.Type = PropertyType.Boolean;
                isCompleteProperty.Value = BitConverter.GetBytes(1);
                isCompleteProperty.IsReadable = true;
                isCompleteProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isCompleteProperty.ToBytes(), 0, 16);
            }

            if (isRecurring)
            {
                NamedProperty isRecurringNamedProperty = new NamedProperty();
                isRecurringNamedProperty.Id = 0x8223;
                isRecurringNamedProperty.Guid = StandardPropertySet.Appointment;
                isRecurringNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isRecurringNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isRecurringNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isRecurringProperty = new Property();
                isRecurringProperty.Tag = propertyTag;
                isRecurringProperty.Type = PropertyType.Boolean;
                isRecurringProperty.Value = BitConverter.GetBytes(1);
                isRecurringProperty.IsReadable = true;
                isRecurringProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isRecurringProperty.ToBytes(), 0, 16);
            }

            if (isAllDayEvent)
            {
                NamedProperty isAllDayEventNamedProperty = new NamedProperty();
                isAllDayEventNamedProperty.Id = 0x8215;
                isAllDayEventNamedProperty.Guid = StandardPropertySet.Appointment;
                isAllDayEventNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isAllDayEventNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isAllDayEventNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isAllDayEventProperty = new Property();
                isAllDayEventProperty.Tag = propertyTag;
                isAllDayEventProperty.Type = PropertyType.Boolean;
                isAllDayEventProperty.Value = BitConverter.GetBytes(1);
                isAllDayEventProperty.IsReadable = true;
                isAllDayEventProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isAllDayEventProperty.ToBytes(), 0, 16);
            }

            if (isReminderSet)
            {
                NamedProperty isReminderSetNamedProperty = new NamedProperty();
                isReminderSetNamedProperty.Id = 0x8503;
                isReminderSetNamedProperty.Guid = StandardPropertySet.Common;
                isReminderSetNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, isReminderSetNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(isReminderSetNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property isReminderSetProperty = new Property();
                isReminderSetProperty.Tag = propertyTag;
                isReminderSetProperty.Type = PropertyType.Boolean;
                isReminderSetProperty.Value = BitConverter.GetBytes(1);
                isReminderSetProperty.IsReadable = true;
                isReminderSetProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isReminderSetProperty.ToBytes(), 0, 16);
            }

            if (reminderTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty reminderTimeNamedProperty = new NamedProperty();
                reminderTimeNamedProperty.Id = 0x8502;
                reminderTimeNamedProperty.Guid = StandardPropertySet.Common;
                reminderTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, reminderTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(reminderTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = reminderTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property reminderTimeProperty = new Property();
                reminderTimeProperty.Tag = propertyTag;
                reminderTimeProperty.Type = PropertyType.Time;
                reminderTimeProperty.Value = ticksBytes;
                reminderTimeProperty.IsReadable = true;
                reminderTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reminderTimeProperty.ToBytes(), 0, 16);
            }

            if (reminderMinutesBeforeStart > 0)
            {
                NamedProperty reminderMinutesBeforeStartNamedProperty = new NamedProperty();
                reminderMinutesBeforeStartNamedProperty.Id = 0x8501;
                reminderMinutesBeforeStartNamedProperty.Guid = StandardPropertySet.Common;
                reminderMinutesBeforeStartNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, reminderMinutesBeforeStartNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(reminderMinutesBeforeStartNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property reminderMinutesBeforeStartProperty = new Property();
                reminderMinutesBeforeStartProperty.Tag = propertyTag;
                reminderMinutesBeforeStartProperty.Type = PropertyType.Integer32;
                reminderMinutesBeforeStartProperty.Value = BitConverter.GetBytes(reminderMinutesBeforeStart);
                reminderMinutesBeforeStartProperty.IsReadable = true;
                reminderMinutesBeforeStartProperty.IsWriteable = true;

                propertiesMemoryStream.Write(reminderMinutesBeforeStartProperty.ToBytes(), 0, 16);
            }

            if (taskStartDate.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty taskStartDateNamedProperty = new NamedProperty();
                taskStartDateNamedProperty.Id = 0x8104;
                taskStartDateNamedProperty.Guid = StandardPropertySet.Task;
                taskStartDateNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, taskStartDateNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(taskStartDateNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = taskStartDate.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property taskStartDateProperty = new Property();
                taskStartDateProperty.Tag = propertyTag;
                taskStartDateProperty.Type = PropertyType.Time;
                taskStartDateProperty.Value = ticksBytes;
                taskStartDateProperty.IsReadable = true;
                taskStartDateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(taskStartDateProperty.ToBytes(), 0, 16);
            }

            if (taskDueDate.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty taskDueDateNamedProperty = new NamedProperty();
                taskDueDateNamedProperty.Id = 0x8105;
                taskDueDateNamedProperty.Guid = StandardPropertySet.Task;
                taskDueDateNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, taskDueDateNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(taskDueDateNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = taskDueDate.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property taskDueDateProperty = new Property();
                taskDueDateProperty.Tag = propertyTag;
                taskDueDateProperty.Type = PropertyType.Time;
                taskDueDateProperty.Value = ticksBytes;
                taskDueDateProperty.IsReadable = true;
                taskDueDateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(taskDueDateProperty.ToBytes(), 0, 16);
            }

            if (dateCompleted.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty dateCompletedNamedProperty = new NamedProperty();
                dateCompletedNamedProperty.Id = 0x810F;
                dateCompletedNamedProperty.Guid = StandardPropertySet.Task;
                dateCompletedNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, dateCompletedNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(dateCompletedNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = dateCompleted.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property dateCompletedProperty = new Property();
                dateCompletedProperty.Tag = propertyTag;
                dateCompletedProperty.Type = PropertyType.Time;
                dateCompletedProperty.Value = ticksBytes;
                dateCompletedProperty.IsReadable = true;
                dateCompletedProperty.IsWriteable = true;

                propertiesMemoryStream.Write(dateCompletedProperty.ToBytes(), 0, 16);
            }

            if (taskStatus != TaskStatus.None)
            {
                NamedProperty taskStatusNamedProperty = new NamedProperty();
                taskStatusNamedProperty.Id = 0x8101;
                taskStatusNamedProperty.Guid = StandardPropertySet.Task;
                taskStatusNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, taskStatusNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(taskStatusNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property taskStatusProperty = new Property();
                taskStatusProperty.Tag = propertyTag;
                taskStatusProperty.Type = PropertyType.Integer32;
                taskStatusProperty.Value = BitConverter.GetBytes(EnumUtil.ParseTaskStatus(taskStatus));
                taskStatusProperty.IsReadable = true;
                taskStatusProperty.IsWriteable = true;

                propertiesMemoryStream.Write(taskStatusProperty.ToBytes(), 0, 16);
            }

            if (taskOwnership != TaskOwnership.None)
            {
                NamedProperty taskOwnershipNamedProperty = new NamedProperty();
                taskOwnershipNamedProperty.Id = 0x8129;
                taskOwnershipNamedProperty.Guid = StandardPropertySet.Task;
                taskOwnershipNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, taskOwnershipNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(taskOwnershipNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property taskOwnershipProperty = new Property();
                taskOwnershipProperty.Tag = propertyTag;
                taskOwnershipProperty.Type = PropertyType.Integer32;
                taskOwnershipProperty.Value = BitConverter.GetBytes(EnumUtil.ParseTaskOwnership(taskOwnership));
                taskOwnershipProperty.IsReadable = true;
                taskOwnershipProperty.IsWriteable = true;

                propertiesMemoryStream.Write(taskOwnershipProperty.ToBytes(), 0, 16);
            }

            if (taskDelegationState != TaskDelegationState.None)
            {
                NamedProperty taskDelegationStateNamedProperty = new NamedProperty();
                taskDelegationStateNamedProperty.Id = 0x812A;
                taskDelegationStateNamedProperty.Guid = StandardPropertySet.Task;
                taskDelegationStateNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, taskDelegationStateNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(taskDelegationStateNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property taskDelegationStateProperty = new Property();
                taskDelegationStateProperty.Tag = propertyTag;
                taskDelegationStateProperty.Type = PropertyType.Integer32;
                taskDelegationStateProperty.Value = BitConverter.GetBytes(EnumUtil.ParseTaskDelegationState(taskDelegationState));
                taskDelegationStateProperty.IsReadable = true;
                taskDelegationStateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(taskDelegationStateProperty.ToBytes(), 0, 16);
            }

            if (noteTop > 0)
            {
                NamedProperty noteTopNamedProperty = new NamedProperty();
                noteTopNamedProperty.Id = 0x8B05;
                noteTopNamedProperty.Guid = StandardPropertySet.Note;
                noteTopNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, noteTopNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(noteTopNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property noteTopProperty = new Property();
                noteTopProperty.Tag = propertyTag;
                noteTopProperty.Type = PropertyType.Integer32;
                noteTopProperty.Value = BitConverter.GetBytes(noteTop);
                noteTopProperty.IsReadable = true;
                noteTopProperty.IsWriteable = true;

                propertiesMemoryStream.Write(noteTopProperty.ToBytes(), 0, 16);
            }

            if (noteLeft > 0)
            {
                NamedProperty noteLeftNamedProperty = new NamedProperty();
                noteLeftNamedProperty.Id = 0x8B04;
                noteLeftNamedProperty.Guid = StandardPropertySet.Note;
                noteLeftNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, noteLeftNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(noteLeftNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property noteLeftProperty = new Property();
                noteLeftProperty.Tag = propertyTag;
                noteLeftProperty.Type = PropertyType.Integer32;
                noteLeftProperty.Value = BitConverter.GetBytes(noteLeft);
                noteLeftProperty.IsReadable = true;
                noteLeftProperty.IsWriteable = true;

                propertiesMemoryStream.Write(noteLeftProperty.ToBytes(), 0, 16);
            }


            if (noteHeight > 0)
            {
                NamedProperty noteHeightNamedProperty = new NamedProperty();
                noteHeightNamedProperty.Id = 0x8B03;
                noteHeightNamedProperty.Guid = StandardPropertySet.Note;
                noteHeightNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, noteHeightNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(noteHeightNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property noteHeightProperty = new Property();
                noteHeightProperty.Tag = propertyTag;
                noteHeightProperty.Type = PropertyType.Integer32;
                noteHeightProperty.Value = BitConverter.GetBytes(noteHeight);
                noteHeightProperty.IsReadable = true;
                noteHeightProperty.IsWriteable = true;

                propertiesMemoryStream.Write(noteHeightProperty.ToBytes(), 0, 16);
            }

            if (noteWidth > 0)
            {
                NamedProperty noteWidthNamedProperty = new NamedProperty();
                noteWidthNamedProperty.Id = 0x8B02;
                noteWidthNamedProperty.Guid = StandardPropertySet.Note;
                noteWidthNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, noteWidthNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(noteWidthNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property noteWidthProperty = new Property();
                noteWidthProperty.Tag = propertyTag;
                noteWidthProperty.Type = PropertyType.Integer32;
                noteWidthProperty.Value = BitConverter.GetBytes(noteWidth);
                noteWidthProperty.IsReadable = true;
                noteWidthProperty.IsWriteable = true;

                propertiesMemoryStream.Write(noteWidthProperty.ToBytes(), 0, 16);
            }

            if (noteColor != NoteColor.None)
            {
                NamedProperty noteColorNamedProperty = new NamedProperty();
                noteColorNamedProperty.Id = 0x8B00;
                noteColorNamedProperty.Guid = StandardPropertySet.Note;
                noteColorNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, noteColorNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(noteColorNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property noteColorProperty = new Property();
                noteColorProperty.Tag = propertyTag;
                noteColorProperty.Type = PropertyType.Integer32;
                noteColorProperty.Value = BitConverter.GetBytes(EnumUtil.ParseNoteColor(noteColor));
                noteColorProperty.IsReadable = true;
                noteColorProperty.IsWriteable = true;

                propertiesMemoryStream.Write(noteColorProperty.ToBytes(), 0, 16);
            }

            if (journalStartTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty journalStartTimeNamedProperty = new NamedProperty();
                journalStartTimeNamedProperty.Id = 0x8706;
                journalStartTimeNamedProperty.Guid = StandardPropertySet.Journal;
                journalStartTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, journalStartTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(journalStartTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = journalStartTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property journalStartTimeProperty = new Property();
                journalStartTimeProperty.Tag = propertyTag;
                journalStartTimeProperty.Type = PropertyType.Time;
                journalStartTimeProperty.Value = ticksBytes;
                journalStartTimeProperty.IsReadable = true;
                journalStartTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(journalStartTimeProperty.ToBytes(), 0, 16);
            }

            if (journalEndTime.CompareTo(DateTime.MinValue) > 0)
            {
                NamedProperty journalEndTimeNamedProperty = new NamedProperty();
                journalEndTimeNamedProperty.Id = 0x8708;
                journalEndTimeNamedProperty.Guid = StandardPropertySet.Journal;
                journalEndTimeNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, journalEndTimeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(journalEndTimeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0040;

                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = journalEndTime.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property journalEndTimeProperty = new Property();
                journalEndTimeProperty.Tag = propertyTag;
                journalEndTimeProperty.Type = PropertyType.Time;
                journalEndTimeProperty.Value = ticksBytes;
                journalEndTimeProperty.IsReadable = true;
                journalEndTimeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(journalEndTimeProperty.ToBytes(), 0, 16);
            }

            if (journalType != null)
            {
                NamedProperty journalTypeNamedProperty = new NamedProperty();
                journalTypeNamedProperty.Id = 0x8700;
                journalTypeNamedProperty.Guid = StandardPropertySet.Journal;
                journalTypeNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, journalTypeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(journalTypeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] journalTypeBuffer = encoding.GetBytes(journalType);
                Independentsoft.IO.StructuredStorage.Stream journalTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, journalTypeBuffer);
                directoryEntries.Add(journalTypeStream);

                Property journalTypeProperty = new Property();
                journalTypeProperty.Tag = propertyTag;
                journalTypeProperty.Type = PropertyType.String8;
                journalTypeProperty.Size = (uint)(journalTypeBuffer.Length + encoding.GetBytes("\0").Length);
                journalTypeProperty.IsReadable = true;
                journalTypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(journalTypeProperty.ToBytes(), 0, 16);
            }

            if (journalTypeDescription != null)
            {
                NamedProperty journalTypeDescriptionNamedProperty = new NamedProperty();
                journalTypeDescriptionNamedProperty.Id = 0x8712;
                journalTypeDescriptionNamedProperty.Guid = StandardPropertySet.Journal;
                journalTypeDescriptionNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, journalTypeDescriptionNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(journalTypeDescriptionNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] journalTypeDescriptionBuffer = encoding.GetBytes(journalTypeDescription);
                Independentsoft.IO.StructuredStorage.Stream journalTypeDescriptionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, journalTypeDescriptionBuffer);
                directoryEntries.Add(journalTypeDescriptionStream);

                Property journalTypeDescriptionProperty = new Property();
                journalTypeDescriptionProperty.Tag = propertyTag;
                journalTypeDescriptionProperty.Type = PropertyType.String8;
                journalTypeDescriptionProperty.Size = (uint)(journalTypeDescriptionBuffer.Length + encoding.GetBytes("\0").Length);
                journalTypeDescriptionProperty.IsReadable = true;
                journalTypeDescriptionProperty.IsWriteable = true;

                propertiesMemoryStream.Write(journalTypeDescriptionProperty.ToBytes(), 0, 16);
            }

            if (journalDuration > 0)
            {
                NamedProperty journalDurationNamedProperty = new NamedProperty();
                journalDurationNamedProperty.Id = 0x8707;
                journalDurationNamedProperty.Guid = StandardPropertySet.Journal;
                journalDurationNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, journalDurationNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(journalDurationNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property journalDurationProperty = new Property();
                journalDurationProperty.Tag = propertyTag;
                journalDurationProperty.Type = PropertyType.Integer32;
                journalDurationProperty.Value = BitConverter.GetBytes(journalDuration);
                journalDurationProperty.IsReadable = true;
                journalDurationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(journalDurationProperty.ToBytes(), 0, 16);
            }

            if (birthday.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = birthday.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property birthdayProperty = new Property();
                birthdayProperty.Tag = 0x3A420040;
                birthdayProperty.Type = PropertyType.Time;
                birthdayProperty.Value = ticksBytes;
                birthdayProperty.IsReadable = true;
                birthdayProperty.IsWriteable = false;

                propertiesMemoryStream.Write(birthdayProperty.ToBytes(), 0, 16);
            }

            if (childrenNames.Count > 0)
            {
                MemoryStream childrenNamesMemoryStream = new MemoryStream();

                for (int i = 0; i < childrenNames.Count; i++)
                {
                    byte[] childrenNamesBuffer = encoding.GetBytes(childrenNames[i] + "\0");
                    uint size = (uint)childrenNamesBuffer.Length;
                    byte[] sizeBuffer = BitConverter.GetBytes(size);

                    childrenNamesMemoryStream.Write(sizeBuffer, 0, 4);

                    string streamName = "__substg1.0_3A58" + multiValueStringTypeMask + "-" + String.Format("{0:X8}", i);

                    Independentsoft.IO.StructuredStorage.Stream childrenNamesStream = new Independentsoft.IO.StructuredStorage.Stream(streamName, childrenNamesBuffer);
                    directoryEntries.Add(childrenNamesStream);
                }

                byte[] childrenNamesStreamBuffer = childrenNamesMemoryStream.ToArray();

                Independentsoft.IO.StructuredStorage.Stream childrenNamesLengthStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A58" + multiValueStringTypeMask, childrenNamesStreamBuffer);
                directoryEntries.Add(childrenNamesLengthStream);

                Property childrenNamesProperty = new Property();
                childrenNamesProperty.Tag = (uint)0x3A58 << 16 | multiValueStringTypeHexMask;
                childrenNamesProperty.Type = PropertyType.MultipleString8;
                childrenNamesProperty.Size = (uint)childrenNamesStreamBuffer.Length;
                childrenNamesProperty.IsReadable = true;
                childrenNamesProperty.IsWriteable = true;

                propertiesMemoryStream.Write(childrenNamesProperty.ToBytes(), 0, 16);
            }

            if (assistentName != null)
            {
                byte[] assistentNameBuffer = encoding.GetBytes(assistentName);
                Independentsoft.IO.StructuredStorage.Stream assistentNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A30" + stringTypeMask, assistentNameBuffer);
                directoryEntries.Add(assistentNameStream);

                Property assistentNameProperty = new Property();
                assistentNameProperty.Tag = (uint)0x3A30 << 16 | stringTypeHexMask;
                assistentNameProperty.Type = PropertyType.String8;
                assistentNameProperty.Size = (uint)(assistentNameBuffer.Length + encoding.GetBytes("\0").Length);
                assistentNameProperty.IsReadable = true;
                assistentNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(assistentNameProperty.ToBytes(), 0, 16);
            }

            if (assistentPhone != null)
            {
                byte[] assistentPhoneBuffer = encoding.GetBytes(assistentPhone);
                Independentsoft.IO.StructuredStorage.Stream assistentPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2E" + stringTypeMask, assistentPhoneBuffer);
                directoryEntries.Add(assistentPhoneStream);

                Property assistentPhoneProperty = new Property();
                assistentPhoneProperty.Tag = (uint)0x3A2E << 16 | stringTypeHexMask;
                assistentPhoneProperty.Type = PropertyType.String8;
                assistentPhoneProperty.Size = (uint)(assistentPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                assistentPhoneProperty.IsReadable = true;
                assistentPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(assistentPhoneProperty.ToBytes(), 0, 16);
            }

            if (businessPhone2 != null)
            {
                byte[] businessPhone2Buffer = encoding.GetBytes(businessPhone2);
                Independentsoft.IO.StructuredStorage.Stream businessPhone2Stream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1B" + stringTypeMask, businessPhone2Buffer);
                directoryEntries.Add(businessPhone2Stream);

                Property businessPhone2Property = new Property();
                businessPhone2Property.Tag = (uint)0x3A1B << 16 | stringTypeHexMask;
                businessPhone2Property.Type = PropertyType.String8;
                businessPhone2Property.Size = (uint)(businessPhone2Buffer.Length + encoding.GetBytes("\0").Length);
                businessPhone2Property.IsReadable = true;
                businessPhone2Property.IsWriteable = true;

                propertiesMemoryStream.Write(businessPhone2Property.ToBytes(), 0, 16);
            }

            if (businessFax != null)
            {
                byte[] businessFaxBuffer = encoding.GetBytes(businessFax);
                Independentsoft.IO.StructuredStorage.Stream businessFaxStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A24" + stringTypeMask, businessFaxBuffer);
                directoryEntries.Add(businessFaxStream);

                Property businessFaxProperty = new Property();
                businessFaxProperty.Tag = (uint)0x3A24 << 16 | stringTypeHexMask;
                businessFaxProperty.Type = PropertyType.String8;
                businessFaxProperty.Size = (uint)(businessFaxBuffer.Length + encoding.GetBytes("\0").Length);
                businessFaxProperty.IsReadable = true;
                businessFaxProperty.IsWriteable = true;

                propertiesMemoryStream.Write(businessFaxProperty.ToBytes(), 0, 16);
            }

            if (businessHomePage != null)
            {
                byte[] businessHomePageBuffer = encoding.GetBytes(businessHomePage);
                Independentsoft.IO.StructuredStorage.Stream businessHomePageStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A51" + stringTypeMask, businessHomePageBuffer);
                directoryEntries.Add(businessHomePageStream);

                Property businessHomePageProperty = new Property();
                businessHomePageProperty.Tag = (uint)0x3A51 << 16 | stringTypeHexMask;
                businessHomePageProperty.Type = PropertyType.String8;
                businessHomePageProperty.Size = (uint)(businessHomePageBuffer.Length + encoding.GetBytes("\0").Length);
                businessHomePageProperty.IsReadable = true;
                businessHomePageProperty.IsWriteable = true;

                propertiesMemoryStream.Write(businessHomePageProperty.ToBytes(), 0, 16);
            }

            if (callbackPhone != null)
            {
                byte[] callbackPhoneBuffer = encoding.GetBytes(callbackPhone);
                Independentsoft.IO.StructuredStorage.Stream callbackPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A02" + stringTypeMask, callbackPhoneBuffer);
                directoryEntries.Add(callbackPhoneStream);

                Property callbackPhoneProperty = new Property();
                callbackPhoneProperty.Tag = (uint)0x3A02 << 16 | stringTypeHexMask;
                callbackPhoneProperty.Type = PropertyType.String8;
                callbackPhoneProperty.Size = (uint)(callbackPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                callbackPhoneProperty.IsReadable = true;
                callbackPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(callbackPhoneProperty.ToBytes(), 0, 16);
            }

            if (carPhone != null)
            {
                byte[] carPhoneBuffer = encoding.GetBytes(carPhone);
                Independentsoft.IO.StructuredStorage.Stream carPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1E" + stringTypeMask, carPhoneBuffer);
                directoryEntries.Add(carPhoneStream);

                Property carPhoneProperty = new Property();
                carPhoneProperty.Tag = (uint)0x3A1E << 16 | stringTypeHexMask;
                carPhoneProperty.Type = PropertyType.String8;
                carPhoneProperty.Size = (uint)(carPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                carPhoneProperty.IsReadable = true;
                carPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(carPhoneProperty.ToBytes(), 0, 16);
            }

            if (cellularPhone != null)
            {
                byte[] cellularPhoneBuffer = encoding.GetBytes(cellularPhone);
                Independentsoft.IO.StructuredStorage.Stream cellularPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1C" + stringTypeMask, cellularPhoneBuffer);
                directoryEntries.Add(cellularPhoneStream);

                Property cellularPhoneProperty = new Property();
                cellularPhoneProperty.Tag = (uint)0x3A1C << 16 | stringTypeHexMask;
                cellularPhoneProperty.Type = PropertyType.String8;
                cellularPhoneProperty.Size = (uint)(cellularPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                cellularPhoneProperty.IsReadable = true;
                cellularPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(cellularPhoneProperty.ToBytes(), 0, 16);
            }

            if (companyMainPhone != null)
            {
                byte[] companyMainPhoneBuffer = encoding.GetBytes(companyMainPhone);
                Independentsoft.IO.StructuredStorage.Stream companyMainPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A57" + stringTypeMask, companyMainPhoneBuffer);
                directoryEntries.Add(companyMainPhoneStream);

                Property companyMainPhoneProperty = new Property();
                companyMainPhoneProperty.Tag = (uint)0x3A57 << 16 | stringTypeHexMask;
                companyMainPhoneProperty.Type = PropertyType.String8;
                companyMainPhoneProperty.Size = (uint)(companyMainPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                companyMainPhoneProperty.IsReadable = true;
                companyMainPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(companyMainPhoneProperty.ToBytes(), 0, 16);
            }

            if (companyName != null)
            {
                byte[] companyNameBuffer = encoding.GetBytes(companyName);
                Independentsoft.IO.StructuredStorage.Stream companyNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A16" + stringTypeMask, companyNameBuffer);
                directoryEntries.Add(companyNameStream);

                Property companyNameProperty = new Property();
                companyNameProperty.Tag = (uint)0x3A16 << 16 | stringTypeHexMask;
                companyNameProperty.Type = PropertyType.String8;
                companyNameProperty.Size = (uint)(companyNameBuffer.Length + encoding.GetBytes("\0").Length);
                companyNameProperty.IsReadable = true;
                companyNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(companyNameProperty.ToBytes(), 0, 16);
            }

            if (computerNetworkName != null)
            {
                byte[] computerNetworkNameBuffer = encoding.GetBytes(computerNetworkName);
                Independentsoft.IO.StructuredStorage.Stream computerNetworkNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A49" + stringTypeMask, computerNetworkNameBuffer);
                directoryEntries.Add(computerNetworkNameStream);

                Property computerNetworkNameProperty = new Property();
                computerNetworkNameProperty.Tag = (uint)0x3A49 << 16 | stringTypeHexMask;
                computerNetworkNameProperty.Type = PropertyType.String8;
                computerNetworkNameProperty.Size = (uint)(computerNetworkNameBuffer.Length + encoding.GetBytes("\0").Length);
                computerNetworkNameProperty.IsReadable = true;
                computerNetworkNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(computerNetworkNameProperty.ToBytes(), 0, 16);
            }

            if (businessAddressCountry != null)
            {
                NamedProperty businessAddressCountryNamedProperty = new NamedProperty();
                businessAddressCountryNamedProperty.Id = 0x8049;
                businessAddressCountryNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressCountryNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressCountryNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressCountryNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressCountryBuffer = encoding.GetBytes(businessAddressCountry);

                Independentsoft.IO.StructuredStorage.Stream businessAddressCountryStream1 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressCountryBuffer);
                directoryEntries.Add(businessAddressCountryStream1);

                Property businessAddressCountryProperty1 = new Property();
                businessAddressCountryProperty1.Tag = propertyTag;
                businessAddressCountryProperty1.Type = PropertyType.String8;
                businessAddressCountryProperty1.Size = (uint)(businessAddressCountryBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressCountryProperty1.IsReadable = true;
                businessAddressCountryProperty1.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressCountryProperty1.ToBytes(), 0, 16);

                Independentsoft.IO.StructuredStorage.Stream businessAddressCountryStream2 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A26" + stringTypeMask, businessAddressCountryBuffer);
                directoryEntries.Add(businessAddressCountryStream2);

                Property businessAddressCountryProperty2 = new Property();
                businessAddressCountryProperty2.Tag = (uint)0x3A26 << 16 | stringTypeHexMask;
                businessAddressCountryProperty2.Type = PropertyType.String8;
                businessAddressCountryProperty2.Size = (uint)(businessAddressCountryBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressCountryProperty2.IsReadable = true;
                businessAddressCountryProperty2.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressCountryProperty2.ToBytes(), 0, 16);
            }

            if (customerId != null)
            {
                byte[] customerIdBuffer = encoding.GetBytes(customerId);
                Independentsoft.IO.StructuredStorage.Stream customerIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A4A" + stringTypeMask, customerIdBuffer);
                directoryEntries.Add(customerIdStream);

                Property customerIdProperty = new Property();
                customerIdProperty.Tag = (uint)0x3A4A << 16 | stringTypeHexMask;
                customerIdProperty.Type = PropertyType.String8;
                customerIdProperty.Size = (uint)(customerIdBuffer.Length + encoding.GetBytes("\0").Length);
                customerIdProperty.IsReadable = true;
                customerIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(customerIdProperty.ToBytes(), 0, 16);
            }

            if (departmentName != null)
            {
                byte[] departmentNameBuffer = encoding.GetBytes(departmentName);
                Independentsoft.IO.StructuredStorage.Stream departmentNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A18" + stringTypeMask, departmentNameBuffer);
                directoryEntries.Add(departmentNameStream);

                Property departmentNameProperty = new Property();
                departmentNameProperty.Tag = (uint)0x3A18 << 16 | stringTypeHexMask;
                departmentNameProperty.Type = PropertyType.String8;
                departmentNameProperty.Size = (uint)(departmentNameBuffer.Length + encoding.GetBytes("\0").Length);
                departmentNameProperty.IsReadable = true;
                departmentNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(departmentNameProperty.ToBytes(), 0, 16);
            }

            if (displayName != null)
            {
                byte[] displayNameBuffer = encoding.GetBytes(displayName);
                Independentsoft.IO.StructuredStorage.Stream displayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3001" + stringTypeMask, displayNameBuffer);
                directoryEntries.Add(displayNameStream);

                Property displayNameProperty = new Property();
                displayNameProperty.Tag = (uint)0x3001 << 16 | stringTypeHexMask;
                displayNameProperty.Type = PropertyType.String8;
                displayNameProperty.Size = (uint)(displayNameBuffer.Length + encoding.GetBytes("\0").Length);
                displayNameProperty.IsReadable = true;
                displayNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(displayNameProperty.ToBytes(), 0, 16);
            }

            if (displayNamePrefix != null)
            {
                byte[] displayNamePrefixBuffer = encoding.GetBytes(displayNamePrefix);
                Independentsoft.IO.StructuredStorage.Stream displayNamePrefixStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A45" + stringTypeMask, displayNamePrefixBuffer);
                directoryEntries.Add(displayNamePrefixStream);

                Property displayNamePrefixProperty = new Property();
                displayNamePrefixProperty.Tag = (uint)0x3A45 << 16 | stringTypeHexMask;
                displayNamePrefixProperty.Type = PropertyType.String8;
                displayNamePrefixProperty.Size = (uint)(displayNamePrefixBuffer.Length + encoding.GetBytes("\0").Length);
                displayNamePrefixProperty.IsReadable = true;
                displayNamePrefixProperty.IsWriteable = true;

                propertiesMemoryStream.Write(displayNamePrefixProperty.ToBytes(), 0, 16);
            }

            if (ftpSite != null)
            {
                byte[] ftpSiteBuffer = encoding.GetBytes(ftpSite);
                Independentsoft.IO.StructuredStorage.Stream ftpSiteStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A4C" + stringTypeMask, ftpSiteBuffer);
                directoryEntries.Add(ftpSiteStream);

                Property ftpSiteProperty = new Property();
                ftpSiteProperty.Tag = (uint)0x3A4C << 16 | stringTypeHexMask;
                ftpSiteProperty.Type = PropertyType.String8;
                ftpSiteProperty.Size = (uint)(ftpSiteBuffer.Length + encoding.GetBytes("\0").Length);
                ftpSiteProperty.IsReadable = true;
                ftpSiteProperty.IsWriteable = true;

                propertiesMemoryStream.Write(ftpSiteProperty.ToBytes(), 0, 16);
            }

            if (generation != null)
            {
                byte[] generationBuffer = encoding.GetBytes(generation);
                Independentsoft.IO.StructuredStorage.Stream generationStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A05" + stringTypeMask, generationBuffer);
                directoryEntries.Add(generationStream);

                Property generationProperty = new Property();
                generationProperty.Tag = (uint)0x3A05 << 16 | stringTypeHexMask;
                generationProperty.Type = PropertyType.String8;
                generationProperty.Size = (uint)(generationBuffer.Length + encoding.GetBytes("\0").Length);
                generationProperty.IsReadable = true;
                generationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(generationProperty.ToBytes(), 0, 16);
            }

            if (givenName != null)
            {
                byte[] givenNameBuffer = encoding.GetBytes(givenName);
                Independentsoft.IO.StructuredStorage.Stream givenNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A06" + stringTypeMask, givenNameBuffer);
                directoryEntries.Add(givenNameStream);

                Property givenNameProperty = new Property();
                givenNameProperty.Tag = (uint)0x3A06 << 16 | stringTypeHexMask;
                givenNameProperty.Type = PropertyType.String8;
                givenNameProperty.Size = (uint)(givenNameBuffer.Length + encoding.GetBytes("\0").Length);
                givenNameProperty.IsReadable = true;
                givenNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(givenNameProperty.ToBytes(), 0, 16);
            }

            if (governmentId != null)
            {
                byte[] governmentIdBuffer = encoding.GetBytes(governmentId);
                Independentsoft.IO.StructuredStorage.Stream governmentIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A07" + stringTypeMask, governmentIdBuffer);
                directoryEntries.Add(governmentIdStream);

                Property governmentIdProperty = new Property();
                governmentIdProperty.Tag = (uint)0x3A07 << 16 | stringTypeHexMask;
                governmentIdProperty.Type = PropertyType.String8;
                governmentIdProperty.Size = (uint)(governmentIdBuffer.Length + encoding.GetBytes("\0").Length);
                governmentIdProperty.IsReadable = true;
                governmentIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(governmentIdProperty.ToBytes(), 0, 16);
            }

            if (hobbies != null)
            {
                byte[] hobbiesBuffer = encoding.GetBytes(hobbies);
                Independentsoft.IO.StructuredStorage.Stream hobbiesStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A43" + stringTypeMask, hobbiesBuffer);
                directoryEntries.Add(hobbiesStream);

                Property hobbiesProperty = new Property();
                hobbiesProperty.Tag = (uint)0x3A43 << 16 | stringTypeHexMask;
                hobbiesProperty.Type = PropertyType.String8;
                hobbiesProperty.Size = (uint)(hobbiesBuffer.Length + encoding.GetBytes("\0").Length);
                hobbiesProperty.IsReadable = true;
                hobbiesProperty.IsWriteable = true;

                propertiesMemoryStream.Write(hobbiesProperty.ToBytes(), 0, 16);
            }

            if (homePhone2 != null)
            {
                byte[] homePhone2Buffer = encoding.GetBytes(homePhone2);
                Independentsoft.IO.StructuredStorage.Stream homePhone2Stream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2F" + stringTypeMask, homePhone2Buffer);
                directoryEntries.Add(homePhone2Stream);

                Property homePhone2Property = new Property();
                homePhone2Property.Tag = (uint)0x3A2F << 16 | stringTypeHexMask;
                homePhone2Property.Type = PropertyType.String8;
                homePhone2Property.Size = (uint)(homePhone2Buffer.Length + encoding.GetBytes("\0").Length);
                homePhone2Property.IsReadable = true;
                homePhone2Property.IsWriteable = true;

                propertiesMemoryStream.Write(homePhone2Property.ToBytes(), 0, 16);
            }

            if (homeAddressCity != null)
            {
                byte[] homeAddressCityBuffer = encoding.GetBytes(homeAddressCity);
                Independentsoft.IO.StructuredStorage.Stream homeAddressCityStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A59" + stringTypeMask, homeAddressCityBuffer);
                directoryEntries.Add(homeAddressCityStream);

                Property homeAddressCityProperty = new Property();
                homeAddressCityProperty.Tag = (uint)0x3A59 << 16 | stringTypeHexMask;
                homeAddressCityProperty.Type = PropertyType.String8;
                homeAddressCityProperty.Size = (uint)(homeAddressCityBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressCityProperty.IsReadable = true;
                homeAddressCityProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressCityProperty.ToBytes(), 0, 16);
            }

            if (homeAddressCountry != null)
            {
                byte[] homeAddressCountryBuffer = encoding.GetBytes(homeAddressCountry);
                Independentsoft.IO.StructuredStorage.Stream homeAddressCountryStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5A" + stringTypeMask, homeAddressCountryBuffer);
                directoryEntries.Add(homeAddressCountryStream);

                Property homeAddressCountryProperty = new Property();
                homeAddressCountryProperty.Tag = (uint)0x3A5A << 16 | stringTypeHexMask;
                homeAddressCountryProperty.Type = PropertyType.String8;
                homeAddressCountryProperty.Size = (uint)(homeAddressCountryBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressCountryProperty.IsReadable = true;
                homeAddressCountryProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressCountryProperty.ToBytes(), 0, 16);
            }

            if (homeAddressPostalCode != null)
            {
                byte[] homeAddressPostalCodeBuffer = encoding.GetBytes(homeAddressPostalCode);
                Independentsoft.IO.StructuredStorage.Stream homeAddressPostalCodeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5B" + stringTypeMask, homeAddressPostalCodeBuffer);
                directoryEntries.Add(homeAddressPostalCodeStream);

                Property homeAddressPostalCodeProperty = new Property();
                homeAddressPostalCodeProperty.Tag = (uint)0x3A5B << 16 | stringTypeHexMask;
                homeAddressPostalCodeProperty.Type = PropertyType.String8;
                homeAddressPostalCodeProperty.Size = (uint)(homeAddressPostalCodeBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressPostalCodeProperty.IsReadable = true;
                homeAddressPostalCodeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressPostalCodeProperty.ToBytes(), 0, 16);
            }

            if (homeAddressPostOfficeBox != null)
            {
                byte[] homeAddressPostOfficeBoxBuffer = encoding.GetBytes(homeAddressPostOfficeBox);
                Independentsoft.IO.StructuredStorage.Stream homeAddressPostOfficeBoxStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5E" + stringTypeMask, homeAddressPostOfficeBoxBuffer);
                directoryEntries.Add(homeAddressPostOfficeBoxStream);

                Property homeAddressPostOfficeBoxProperty = new Property();
                homeAddressPostOfficeBoxProperty.Tag = (uint)0x3A5E << 16 | stringTypeHexMask;
                homeAddressPostOfficeBoxProperty.Type = PropertyType.String8;
                homeAddressPostOfficeBoxProperty.Size = (uint)(homeAddressPostOfficeBoxBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressPostOfficeBoxProperty.IsReadable = true;
                homeAddressPostOfficeBoxProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressPostOfficeBoxProperty.ToBytes(), 0, 16);
            }

            if (homeAddressState != null)
            {
                byte[] homeAddressStateBuffer = encoding.GetBytes(homeAddressState);
                Independentsoft.IO.StructuredStorage.Stream homeAddressStateStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5C" + stringTypeMask, homeAddressStateBuffer);
                directoryEntries.Add(homeAddressStateStream);

                Property homeAddressStateProperty = new Property();
                homeAddressStateProperty.Tag = (uint)0x3A5C << 16 | stringTypeHexMask;
                homeAddressStateProperty.Type = PropertyType.String8;
                homeAddressStateProperty.Size = (uint)(homeAddressStateBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressStateProperty.IsReadable = true;
                homeAddressStateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressStateProperty.ToBytes(), 0, 16);
            }

            if (homeAddressStreet != null)
            {
                byte[] homeAddressStreetBuffer = encoding.GetBytes(homeAddressStreet);
                Independentsoft.IO.StructuredStorage.Stream homeAddressStreetStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5D" + stringTypeMask, homeAddressStreetBuffer);
                directoryEntries.Add(homeAddressStreetStream);

                Property homeAddressStreetProperty = new Property();
                homeAddressStreetProperty.Tag = (uint)0x3A5D << 16 | stringTypeHexMask;
                homeAddressStreetProperty.Type = PropertyType.String8;
                homeAddressStreetProperty.Size = (uint)(homeAddressStreetBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressStreetProperty.IsReadable = true;
                homeAddressStreetProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressStreetProperty.ToBytes(), 0, 16);
            }

            if (homeFax != null)
            {
                byte[] homeFaxBuffer = encoding.GetBytes(homeFax);
                Independentsoft.IO.StructuredStorage.Stream homeFaxStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A25" + stringTypeMask, homeFaxBuffer);
                directoryEntries.Add(homeFaxStream);

                Property homeFaxProperty = new Property();
                homeFaxProperty.Tag = (uint)0x3A25 << 16 | stringTypeHexMask;
                homeFaxProperty.Type = PropertyType.String8;
                homeFaxProperty.Size = (uint)(homeFaxBuffer.Length + encoding.GetBytes("\0").Length);
                homeFaxProperty.IsReadable = true;
                homeFaxProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeFaxProperty.ToBytes(), 0, 16);
            }

            if (homePhone != null)
            {
                byte[] homePhoneBuffer = encoding.GetBytes(homePhone);
                Independentsoft.IO.StructuredStorage.Stream homePhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A09" + stringTypeMask, homePhoneBuffer);
                directoryEntries.Add(homePhoneStream);

                Property homePhoneProperty = new Property();
                homePhoneProperty.Tag = (uint)0x3A09 << 16 | stringTypeHexMask;
                homePhoneProperty.Type = PropertyType.String8;
                homePhoneProperty.Size = (uint)(homePhoneBuffer.Length + encoding.GetBytes("\0").Length);
                homePhoneProperty.IsReadable = true;
                homePhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homePhoneProperty.ToBytes(), 0, 16);
            }

            if (initials != null)
            {
                byte[] initialsBuffer = encoding.GetBytes(initials);
                Independentsoft.IO.StructuredStorage.Stream initialsStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A0A" + stringTypeMask, initialsBuffer);
                directoryEntries.Add(initialsStream);

                Property initialsProperty = new Property();
                initialsProperty.Tag = (uint)0x3A0A << 16 | stringTypeHexMask;
                initialsProperty.Type = PropertyType.String8;
                initialsProperty.Size = (uint)(initialsBuffer.Length + encoding.GetBytes("\0").Length);
                initialsProperty.IsReadable = true;
                initialsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(initialsProperty.ToBytes(), 0, 16);
            }

            if (isdn != null)
            {
                byte[] isdnBuffer = encoding.GetBytes(isdn);
                Independentsoft.IO.StructuredStorage.Stream isdnStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2D" + stringTypeMask, isdnBuffer);
                directoryEntries.Add(isdnStream);

                Property isdnProperty = new Property();
                isdnProperty.Tag = (uint)0x3A2D << 16 | stringTypeHexMask;
                isdnProperty.Type = PropertyType.String8;
                isdnProperty.Size = (uint)(isdnBuffer.Length + encoding.GetBytes("\0").Length);
                isdnProperty.IsReadable = true;
                isdnProperty.IsWriteable = true;

                propertiesMemoryStream.Write(isdnProperty.ToBytes(), 0, 16);
            }

            if (businessAddressCity != null)
            {
                NamedProperty businessAddressCityNamedProperty = new NamedProperty();
                businessAddressCityNamedProperty.Id = 0x8046;
                businessAddressCityNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressCityNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressCityNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressCityNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressCityBuffer = encoding.GetBytes(businessAddressCity);

                Independentsoft.IO.StructuredStorage.Stream businessAddressCityStream1 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressCityBuffer);
                directoryEntries.Add(businessAddressCityStream1);

                Property businessAddressCityProperty1 = new Property();
                businessAddressCityProperty1.Tag = propertyTag;
                businessAddressCityProperty1.Type = PropertyType.String8;
                businessAddressCityProperty1.Size = (uint)(businessAddressCityBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressCityProperty1.IsReadable = true;
                businessAddressCityProperty1.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressCityProperty1.ToBytes(), 0, 16);

                Independentsoft.IO.StructuredStorage.Stream businessAddressCityStream2 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A27" + stringTypeMask, businessAddressCityBuffer);
                directoryEntries.Add(businessAddressCityStream2);

                Property businessAddressCityProperty2 = new Property();
                businessAddressCityProperty2.Tag = (uint)0x3A27 << 16 | stringTypeHexMask;
                businessAddressCityProperty2.Type = PropertyType.String8;
                businessAddressCityProperty2.Size = (uint)(businessAddressCityBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressCityProperty2.IsReadable = true;
                businessAddressCityProperty2.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressCityProperty2.ToBytes(), 0, 16);
            }

            if (managerName != null)
            {
                byte[] managerNameBuffer = encoding.GetBytes(managerName);
                Independentsoft.IO.StructuredStorage.Stream managerNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A4E" + stringTypeMask, managerNameBuffer);
                directoryEntries.Add(managerNameStream);

                Property managerNameProperty = new Property();
                managerNameProperty.Tag = (uint)0x3A4E << 16 | stringTypeHexMask;
                managerNameProperty.Type = PropertyType.String8;
                managerNameProperty.Size = (uint)(managerNameBuffer.Length + encoding.GetBytes("\0").Length);
                managerNameProperty.IsReadable = true;
                managerNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(managerNameProperty.ToBytes(), 0, 16);
            }

            if (middleName != null)
            {
                byte[] middleNameBuffer = encoding.GetBytes(middleName);
                Independentsoft.IO.StructuredStorage.Stream middleNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A44" + stringTypeMask, middleNameBuffer);
                directoryEntries.Add(middleNameStream);

                Property middleNameProperty = new Property();
                middleNameProperty.Tag = (uint)0x3A44 << 16 | stringTypeHexMask;
                middleNameProperty.Type = PropertyType.String8;
                middleNameProperty.Size = (uint)(middleNameBuffer.Length + encoding.GetBytes("\0").Length);
                middleNameProperty.IsReadable = true;
                middleNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(middleNameProperty.ToBytes(), 0, 16);
            }

            if (nickname != null)
            {
                byte[] nicknameBuffer = encoding.GetBytes(nickname);
                Independentsoft.IO.StructuredStorage.Stream nicknameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A4F" + stringTypeMask, nicknameBuffer);
                directoryEntries.Add(nicknameStream);

                Property nicknameProperty = new Property();
                nicknameProperty.Tag = (uint)0x3A4F << 16 | stringTypeHexMask;
                nicknameProperty.Type = PropertyType.String8;
                nicknameProperty.Size = (uint)(nicknameBuffer.Length + encoding.GetBytes("\0").Length);
                nicknameProperty.IsReadable = true;
                nicknameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(nicknameProperty.ToBytes(), 0, 16);
            }

            if (officeLocation != null)
            {
                byte[] officeLocationBuffer = encoding.GetBytes(officeLocation);
                Independentsoft.IO.StructuredStorage.Stream officeLocationStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A19" + stringTypeMask, officeLocationBuffer);
                directoryEntries.Add(officeLocationStream);

                Property officeLocationProperty = new Property();
                officeLocationProperty.Tag = (uint)0x3A19 << 16 | stringTypeHexMask;
                officeLocationProperty.Type = PropertyType.String8;
                officeLocationProperty.Size = (uint)(officeLocationBuffer.Length + encoding.GetBytes("\0").Length);
                officeLocationProperty.IsReadable = true;
                officeLocationProperty.IsWriteable = true;

                propertiesMemoryStream.Write(officeLocationProperty.ToBytes(), 0, 16);
            }

            if (businessPhone != null)
            {
                byte[] businessPhoneBuffer = encoding.GetBytes(businessPhone);
                Independentsoft.IO.StructuredStorage.Stream businessPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A08" + stringTypeMask, businessPhoneBuffer);
                directoryEntries.Add(businessPhoneStream);

                Property businessPhoneProperty = new Property();
                businessPhoneProperty.Tag = (uint)0x3A08 << 16 | stringTypeHexMask;
                businessPhoneProperty.Type = PropertyType.String8;
                businessPhoneProperty.Size = (uint)(businessPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                businessPhoneProperty.IsReadable = true;
                businessPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(businessPhoneProperty.ToBytes(), 0, 16);
            }

            if (otherAddressCity != null)
            {
                byte[] otherAddressCityBuffer = encoding.GetBytes(otherAddressCity);
                Independentsoft.IO.StructuredStorage.Stream otherAddressCityStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A5F" + stringTypeMask, otherAddressCityBuffer);
                directoryEntries.Add(otherAddressCityStream);

                Property otherAddressCityProperty = new Property();
                otherAddressCityProperty.Tag = (uint)0x3A5F << 16 | stringTypeHexMask;
                otherAddressCityProperty.Type = PropertyType.String8;
                otherAddressCityProperty.Size = (uint)(otherAddressCityBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressCityProperty.IsReadable = true;
                otherAddressCityProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressCityProperty.ToBytes(), 0, 16);
            }

            if (otherAddressCountry != null)
            {
                byte[] otherAddressCountryBuffer = encoding.GetBytes(otherAddressCountry);
                Independentsoft.IO.StructuredStorage.Stream otherAddressCountryStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A60" + stringTypeMask, otherAddressCountryBuffer);
                directoryEntries.Add(otherAddressCountryStream);

                Property otherAddressCountryProperty = new Property();
                otherAddressCountryProperty.Tag = (uint)0x3A60 << 16 | stringTypeHexMask;
                otherAddressCountryProperty.Type = PropertyType.String8;
                otherAddressCountryProperty.Size = (uint)(otherAddressCountryBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressCountryProperty.IsReadable = true;
                otherAddressCountryProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressCountryProperty.ToBytes(), 0, 16);
            }

            if (otherAddressPostalCode != null)
            {
                byte[] otherAddressPostalCodeBuffer = encoding.GetBytes(otherAddressPostalCode);
                Independentsoft.IO.StructuredStorage.Stream otherAddressPostalCodeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A61" + stringTypeMask, otherAddressPostalCodeBuffer);
                directoryEntries.Add(otherAddressPostalCodeStream);

                Property otherAddressPostalCodeProperty = new Property();
                otherAddressPostalCodeProperty.Tag = (uint)0x3A61 << 16 | stringTypeHexMask;
                otherAddressPostalCodeProperty.Type = PropertyType.String8;
                otherAddressPostalCodeProperty.Size = (uint)(otherAddressPostalCodeBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressPostalCodeProperty.IsReadable = true;
                otherAddressPostalCodeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressPostalCodeProperty.ToBytes(), 0, 16);
            }

            if (otherAddressState != null)
            {
                byte[] otherAddressStateBuffer = encoding.GetBytes(otherAddressState);
                Independentsoft.IO.StructuredStorage.Stream otherAddressStateStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A62" + stringTypeMask, otherAddressStateBuffer);
                directoryEntries.Add(otherAddressStateStream);

                Property otherAddressStateProperty = new Property();
                otherAddressStateProperty.Tag = (uint)0x3A62 << 16 | stringTypeHexMask;
                otherAddressStateProperty.Type = PropertyType.String8;
                otherAddressStateProperty.Size = (uint)(otherAddressStateBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressStateProperty.IsReadable = true;
                otherAddressStateProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressStateProperty.ToBytes(), 0, 16);
            }

            if (otherAddressStreet != null)
            {
                byte[] otherAddressStreetBuffer = encoding.GetBytes(otherAddressStreet);
                Independentsoft.IO.StructuredStorage.Stream otherAddressStreetStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A63" + stringTypeMask, otherAddressStreetBuffer);
                directoryEntries.Add(otherAddressStreetStream);

                Property otherAddressStreetProperty = new Property();
                otherAddressStreetProperty.Tag = (uint)0x3A63 << 16 | stringTypeHexMask;
                otherAddressStreetProperty.Type = PropertyType.String8;
                otherAddressStreetProperty.Size = (uint)(otherAddressStreetBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressStreetProperty.IsReadable = true;
                otherAddressStreetProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressStreetProperty.ToBytes(), 0, 16);
            }

            if (otherPhone != null)
            {
                byte[] otherPhoneBuffer = encoding.GetBytes(otherPhone);
                Independentsoft.IO.StructuredStorage.Stream otherPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1F" + stringTypeMask, otherPhoneBuffer);
                directoryEntries.Add(otherPhoneStream);

                Property otherPhoneProperty = new Property();
                otherPhoneProperty.Tag = (uint)0x3A1F << 16 | stringTypeHexMask;
                otherPhoneProperty.Type = PropertyType.String8;
                otherPhoneProperty.Size = (uint)(otherPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                otherPhoneProperty.IsReadable = true;
                otherPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherPhoneProperty.ToBytes(), 0, 16);
            }

            if (pager != null)
            {
                byte[] pagerBuffer = encoding.GetBytes(pager);
                Independentsoft.IO.StructuredStorage.Stream pagerStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A21" + stringTypeMask, pagerBuffer);
                directoryEntries.Add(pagerStream);

                Property pagerProperty = new Property();
                pagerProperty.Tag = (uint)0x3A21 << 16 | stringTypeHexMask;
                pagerProperty.Type = PropertyType.String8;
                pagerProperty.Size = (uint)(pagerBuffer.Length + encoding.GetBytes("\0").Length);
                pagerProperty.IsReadable = true;
                pagerProperty.IsWriteable = true;

                propertiesMemoryStream.Write(pagerProperty.ToBytes(), 0, 16);
            }

            if (personalHomePage != null)
            {
                byte[] personalHomePageBuffer = encoding.GetBytes(personalHomePage);
                Independentsoft.IO.StructuredStorage.Stream personalHomePageStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A50" + stringTypeMask, personalHomePageBuffer);
                directoryEntries.Add(personalHomePageStream);

                Property personalHomePageProperty = new Property();
                personalHomePageProperty.Tag = (uint)0x3A50 << 16 | stringTypeHexMask;
                personalHomePageProperty.Type = PropertyType.String8;
                personalHomePageProperty.Size = (uint)(personalHomePageBuffer.Length + encoding.GetBytes("\0").Length);
                personalHomePageProperty.IsReadable = true;
                personalHomePageProperty.IsWriteable = true;

                propertiesMemoryStream.Write(personalHomePageProperty.ToBytes(), 0, 16);
            }

            if (postalAddress != null)
            {
                byte[] postalAddressBuffer = encoding.GetBytes(postalAddress);
                Independentsoft.IO.StructuredStorage.Stream postalAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A15" + stringTypeMask, postalAddressBuffer);
                directoryEntries.Add(postalAddressStream);

                Property postalAddressProperty = new Property();
                postalAddressProperty.Tag = (uint)0x3A15 << 16 | stringTypeHexMask;
                postalAddressProperty.Type = PropertyType.String8;
                postalAddressProperty.Size = (uint)(postalAddressBuffer.Length + encoding.GetBytes("\0").Length);
                postalAddressProperty.IsReadable = true;
                postalAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(postalAddressProperty.ToBytes(), 0, 16);
            }

            if (businessAddressPostalCode != null)
            {
                NamedProperty businessAddressPostalCodeNamedProperty = new NamedProperty();
                businessAddressPostalCodeNamedProperty.Id = 0x8048;
                businessAddressPostalCodeNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressPostalCodeNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressPostalCodeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressPostalCodeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressPostalCodeBuffer = encoding.GetBytes(businessAddressPostalCode);

                Independentsoft.IO.StructuredStorage.Stream businessAddressPostalCodeStream1 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressPostalCodeBuffer);
                directoryEntries.Add(businessAddressPostalCodeStream1);

                Property businessAddressPostalCodeProperty1 = new Property();
                businessAddressPostalCodeProperty1.Tag = propertyTag;
                businessAddressPostalCodeProperty1.Type = PropertyType.String8;
                businessAddressPostalCodeProperty1.Size = (uint)(businessAddressPostalCodeBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressPostalCodeProperty1.IsReadable = true;
                businessAddressPostalCodeProperty1.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressPostalCodeProperty1.ToBytes(), 0, 16);

                Independentsoft.IO.StructuredStorage.Stream businessAddressPostalCodeStream2 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2A" + stringTypeMask, businessAddressPostalCodeBuffer);
                directoryEntries.Add(businessAddressPostalCodeStream2);

                Property businessAddressPostalCodeProperty2 = new Property();
                businessAddressPostalCodeProperty2.Tag = (uint)0x3A2A << 16 | stringTypeHexMask;
                businessAddressPostalCodeProperty2.Type = PropertyType.String8;
                businessAddressPostalCodeProperty2.Size = (uint)(businessAddressPostalCodeBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressPostalCodeProperty2.IsReadable = true;
                businessAddressPostalCodeProperty2.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressPostalCodeProperty2.ToBytes(), 0, 16);
            }

            if (businessAddressPostOfficeBox != null)
            {
                byte[] businessAddressPostOfficeBoxBuffer = encoding.GetBytes(businessAddressPostOfficeBox);
                Independentsoft.IO.StructuredStorage.Stream businessAddressPostOfficeBoxStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2B" + stringTypeMask, businessAddressPostOfficeBoxBuffer);
                directoryEntries.Add(businessAddressPostOfficeBoxStream);

                Property businessAddressPostOfficeBoxProperty = new Property();
                businessAddressPostOfficeBoxProperty.Tag = (uint)0x3A2B << 16 | stringTypeHexMask;
                businessAddressPostOfficeBoxProperty.Type = PropertyType.String8;
                businessAddressPostOfficeBoxProperty.Size = (uint)(businessAddressPostOfficeBoxBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressPostOfficeBoxProperty.IsReadable = true;
                businessAddressPostOfficeBoxProperty.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressPostOfficeBoxProperty.ToBytes(), 0, 16);
            }

            if (businessAddressState != null)
            {
                NamedProperty businessAddressStateNamedProperty = new NamedProperty();
                businessAddressStateNamedProperty.Id = 0x8047;
                businessAddressStateNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressStateNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressStateNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressStateNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressStateBuffer = encoding.GetBytes(businessAddressState);

                Independentsoft.IO.StructuredStorage.Stream businessAddressStateStream1 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressStateBuffer);
                directoryEntries.Add(businessAddressStateStream1);

                Property businessAddressStateProperty1 = new Property();
                businessAddressStateProperty1.Tag = propertyTag;
                businessAddressStateProperty1.Type = PropertyType.String8;
                businessAddressStateProperty1.Size = (uint)(businessAddressStateBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressStateProperty1.IsReadable = true;
                businessAddressStateProperty1.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressStateProperty1.ToBytes(), 0, 16);

                Independentsoft.IO.StructuredStorage.Stream businessAddressStateStream2 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A28" + stringTypeMask, businessAddressStateBuffer);
                directoryEntries.Add(businessAddressStateStream2);

                Property businessAddressStateProperty2 = new Property();
                businessAddressStateProperty2.Tag = (uint)0x3A28 << 16 | stringTypeHexMask;
                businessAddressStateProperty2.Type = PropertyType.String8;
                businessAddressStateProperty2.Size = (uint)(businessAddressStateBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressStateProperty2.IsReadable = true;
                businessAddressStateProperty2.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressStateProperty2.ToBytes(), 0, 16);
            }

            if (businessAddressStreet != null)
            {
                NamedProperty businessAddressStreetNamedProperty = new NamedProperty();
                businessAddressStreetNamedProperty.Id = 0x8045;
                businessAddressStreetNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressStreetNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressStreetNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressStreetNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressStreetBuffer = encoding.GetBytes(businessAddressStreet);

                Independentsoft.IO.StructuredStorage.Stream businessAddressStreetStream1 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressStreetBuffer);
                directoryEntries.Add(businessAddressStreetStream1);

                Property businessAddressStreetProperty1 = new Property();
                businessAddressStreetProperty1.Tag = propertyTag;
                businessAddressStreetProperty1.Type = PropertyType.String8;
                businessAddressStreetProperty1.Size = (uint)(businessAddressStreetBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressStreetProperty1.IsReadable = true;
                businessAddressStreetProperty1.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressStreetProperty1.ToBytes(), 0, 16);

                Independentsoft.IO.StructuredStorage.Stream businessAddressStreetStream2 = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A29" + stringTypeMask, businessAddressStreetBuffer);
                directoryEntries.Add(businessAddressStreetStream2);

                Property businessAddressStreetProperty2 = new Property();
                businessAddressStreetProperty2.Tag = (uint)0x3A29 << 16 | stringTypeHexMask;
                businessAddressStreetProperty2.Type = PropertyType.String8;
                businessAddressStreetProperty2.Size = (uint)(businessAddressStreetBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressStreetProperty2.IsReadable = true;
                businessAddressStreetProperty2.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressStreetProperty2.ToBytes(), 0, 16);
            }

            if (primaryFax != null)
            {
                byte[] primaryFaxBuffer = encoding.GetBytes(primaryFax);
                Independentsoft.IO.StructuredStorage.Stream primaryFaxStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A23" + stringTypeMask, primaryFaxBuffer);
                directoryEntries.Add(primaryFaxStream);

                Property primaryFaxProperty = new Property();
                primaryFaxProperty.Tag = (uint)0x3A23 << 16 | stringTypeHexMask;
                primaryFaxProperty.Type = PropertyType.String8;
                primaryFaxProperty.Size = (uint)(primaryFaxBuffer.Length + encoding.GetBytes("\0").Length);
                primaryFaxProperty.IsReadable = true;
                primaryFaxProperty.IsWriteable = true;

                propertiesMemoryStream.Write(primaryFaxProperty.ToBytes(), 0, 16);
            }

            if (primaryPhone != null)
            {
                byte[] primaryPhoneBuffer = encoding.GetBytes(primaryPhone);
                Independentsoft.IO.StructuredStorage.Stream primaryPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1A" + stringTypeMask, primaryPhoneBuffer);
                directoryEntries.Add(primaryPhoneStream);

                Property primaryPhoneProperty = new Property();
                primaryPhoneProperty.Tag = (uint)0x3A1A << 16 | stringTypeHexMask;
                primaryPhoneProperty.Type = PropertyType.String8;
                primaryPhoneProperty.Size = (uint)(primaryPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                primaryPhoneProperty.IsReadable = true;
                primaryPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(primaryPhoneProperty.ToBytes(), 0, 16);
            }

            if (profession != null)
            {
                byte[] professionBuffer = encoding.GetBytes(profession);
                Independentsoft.IO.StructuredStorage.Stream professionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A46" + stringTypeMask, professionBuffer);
                directoryEntries.Add(professionStream);

                Property professionProperty = new Property();
                professionProperty.Tag = (uint)0x3A46 << 16 | stringTypeHexMask;
                professionProperty.Type = PropertyType.String8;
                professionProperty.Size = (uint)(professionBuffer.Length + encoding.GetBytes("\0").Length);
                professionProperty.IsReadable = true;
                professionProperty.IsWriteable = true;

                propertiesMemoryStream.Write(professionProperty.ToBytes(), 0, 16);
            }

            if (radioPhone != null)
            {
                byte[] radioPhoneBuffer = encoding.GetBytes(radioPhone);
                Independentsoft.IO.StructuredStorage.Stream radioPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A1D" + stringTypeMask, radioPhoneBuffer);
                directoryEntries.Add(radioPhoneStream);

                Property radioPhoneProperty = new Property();
                radioPhoneProperty.Tag = (uint)0x3A1D << 16 | stringTypeHexMask;
                radioPhoneProperty.Type = PropertyType.String8;
                radioPhoneProperty.Size = (uint)(radioPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                radioPhoneProperty.IsReadable = true;
                radioPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(radioPhoneProperty.ToBytes(), 0, 16);
            }

            if (spouseName != null)
            {
                byte[] spouseNameBuffer = encoding.GetBytes(spouseName);
                Independentsoft.IO.StructuredStorage.Stream spouseNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A48" + stringTypeMask, spouseNameBuffer);
                directoryEntries.Add(spouseNameStream);

                Property spouseNameProperty = new Property();
                spouseNameProperty.Tag = (uint)0x3A48 << 16 | stringTypeHexMask;
                spouseNameProperty.Type = PropertyType.String8;
                spouseNameProperty.Size = (uint)(spouseNameBuffer.Length + encoding.GetBytes("\0").Length);
                spouseNameProperty.IsReadable = true;
                spouseNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(spouseNameProperty.ToBytes(), 0, 16);
            }

            if (surname != null)
            {
                byte[] surnameBuffer = encoding.GetBytes(surname);
                Independentsoft.IO.StructuredStorage.Stream surnameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A11" + stringTypeMask, surnameBuffer);
                directoryEntries.Add(surnameStream);

                Property surnameProperty = new Property();
                surnameProperty.Tag = (uint)0x3A11 << 16 | stringTypeHexMask;
                surnameProperty.Type = PropertyType.String8;
                surnameProperty.Size = (uint)(surnameBuffer.Length + encoding.GetBytes("\0").Length);
                surnameProperty.IsReadable = true;
                surnameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(surnameProperty.ToBytes(), 0, 16);
            }

            if (telex != null)
            {
                byte[] telexBuffer = encoding.GetBytes(telex);
                Independentsoft.IO.StructuredStorage.Stream telexStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A2C" + stringTypeMask, telexBuffer);
                directoryEntries.Add(telexStream);

                Property telexProperty = new Property();
                telexProperty.Tag = (uint)0x3A2C << 16 | stringTypeHexMask;
                telexProperty.Type = PropertyType.String8;
                telexProperty.Size = (uint)(telexBuffer.Length + encoding.GetBytes("\0").Length);
                telexProperty.IsReadable = true;
                telexProperty.IsWriteable = true;

                propertiesMemoryStream.Write(telexProperty.ToBytes(), 0, 16);
            }

            if (title != null)
            {
                byte[] titleBuffer = encoding.GetBytes(title);
                Independentsoft.IO.StructuredStorage.Stream titleStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A17" + stringTypeMask, titleBuffer);
                directoryEntries.Add(titleStream);

                Property titleProperty = new Property();
                titleProperty.Tag = (uint)0x3A17 << 16 | stringTypeHexMask;
                titleProperty.Type = PropertyType.String8;
                titleProperty.Size = (uint)(titleBuffer.Length + encoding.GetBytes("\0").Length);
                titleProperty.IsReadable = true;
                titleProperty.IsWriteable = true;

                propertiesMemoryStream.Write(titleProperty.ToBytes(), 0, 16);
            }

            if (ttyTddPhone != null)
            {
                byte[] ttyTddPhoneBuffer = encoding.GetBytes(ttyTddPhone);
                Independentsoft.IO.StructuredStorage.Stream ttyTddPhoneStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A4B" + stringTypeMask, ttyTddPhoneBuffer);
                directoryEntries.Add(ttyTddPhoneStream);

                Property ttyTddPhoneProperty = new Property();
                ttyTddPhoneProperty.Tag = (uint)0x3A4B << 16 | stringTypeHexMask;
                ttyTddPhoneProperty.Type = PropertyType.String8;
                ttyTddPhoneProperty.Size = (uint)(ttyTddPhoneBuffer.Length + encoding.GetBytes("\0").Length);
                ttyTddPhoneProperty.IsReadable = true;
                ttyTddPhoneProperty.IsWriteable = true;

                propertiesMemoryStream.Write(ttyTddPhoneProperty.ToBytes(), 0, 16);
            }

            if (weddingAnniversary.CompareTo(DateTime.MinValue) > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);
                TimeSpan timeSpan = weddingAnniversary.ToUniversalTime().Subtract(year1601);

                long ticks = timeSpan.Ticks;

                byte[] ticksBytes = BitConverter.GetBytes(ticks);

                Property weddingAnniversaryProperty = new Property();
                weddingAnniversaryProperty.Tag = 0x3A410040;
                weddingAnniversaryProperty.Type = PropertyType.Time;
                weddingAnniversaryProperty.Value = ticksBytes;
                weddingAnniversaryProperty.IsReadable = true;
                weddingAnniversaryProperty.IsWriteable = false;

                propertiesMemoryStream.Write(weddingAnniversaryProperty.ToBytes(), 0, 16);
            }

            if (gender != Gender.None)
            {
                Property genderProperty = new Property();
                genderProperty.Tag = 0x3A4D0002;
                genderProperty.Type = PropertyType.Integer16;
                genderProperty.Value = BitConverter.GetBytes(EnumUtil.ParseGender(gender));
                genderProperty.IsReadable = true;
                genderProperty.IsWriteable = true;

                propertiesMemoryStream.Write(genderProperty.ToBytes(), 0, 16);
            }

            if (selectedMailingAddress != SelectedMailingAddress.None)
            {
                NamedProperty selectedMailingAddressNamedProperty = new NamedProperty();
                selectedMailingAddressNamedProperty.Id = 0x8022;
                selectedMailingAddressNamedProperty.Guid = StandardPropertySet.Address;
                selectedMailingAddressNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, selectedMailingAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(selectedMailingAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0003;

                Property selectedMailingAddressProperty = new Property();
                selectedMailingAddressProperty.Tag = propertyTag;
                selectedMailingAddressProperty.Type = PropertyType.Integer32;
                selectedMailingAddressProperty.Value = BitConverter.GetBytes(EnumUtil.ParseSelectedMailingAddress(selectedMailingAddress));
                selectedMailingAddressProperty.IsReadable = true;
                selectedMailingAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(selectedMailingAddressProperty.ToBytes(), 0, 16);
            }

            if (contactHasPicture)
            {
                NamedProperty contactHasPictureNamedProperty = new NamedProperty();
                contactHasPictureNamedProperty.Id = 0x8015;
                contactHasPictureNamedProperty.Guid = StandardPropertySet.Address;
                contactHasPictureNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, contactHasPictureNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(contactHasPictureNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x000B;

                Property contactHasPictureProperty = new Property();
                contactHasPictureProperty.Tag = propertyTag;
                contactHasPictureProperty.Type = PropertyType.Boolean;
                contactHasPictureProperty.Value = BitConverter.GetBytes(1);
                contactHasPictureProperty.IsReadable = true;
                contactHasPictureProperty.IsWriteable = true;

                propertiesMemoryStream.Write(contactHasPictureProperty.ToBytes(), 0, 16);
            }

            if (fileAs != null)
            {
                NamedProperty fileAsNamedProperty = new NamedProperty();
                fileAsNamedProperty.Id = 0x8005;
                fileAsNamedProperty.Guid = StandardPropertySet.Address;
                fileAsNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, fileAsNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(fileAsNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] fileAsBuffer = encoding.GetBytes(fileAs);
                Independentsoft.IO.StructuredStorage.Stream fileAsStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, fileAsBuffer);
                directoryEntries.Add(fileAsStream);

                Property fileAsProperty = new Property();
                fileAsProperty.Tag = propertyTag;
                fileAsProperty.Type = PropertyType.String8;
                fileAsProperty.Size = (uint)(fileAsBuffer.Length + encoding.GetBytes("\0").Length);
                fileAsProperty.IsReadable = true;
                fileAsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(fileAsProperty.ToBytes(), 0, 16);
            }

            if (instantMessengerAddress != null)
            {
                NamedProperty instantMessengerAddressNamedProperty = new NamedProperty();
                instantMessengerAddressNamedProperty.Id = 0x8062;
                instantMessengerAddressNamedProperty.Guid = StandardPropertySet.Address;
                instantMessengerAddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, instantMessengerAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(instantMessengerAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] instantMessengerAddressBuffer = encoding.GetBytes(instantMessengerAddress);
                Independentsoft.IO.StructuredStorage.Stream instantMessengerAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, instantMessengerAddressBuffer);
                directoryEntries.Add(instantMessengerAddressStream);

                Property instantMessengerAddressProperty = new Property();
                instantMessengerAddressProperty.Tag = propertyTag;
                instantMessengerAddressProperty.Type = PropertyType.String8;
                instantMessengerAddressProperty.Size = (uint)(instantMessengerAddressBuffer.Length + encoding.GetBytes("\0").Length);
                instantMessengerAddressProperty.IsReadable = true;
                instantMessengerAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(instantMessengerAddressProperty.ToBytes(), 0, 16);
            }

            if (internetFreeBusyAddress != null)
            {
                NamedProperty internetFreeBusyAddressNamedProperty = new NamedProperty();
                internetFreeBusyAddressNamedProperty.Id = 0x80D8;
                internetFreeBusyAddressNamedProperty.Guid = StandardPropertySet.Address;
                internetFreeBusyAddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, internetFreeBusyAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(internetFreeBusyAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] internetFreeBusyAddressBuffer = encoding.GetBytes(internetFreeBusyAddress);
                Independentsoft.IO.StructuredStorage.Stream internetFreeBusyAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, internetFreeBusyAddressBuffer);
                directoryEntries.Add(internetFreeBusyAddressStream);

                Property internetFreeBusyAddressProperty = new Property();
                internetFreeBusyAddressProperty.Tag = propertyTag;
                internetFreeBusyAddressProperty.Type = PropertyType.String8;
                internetFreeBusyAddressProperty.Size = (uint)(internetFreeBusyAddressBuffer.Length + encoding.GetBytes("\0").Length);
                internetFreeBusyAddressProperty.IsReadable = true;
                internetFreeBusyAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(internetFreeBusyAddressProperty.ToBytes(), 0, 16);
            }

            if (businessAddress != null)
            {
                NamedProperty businessAddressNamedProperty = new NamedProperty();
                businessAddressNamedProperty.Id = 0x801B;
                businessAddressNamedProperty.Guid = StandardPropertySet.Address;
                businessAddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, businessAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(businessAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] businessAddressBuffer = encoding.GetBytes(businessAddress);
                Independentsoft.IO.StructuredStorage.Stream businessAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, businessAddressBuffer);
                directoryEntries.Add(businessAddressStream);

                Property businessAddressProperty = new Property();
                businessAddressProperty.Tag = propertyTag;
                businessAddressProperty.Type = PropertyType.String8;
                businessAddressProperty.Size = (uint)(businessAddressBuffer.Length + encoding.GetBytes("\0").Length);
                businessAddressProperty.IsReadable = true;
                businessAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(businessAddressProperty.ToBytes(), 0, 16);
            }

            if (homeAddress != null)
            {
                NamedProperty homeAddressNamedProperty = new NamedProperty();
                homeAddressNamedProperty.Id = 0x801A;
                homeAddressNamedProperty.Guid = StandardPropertySet.Address;
                homeAddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, homeAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(homeAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] homeAddressBuffer = encoding.GetBytes(homeAddress);
                Independentsoft.IO.StructuredStorage.Stream homeAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, homeAddressBuffer);
                directoryEntries.Add(homeAddressStream);

                Property homeAddressProperty = new Property();
                homeAddressProperty.Tag = propertyTag;
                homeAddressProperty.Type = PropertyType.String8;
                homeAddressProperty.Size = (uint)(homeAddressBuffer.Length + encoding.GetBytes("\0").Length);
                homeAddressProperty.IsReadable = true;
                homeAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(homeAddressProperty.ToBytes(), 0, 16);
            }


            if (otherAddress != null)
            {
                NamedProperty otherAddressNamedProperty = new NamedProperty();
                otherAddressNamedProperty.Id = 0x801C;
                otherAddressNamedProperty.Guid = StandardPropertySet.Address;
                otherAddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, otherAddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(otherAddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] otherAddressBuffer = encoding.GetBytes(otherAddress);
                Independentsoft.IO.StructuredStorage.Stream otherAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, otherAddressBuffer);
                directoryEntries.Add(otherAddressStream);

                Property otherAddressProperty = new Property();
                otherAddressProperty.Tag = propertyTag;
                otherAddressProperty.Type = PropertyType.String8;
                otherAddressProperty.Size = (uint)(otherAddressBuffer.Length + encoding.GetBytes("\0").Length);
                otherAddressProperty.IsReadable = true;
                otherAddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(otherAddressProperty.ToBytes(), 0, 16);
            }

            if (email1Address != null)
            {
                NamedProperty email1AddressNamedProperty = new NamedProperty();
                email1AddressNamedProperty.Id = 0x8083;
                email1AddressNamedProperty.Guid = StandardPropertySet.Address;
                email1AddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email1AddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email1AddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email1AddressBuffer = encoding.GetBytes(email1Address);
                Independentsoft.IO.StructuredStorage.Stream email1AddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email1AddressBuffer);
                directoryEntries.Add(email1AddressStream);

                Property email1AddressProperty = new Property();
                email1AddressProperty.Tag = propertyTag;
                email1AddressProperty.Type = PropertyType.String8;
                email1AddressProperty.Size = (uint)(email1AddressBuffer.Length + encoding.GetBytes("\0").Length);
                email1AddressProperty.IsReadable = true;
                email1AddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email1AddressProperty.ToBytes(), 0, 16);
            }

            if (email2Address != null)
            {
                NamedProperty email2AddressNamedProperty = new NamedProperty();
                email2AddressNamedProperty.Id = 0x8093;
                email2AddressNamedProperty.Guid = StandardPropertySet.Address;
                email2AddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email2AddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email2AddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email2AddressBuffer = encoding.GetBytes(email2Address);
                Independentsoft.IO.StructuredStorage.Stream email2AddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email2AddressBuffer);
                directoryEntries.Add(email2AddressStream);

                Property email2AddressProperty = new Property();
                email2AddressProperty.Tag = propertyTag;
                email2AddressProperty.Type = PropertyType.String8;
                email2AddressProperty.Size = (uint)(email2AddressBuffer.Length + encoding.GetBytes("\0").Length);
                email2AddressProperty.IsReadable = true;
                email2AddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email2AddressProperty.ToBytes(), 0, 16);
            }

            if (email3Address != null)
            {
                NamedProperty email3AddressNamedProperty = new NamedProperty();
                email3AddressNamedProperty.Id = 0x80A3;
                email3AddressNamedProperty.Guid = StandardPropertySet.Address;
                email3AddressNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email3AddressNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email3AddressNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email3AddressBuffer = encoding.GetBytes(email3Address);
                Independentsoft.IO.StructuredStorage.Stream email3AddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email3AddressBuffer);
                directoryEntries.Add(email3AddressStream);

                Property email3AddressProperty = new Property();
                email3AddressProperty.Tag = propertyTag;
                email3AddressProperty.Type = PropertyType.String8;
                email3AddressProperty.Size = (uint)(email3AddressBuffer.Length + encoding.GetBytes("\0").Length);
                email3AddressProperty.IsReadable = true;
                email3AddressProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email3AddressProperty.ToBytes(), 0, 16);
            }

            if (email1DisplayName != null)
            {
                NamedProperty email1DisplayNameNamedProperty = new NamedProperty();
                email1DisplayNameNamedProperty.Id = 0x8084;
                email1DisplayNameNamedProperty.Guid = StandardPropertySet.Address;
                email1DisplayNameNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email1DisplayNameNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email1DisplayNameNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email1DisplayNameBuffer = encoding.GetBytes(email1DisplayName);
                Independentsoft.IO.StructuredStorage.Stream email1DisplayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email1DisplayNameBuffer);
                directoryEntries.Add(email1DisplayNameStream);

                Property email1DisplayNameProperty = new Property();
                email1DisplayNameProperty.Tag = propertyTag;
                email1DisplayNameProperty.Type = PropertyType.String8;
                email1DisplayNameProperty.Size = (uint)(email1DisplayNameBuffer.Length + encoding.GetBytes("\0").Length);
                email1DisplayNameProperty.IsReadable = true;
                email1DisplayNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email1DisplayNameProperty.ToBytes(), 0, 16);
            }

            if (email2DisplayName != null)
            {
                NamedProperty email2DisplayNameNamedProperty = new NamedProperty();
                email2DisplayNameNamedProperty.Id = 0x8094;
                email2DisplayNameNamedProperty.Guid = StandardPropertySet.Address;
                email2DisplayNameNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email2DisplayNameNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email2DisplayNameNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email2DisplayNameBuffer = encoding.GetBytes(email2DisplayName);
                Independentsoft.IO.StructuredStorage.Stream email2DisplayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email2DisplayNameBuffer);
                directoryEntries.Add(email2DisplayNameStream);

                Property email2DisplayNameProperty = new Property();
                email2DisplayNameProperty.Tag = propertyTag;
                email2DisplayNameProperty.Type = PropertyType.String8;
                email2DisplayNameProperty.Size = (uint)(email2DisplayNameBuffer.Length + encoding.GetBytes("\0").Length);
                email2DisplayNameProperty.IsReadable = true;
                email2DisplayNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email2DisplayNameProperty.ToBytes(), 0, 16);
            }

            if (email3DisplayName != null)
            {
                NamedProperty email3DisplayNameNamedProperty = new NamedProperty();
                email3DisplayNameNamedProperty.Id = 0x80A4;
                email3DisplayNameNamedProperty.Guid = StandardPropertySet.Address;
                email3DisplayNameNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email3DisplayNameNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email3DisplayNameNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email3DisplayNameBuffer = encoding.GetBytes(email3DisplayName);
                Independentsoft.IO.StructuredStorage.Stream email3DisplayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email3DisplayNameBuffer);
                directoryEntries.Add(email3DisplayNameStream);

                Property email3DisplayNameProperty = new Property();
                email3DisplayNameProperty.Tag = propertyTag;
                email3DisplayNameProperty.Type = PropertyType.String8;
                email3DisplayNameProperty.Size = (uint)(email3DisplayNameBuffer.Length + encoding.GetBytes("\0").Length);
                email3DisplayNameProperty.IsReadable = true;
                email3DisplayNameProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email3DisplayNameProperty.ToBytes(), 0, 16);
            }

            if (email1DisplayAs != null)
            {
                NamedProperty email1DisplayAsNamedProperty = new NamedProperty();
                email1DisplayAsNamedProperty.Id = 0x8080;
                email1DisplayAsNamedProperty.Guid = StandardPropertySet.Address;
                email1DisplayAsNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email1DisplayAsNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email1DisplayAsNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email1DisplayAsBuffer = encoding.GetBytes(email1DisplayAs);
                Independentsoft.IO.StructuredStorage.Stream email1DisplayAsStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email1DisplayAsBuffer);
                directoryEntries.Add(email1DisplayAsStream);

                Property email1DisplayAsProperty = new Property();
                email1DisplayAsProperty.Tag = propertyTag;
                email1DisplayAsProperty.Type = PropertyType.String8;
                email1DisplayAsProperty.Size = (uint)(email1DisplayAsBuffer.Length + encoding.GetBytes("\0").Length);
                email1DisplayAsProperty.IsReadable = true;
                email1DisplayAsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email1DisplayAsProperty.ToBytes(), 0, 16);
            }

            if (email2DisplayAs != null)
            {
                NamedProperty email2DisplayAsNamedProperty = new NamedProperty();
                email2DisplayAsNamedProperty.Id = 0x8090;
                email2DisplayAsNamedProperty.Guid = StandardPropertySet.Address;
                email2DisplayAsNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email2DisplayAsNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email2DisplayAsNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email2DisplayAsBuffer = encoding.GetBytes(email2DisplayAs);
                Independentsoft.IO.StructuredStorage.Stream email2DisplayAsStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email2DisplayAsBuffer);
                directoryEntries.Add(email2DisplayAsStream);

                Property email2DisplayAsProperty = new Property();
                email2DisplayAsProperty.Tag = propertyTag;
                email2DisplayAsProperty.Type = PropertyType.String8;
                email2DisplayAsProperty.Size = (uint)(email2DisplayAsBuffer.Length + encoding.GetBytes("\0").Length);
                email2DisplayAsProperty.IsReadable = true;
                email2DisplayAsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email2DisplayAsProperty.ToBytes(), 0, 16);
            }


            if (email3DisplayAs != null)
            {
                NamedProperty email3DisplayAsNamedProperty = new NamedProperty();
                email3DisplayAsNamedProperty.Id = 0x80A0;
                email3DisplayAsNamedProperty.Guid = StandardPropertySet.Address;
                email3DisplayAsNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email3DisplayAsNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email3DisplayAsNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email3DisplayAsBuffer = encoding.GetBytes(email3DisplayAs);
                Independentsoft.IO.StructuredStorage.Stream email3DisplayAsStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email3DisplayAsBuffer);
                directoryEntries.Add(email3DisplayAsStream);

                Property email3DisplayAsProperty = new Property();
                email3DisplayAsProperty.Tag = propertyTag;
                email3DisplayAsProperty.Type = PropertyType.String8;
                email3DisplayAsProperty.Size = (uint)(email3DisplayAsBuffer.Length + encoding.GetBytes("\0").Length);
                email3DisplayAsProperty.IsReadable = true;
                email3DisplayAsProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email3DisplayAsProperty.ToBytes(), 0, 16);
            }

            if (email1Type != null)
            {
                NamedProperty email1TypeNamedProperty = new NamedProperty();
                email1TypeNamedProperty.Id = 0x8082;
                email1TypeNamedProperty.Guid = StandardPropertySet.Address;
                email1TypeNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email1TypeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email1TypeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email1TypeBuffer = encoding.GetBytes(email1Type);
                Independentsoft.IO.StructuredStorage.Stream email1TypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email1TypeBuffer);
                directoryEntries.Add(email1TypeStream);

                Property email1TypeProperty = new Property();
                email1TypeProperty.Tag = propertyTag;
                email1TypeProperty.Type = PropertyType.String8;
                email1TypeProperty.Size = (uint)(email1TypeBuffer.Length + encoding.GetBytes("\0").Length);
                email1TypeProperty.IsReadable = true;
                email1TypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email1TypeProperty.ToBytes(), 0, 16);
            }

            if (email2Type != null)
            {
                NamedProperty email2TypeNamedProperty = new NamedProperty();
                email2TypeNamedProperty.Id = 0x8092;
                email2TypeNamedProperty.Guid = StandardPropertySet.Address;
                email2TypeNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email2TypeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email2TypeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email2TypeBuffer = encoding.GetBytes(email2Type);
                Independentsoft.IO.StructuredStorage.Stream email2TypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email2TypeBuffer);
                directoryEntries.Add(email2TypeStream);

                Property email2TypeProperty = new Property();
                email2TypeProperty.Tag = propertyTag;
                email2TypeProperty.Type = PropertyType.String8;
                email2TypeProperty.Size = (uint)(email2TypeBuffer.Length + encoding.GetBytes("\0").Length);
                email2TypeProperty.IsReadable = true;
                email2TypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email2TypeProperty.ToBytes(), 0, 16);
            }

            if (email3Type != null)
            {
                NamedProperty email3TypeNamedProperty = new NamedProperty();
                email3TypeNamedProperty.Id = 0x80A2;
                email3TypeNamedProperty.Guid = StandardPropertySet.Address;
                email3TypeNamedProperty.Type = NamedPropertyType.String;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email3TypeNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email3TypeNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                byte[] email3TypeBuffer = encoding.GetBytes(email3Type);
                Independentsoft.IO.StructuredStorage.Stream email3TypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email3TypeBuffer);
                directoryEntries.Add(email3TypeStream);

                Property email3TypeProperty = new Property();
                email3TypeProperty.Tag = propertyTag;
                email3TypeProperty.Type = PropertyType.String8;
                email3TypeProperty.Size = (uint)(email3TypeBuffer.Length + encoding.GetBytes("\0").Length);
                email3TypeProperty.IsReadable = true;
                email3TypeProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email3TypeProperty.ToBytes(), 0, 16);
            }

            if (email1EntryId != null)
            {
                NamedProperty email1EntryIdNamedProperty = new NamedProperty();
                email1EntryIdNamedProperty.Id = 0x8085;
                email1EntryIdNamedProperty.Guid = StandardPropertySet.Address;
                email1EntryIdNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email1EntryIdNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email1EntryIdNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0102;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                Independentsoft.IO.StructuredStorage.Stream email1EntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email1EntryId);
                directoryEntries.Add(email1EntryIdStream);

                Property email1EntryIdProperty = new Property();
                email1EntryIdProperty.Tag = propertyTag;
                email1EntryIdProperty.Type = PropertyType.Integer32;
                email1EntryIdProperty.Size = (uint)email1EntryId.Length;
                email1EntryIdProperty.IsReadable = true;
                email1EntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email1EntryIdProperty.ToBytes(), 0, 16);
            }

            if (email2EntryId != null)
            {
                NamedProperty email2EntryIdNamedProperty = new NamedProperty();
                email2EntryIdNamedProperty.Id = 0x8095;
                email2EntryIdNamedProperty.Guid = StandardPropertySet.Address;
                email2EntryIdNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email2EntryIdNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email2EntryIdNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0102;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                Independentsoft.IO.StructuredStorage.Stream email2EntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email2EntryId);
                directoryEntries.Add(email2EntryIdStream);

                Property email2EntryIdProperty = new Property();
                email2EntryIdProperty.Tag = propertyTag;
                email2EntryIdProperty.Type = PropertyType.Integer32;
                email2EntryIdProperty.Size = (uint)email2EntryId.Length;
                email2EntryIdProperty.IsReadable = true;
                email2EntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email2EntryIdProperty.ToBytes(), 0, 16);
            }

            if (email3EntryId != null)
            {
                NamedProperty email3EntryIdNamedProperty = new NamedProperty();
                email3EntryIdNamedProperty.Id = 0x80A5;
                email3EntryIdNamedProperty.Guid = StandardPropertySet.Address;
                email3EntryIdNamedProperty.Type = NamedPropertyType.Numerical;

                int propertyIndex = Util.IndexOfNamedProperty(namedProperties, email3EntryIdNamedProperty);

                if (propertyIndex == -1)
                {
                    namedProperties.Add(email3EntryIdNamedProperty);
                    propertyIndex = namedProperties.Count - 1;
                }

                uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | 0x0102;
                string propertyTagString = String.Format("{0:X8}", propertyTag);

                Independentsoft.IO.StructuredStorage.Stream email3EntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, email3EntryId);
                directoryEntries.Add(email3EntryIdStream);

                Property email3EntryIdProperty = new Property();
                email3EntryIdProperty.Tag = propertyTag;
                email3EntryIdProperty.Type = PropertyType.Integer32;
                email3EntryIdProperty.Size = (uint)email3EntryId.Length;
                email3EntryIdProperty.IsReadable = true;
                email3EntryIdProperty.IsWriteable = true;

                propertiesMemoryStream.Write(email3EntryIdProperty.ToBytes(), 0, 16);
            }

            //Extended properties
            for (int e = 0; e < extendedProperties.Count; e++)
            {
                if (extendedProperties[e].Value != null)
                {
                    NamedProperty namedProperty = new NamedProperty();

                    if (extendedProperties[e].Tag is ExtendedPropertyId)
                    {
                        ExtendedPropertyId tag = (ExtendedPropertyId)extendedProperties[e].Tag;

                        namedProperty.Id = (uint)tag.Id;
                        namedProperty.Guid = tag.Guid;
                        namedProperty.Type = NamedPropertyType.Numerical;
                    }
                    else
                    {
                        ExtendedPropertyName tag = (ExtendedPropertyName)extendedProperties[e].Tag;

                        namedProperty.Name = tag.Name;
                        namedProperty.Guid = tag.Guid;
                        namedProperty.Type = NamedPropertyType.String;
                    }

                    if (Util.IndexOfNamedProperty(namedProperties, namedProperty) == -1)
                    {
                         namedProperties.Add(namedProperty);
                         int propertyIndex = namedProperties.Count - 1;

                        uint propertyTag = (uint)(0x8000 + propertyIndex) << 16 | Util.GetTypeHexMask(extendedProperties[e].Tag.Type);

                        if (extendedProperties[e].Tag.Type == PropertyType.Boolean || extendedProperties[e].Tag.Type == PropertyType.Integer16 || extendedProperties[e].Tag.Type == PropertyType.Integer32 || extendedProperties[e].Tag.Type == PropertyType.Integer64
                            || extendedProperties[e].Tag.Type == PropertyType.Floating32 || extendedProperties[e].Tag.Type == PropertyType.Floating64 || extendedProperties[e].Tag.Type == PropertyType.FloatingTime || extendedProperties[e].Tag.Type == PropertyType.Time)
                        {
                            Property extendedProperty = new Property();
                            extendedProperty.Tag = propertyTag;
                            extendedProperty.Type = extendedProperties[e].Tag.Type;
                            extendedProperty.Value = extendedProperties[e].Value;
                            extendedProperty.IsReadable = true;
                            extendedProperty.IsWriteable = true;

                            propertiesMemoryStream.Write(extendedProperty.ToBytes(), 0, 16);
                        }
                        else if (extendedProperties[e].Tag.Type == PropertyType.MultipleCurrency || extendedProperties[e].Tag.Type == PropertyType.MultipleFloating32 || extendedProperties[e].Tag.Type == PropertyType.MultipleFloating64 || extendedProperties[e].Tag.Type == PropertyType.MultipleFloatingTime
                                || extendedProperties[e].Tag.Type == PropertyType.MultipleGuid || extendedProperties[e].Tag.Type == PropertyType.MultipleInteger16 || extendedProperties[e].Tag.Type == PropertyType.MultipleInteger32
                                || extendedProperties[e].Tag.Type == PropertyType.MultipleInteger64 || extendedProperties[e].Tag.Type == PropertyType.MultipleString || extendedProperties[e].Tag.Type == PropertyType.MultipleString8 || extendedProperties[e].Tag.Type == PropertyType.MultipleTime)
                        {
                            //Not implemented. Skip it. Implementation is same as Keywords property or MultiBinary
                        }
                        else if (extendedProperties[e].Tag.Type == PropertyType.MultipleBinary)
                        {
                            string propertyTagString = String.Format("{0:X8}", propertyTag);

                            MemoryStream memoryStream = new MemoryStream();

                            int valueCount = BitConverter.ToInt32(extendedProperties[e].Value, 0);

                            int[] positions = new int[valueCount + 1];

                            for (int i = 0; i < valueCount; i++)
                            {
                                positions[i] = BitConverter.ToInt32(extendedProperties[e].Value, 4 + i * 4);
                            }

                            positions[valueCount] = extendedProperties[e].Value.Length;

                            for (int i = 0; i < positions.Length - 1; i++)
                            {
                                long size = (long)(positions[i + 1] - positions[i]);
                                byte[] sizeBuffer = BitConverter.GetBytes(size);

                                memoryStream.Write(sizeBuffer, 0, 8);

                                byte[] valueBuffer = new byte[(int)size];
                                System.Array.Copy(extendedProperties[e].Value, positions[i], valueBuffer, 0, (int)size);

                                string streamName = "__substg1.0_" + propertyTagString + "-" + String.Format("{0:X8}", i);

                                Independentsoft.IO.StructuredStorage.Stream valueStream = new Independentsoft.IO.StructuredStorage.Stream(streamName, valueBuffer);
                                directoryEntries.Add(valueStream);
                            }

                            byte[] propertyLengthStreamBuffer = memoryStream.ToArray();

                            Independentsoft.IO.StructuredStorage.Stream propertyLengthStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, propertyLengthStreamBuffer);
                            directoryEntries.Add(propertyLengthStream);

                            Property lengthProperty = new Property();
                            lengthProperty.Tag = propertyTag;
                            lengthProperty.Type = PropertyType.MultipleBinary;
                            lengthProperty.Size = (uint)propertyLengthStreamBuffer.Length;
                            lengthProperty.IsReadable = true;
                            lengthProperty.IsWriteable = true;

                            propertiesMemoryStream.Write(lengthProperty.ToBytes(), 0, 16);
                        }
                        else
                        {
                            byte[] valueBuffer = extendedProperties[e].Value;

                            if (valueBuffer != null && extendedProperties[e].Tag.Type == PropertyType.String && !(encoding == System.Text.Encoding.Unicode || encoding is System.Text.UnicodeEncoding))
                            {
                                string stringValue = System.Text.Encoding.Unicode.GetString(valueBuffer, 0, valueBuffer.Length);
                                valueBuffer = encoding.GetBytes(stringValue);

                                propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                            }
                            else if (valueBuffer != null && extendedProperties[e].Tag.Type == PropertyType.String8 && (encoding == System.Text.Encoding.Unicode || encoding is System.Text.UnicodeEncoding))
                            {
                                string stringValue = System.Text.Encoding.UTF8.GetString(valueBuffer, 0, valueBuffer.Length);
                                valueBuffer = encoding.GetBytes(stringValue);

                                propertyTag = (uint)(0x8000 + propertyIndex) << 16 | stringTypeHexMask;
                            }

                            String propertyTagString = String.Format("{0:X8}", propertyTag);

                            Independentsoft.IO.StructuredStorage.Stream extendedPropertyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_" + propertyTagString, valueBuffer);
                            directoryEntries.Add(extendedPropertyStream);

                            Property extendedProperty = new Property();
                            extendedProperty.Tag = propertyTag;

                            if (extendedProperties[e].Tag.Type == PropertyType.Binary || extendedProperties[e].Tag.Type == PropertyType.Object)
                            {
                                extendedProperty.Type = PropertyType.Integer32;
                            }
                            else
                            {
                                extendedProperty.Type = PropertyType.String8;
                            }

                            extendedProperty.Size = Util.GetPropertySize(valueBuffer, extendedProperties[e].Tag.Type, encoding);
                            extendedProperty.IsReadable = true;
                            extendedProperty.IsWriteable = true;

                            propertiesMemoryStream.Write(extendedProperty.ToBytes(), 0, 16);
                        }
                    }
                }
            }

            Independentsoft.IO.StructuredStorage.Stream propertiesStream = new Independentsoft.IO.StructuredStorage.Stream("__properties_version1.0", propertiesMemoryStream.ToArray());

            directoryEntries.Add(propertiesStream);

            //Recipients
            for (int i = 0; i < recipients.Count; i++)
            {
                string recipientStorageName = String.Format("__recip_version1.0_#{0:X8}", i);
                Independentsoft.IO.StructuredStorage.Storage recipientStorage = new Independentsoft.IO.StructuredStorage.Storage(recipientStorageName);

                Recipient recipient = recipients[i];

                MemoryStream recipentPropertiesMemoryStream = new MemoryStream();

                //empty header
                byte[] recipentPropertiesHeader = new byte[8];
                recipentPropertiesMemoryStream.Write(recipentPropertiesHeader, 0, 8);


                //PR_ROWID must present
                Property rowIdProperty = new Property();
                rowIdProperty.Tag = 0x30000003;
                rowIdProperty.Type = PropertyType.Integer32;
                rowIdProperty.Value = BitConverter.GetBytes(i);
                rowIdProperty.IsReadable = true;
                rowIdProperty.IsWriteable = true;

                recipentPropertiesMemoryStream.Write(rowIdProperty.ToBytes(), 0, 16);

                if (recipient.DisplayType != DisplayType.None)
                {
                    Property displayTypeProperty = new Property();
                    displayTypeProperty.Tag = 0x39000003;
                    displayTypeProperty.Type = PropertyType.Integer32;
                    displayTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseDisplayType(recipient.DisplayType));
                    displayTypeProperty.IsReadable = true;
                    displayTypeProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(displayTypeProperty.ToBytes(), 0, 16);
                }

                if (recipient.ObjectType != ObjectType.None)
                {
                    Property objectTypeProperty = new Property();
                    objectTypeProperty.Tag = 0x0FFE0003;
                    objectTypeProperty.Type = PropertyType.Integer32;
                    objectTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseObjectType(recipient.ObjectType));
                    objectTypeProperty.IsReadable = true;
                    objectTypeProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(objectTypeProperty.ToBytes(), 0, 16);
                }

                if (recipient.RecipientType != RecipientType.None)
                {
                    Property recipientTypeProperty = new Property();
                    recipientTypeProperty.Tag = 0x0C150003;
                    recipientTypeProperty.Type = PropertyType.Integer32;
                    recipientTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseRecipientType(recipient.RecipientType));
                    recipientTypeProperty.IsReadable = true;
                    recipientTypeProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(recipientTypeProperty.ToBytes(), 0, 16);
                }

                if (recipient.DisplayName != null)
                {
                    byte[] displayNameBuffer = encoding.GetBytes(recipient.DisplayName);

                    Independentsoft.IO.StructuredStorage.Stream displayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3001" + stringTypeMask, displayNameBuffer);
                    recipientStorage.DirectoryEntries.Add(displayNameStream);

                    Property displayNameProperty = new Property();
                    displayNameProperty.Tag = (uint)0x3001 << 16 | stringTypeHexMask;
                    displayNameProperty.Type = PropertyType.String8;
                    displayNameProperty.Size = (uint)(displayNameBuffer.Length + encoding.GetBytes("\0").Length);
                    displayNameProperty.IsReadable = true;
                    displayNameProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(displayNameProperty.ToBytes(), 0, 16);

                    Independentsoft.IO.StructuredStorage.Stream recipientDisplayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_5FF6" + stringTypeMask, displayNameBuffer);
                    recipientStorage.DirectoryEntries.Add(recipientDisplayNameStream);

                    Property recipientDisplayNameProperty = new Property();
                    recipientDisplayNameProperty.Tag = (uint)0x5FF6 << 16 | stringTypeHexMask;
                    recipientDisplayNameProperty.Type = PropertyType.String8;
                    recipientDisplayNameProperty.Size = (uint)(displayNameBuffer.Length + encoding.GetBytes("\0").Length);
                    recipientDisplayNameProperty.IsReadable = true;
                    recipientDisplayNameProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(recipientDisplayNameProperty.ToBytes(), 0, 16);
                }

                if (recipient.EmailAddress != null)
                {
                    byte[] emailAddressBuffer = encoding.GetBytes(recipient.EmailAddress);
                    Independentsoft.IO.StructuredStorage.Stream emailAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3003" + stringTypeMask, emailAddressBuffer);
                    recipientStorage.DirectoryEntries.Add(emailAddressStream);

                    Property emailAddressProperty = new Property();
                    emailAddressProperty.Tag = (uint)0x3003 << 16 | stringTypeHexMask;
                    emailAddressProperty.Type = PropertyType.String8;
                    emailAddressProperty.Size = (uint)(emailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                    emailAddressProperty.IsReadable = true;
                    emailAddressProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(emailAddressProperty.ToBytes(), 0, 16);
                }

                if (recipient.AddressType != null)
                {
                    byte[] addressTypeBuffer = encoding.GetBytes(recipient.AddressType);
                    Independentsoft.IO.StructuredStorage.Stream addressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3002" + stringTypeMask, addressTypeBuffer);
                    recipientStorage.DirectoryEntries.Add(addressTypeStream);

                    Property addressTypeProperty = new Property();
                    addressTypeProperty.Tag = (uint)0x3002 << 16 | stringTypeHexMask;
                    addressTypeProperty.Type = PropertyType.String8;
                    addressTypeProperty.Size = (uint)(addressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                    addressTypeProperty.IsReadable = true;
                    addressTypeProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(addressTypeProperty.ToBytes(), 0, 16);
                }

                if (recipient.EntryId != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream entryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0FFF0102", recipient.EntryId);
                    recipientStorage.DirectoryEntries.Add(entryIdStream);

                    Property entryIdProperty = new Property();
                    entryIdProperty.Tag = 0x0FFF0102;
                    entryIdProperty.Type = PropertyType.Binary;
                    entryIdProperty.Size = (uint)recipient.EntryId.Length;
                    entryIdProperty.IsReadable = true;
                    entryIdProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(entryIdProperty.ToBytes(), 0, 16);

                    Independentsoft.IO.StructuredStorage.Stream recipientEntryIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_5FF70102", recipient.EntryId);
                    recipientStorage.DirectoryEntries.Add(recipientEntryIdStream);

                    Property recipientEntryIdProperty = new Property();
                    recipientEntryIdProperty.Tag = 0x5FF70102;
                    recipientEntryIdProperty.Type = PropertyType.Binary;
                    recipientEntryIdProperty.Size = (uint)recipient.EntryId.Length;
                    recipientEntryIdProperty.IsReadable = true;
                    recipientEntryIdProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(recipientEntryIdProperty.ToBytes(), 0, 16);
                }

                if (recipient.SearchKey != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream searchKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_300B0102", recipient.SearchKey);
                    recipientStorage.DirectoryEntries.Add(searchKeyStream);

                    Property searchKeyProperty = new Property();
                    searchKeyProperty.Tag = 0x300B0102;
                    searchKeyProperty.Type = PropertyType.Binary;
                    searchKeyProperty.Size = (uint)recipient.SearchKey.Length;
                    searchKeyProperty.IsReadable = true;
                    searchKeyProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(searchKeyProperty.ToBytes(), 0, 16);
                }

                if (recipient.InstanceKey != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream instanceKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0FF60102", recipient.InstanceKey);
                    recipientStorage.DirectoryEntries.Add(instanceKeyStream);

                    Property instanceKeyProperty = new Property();
                    instanceKeyProperty.Tag = 0x0FF60102;
                    instanceKeyProperty.Type = PropertyType.Binary;
                    instanceKeyProperty.Size = (uint)recipient.InstanceKey.Length;
                    instanceKeyProperty.IsReadable = true;
                    instanceKeyProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(instanceKeyProperty.ToBytes(), 0, 16);
                }

                if (recipient.Responsibility)
                {
                    Property responsibilityProperty = new Property();
                    responsibilityProperty.Tag = 0x0E0F000B;
                    responsibilityProperty.Type = PropertyType.Boolean;
                    responsibilityProperty.Value = BitConverter.GetBytes(1);
                    responsibilityProperty.IsReadable = true;
                    responsibilityProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(responsibilityProperty.ToBytes(), 0, 16);
                }

                if (recipient.SendRichInfo)
                {
                    Property sendRichInfoProperty = new Property();
                    sendRichInfoProperty.Tag = 0x3A40000B;
                    sendRichInfoProperty.Type = PropertyType.Boolean;
                    sendRichInfoProperty.Value = BitConverter.GetBytes(1);
                    sendRichInfoProperty.IsReadable = true;
                    sendRichInfoProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(sendRichInfoProperty.ToBytes(), 0, 16);
                }

                if (recipient.SendInternetEncoding > 0)
                {
                    Property sendInternetEncodingProperty = new Property();
                    sendInternetEncodingProperty.Tag = 0x3A710003;
                    sendInternetEncodingProperty.Type = PropertyType.Integer32;
                    sendInternetEncodingProperty.Value = BitConverter.GetBytes(recipient.SendInternetEncoding);
                    sendInternetEncodingProperty.IsReadable = true;
                    sendInternetEncodingProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(sendInternetEncodingProperty.ToBytes(), 0, 16);
                }

                if (recipient.SmtpAddress != null)
                {
                    byte[] smtpAddressBuffer = encoding.GetBytes(recipient.SmtpAddress);
                    Independentsoft.IO.StructuredStorage.Stream smtpAddressStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_39FE" + stringTypeMask, smtpAddressBuffer);
                    recipientStorage.DirectoryEntries.Add(smtpAddressStream);

                    Property smtpAddressProperty = new Property();
                    smtpAddressProperty.Tag = (uint)0x39FE << 16 | stringTypeHexMask;
                    smtpAddressProperty.Type = PropertyType.String8;
                    smtpAddressProperty.Size = (uint)(smtpAddressBuffer.Length + encoding.GetBytes("\0").Length);
                    smtpAddressProperty.IsReadable = true;
                    smtpAddressProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(smtpAddressProperty.ToBytes(), 0, 16);
                }

                if (recipient.DisplayName7Bit != null)
                {
                    byte[] displayName7BitBuffer = encoding.GetBytes(recipient.DisplayName7Bit);
                    Independentsoft.IO.StructuredStorage.Stream displayName7BitStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_39FF" + stringTypeMask, displayName7BitBuffer);
                    recipientStorage.DirectoryEntries.Add(displayName7BitStream);

                    Property displayName7BitProperty = new Property();
                    displayName7BitProperty.Tag = (uint)0x39FF << 16 | stringTypeHexMask;
                    displayName7BitProperty.Type = PropertyType.String8;
                    displayName7BitProperty.Size = (uint)(displayName7BitBuffer.Length + encoding.GetBytes("\0").Length);
                    displayName7BitProperty.IsReadable = true;
                    displayName7BitProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(displayName7BitProperty.ToBytes(), 0, 16);
                }

                if (recipient.TransmitableDisplayName != null)
                {
                    byte[] transmitableDisplayNameBuffer = encoding.GetBytes(recipient.TransmitableDisplayName);
                    Independentsoft.IO.StructuredStorage.Stream transmitableDisplayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3A20" + stringTypeMask, transmitableDisplayNameBuffer);
                    recipientStorage.DirectoryEntries.Add(transmitableDisplayNameStream);

                    Property transmitableDisplayNameProperty = new Property();
                    transmitableDisplayNameProperty.Tag = (uint)0x3A20 << 16 | stringTypeHexMask;
                    transmitableDisplayNameProperty.Type = PropertyType.String8;
                    transmitableDisplayNameProperty.Size = (uint)(transmitableDisplayNameBuffer.Length + encoding.GetBytes("\0").Length);
                    transmitableDisplayNameProperty.IsReadable = true;
                    transmitableDisplayNameProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(transmitableDisplayNameProperty.ToBytes(), 0, 16);
                }

                if (recipient.OriginatingAddressType != null)
                {
                    byte[] originatingAddressTypeBuffer = encoding.GetBytes(recipient.OriginatingAddressType);
                    Independentsoft.IO.StructuredStorage.Stream originatingAddressTypeStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_403D" + stringTypeMask, originatingAddressTypeBuffer);
                    recipientStorage.DirectoryEntries.Add(originatingAddressTypeStream);

                    Property originatingAddressTypeProperty = new Property();
                    originatingAddressTypeProperty.Tag = (uint)0x403D << 16 | stringTypeHexMask;
                    originatingAddressTypeProperty.Type = PropertyType.String8;
                    originatingAddressTypeProperty.Size = (uint)(originatingAddressTypeBuffer.Length + encoding.GetBytes("\0").Length);
                    originatingAddressTypeProperty.IsReadable = true;
                    originatingAddressTypeProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(originatingAddressTypeProperty.ToBytes(), 0, 16);
                }

                if (recipient.OriginatingEmailAddress != null)
                {
                    byte[] originatingEmailAddressBuffer = encoding.GetBytes(recipient.OriginatingEmailAddress);
                    Independentsoft.IO.StructuredStorage.Stream originatingEmailAddressBufferStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_403E" + stringTypeMask, originatingEmailAddressBuffer);
                    recipientStorage.DirectoryEntries.Add(originatingEmailAddressBufferStream);

                    Property originatingEmailAddressBufferProperty = new Property();
                    originatingEmailAddressBufferProperty.Tag = (uint)0x403E << 16 | stringTypeHexMask;
                    originatingEmailAddressBufferProperty.Type = PropertyType.String8;
                    originatingEmailAddressBufferProperty.Size = (uint)(originatingEmailAddressBuffer.Length + encoding.GetBytes("\0").Length);
                    originatingEmailAddressBufferProperty.IsReadable = true;
                    originatingEmailAddressBufferProperty.IsWriteable = true;

                    recipentPropertiesMemoryStream.Write(originatingEmailAddressBufferProperty.ToBytes(), 0, 16);
                }

                Independentsoft.IO.StructuredStorage.Stream recipientPropertiesStream = new Independentsoft.IO.StructuredStorage.Stream("__properties_version1.0", recipentPropertiesMemoryStream.ToArray());

                recipientStorage.DirectoryEntries.Add(recipientPropertiesStream);

                directoryEntries.Add(recipientStorage);
            }

            //Attachments
            for (int i = 0; i < attachments.Count; i++)
            {
                string attachmentStorageName = String.Format("__attach_version1.0_#{0:X8}", i);
                Independentsoft.IO.StructuredStorage.Storage attachmentStorage = new Independentsoft.IO.StructuredStorage.Storage(attachmentStorageName);

                Attachment attachment = attachments[i];

                MemoryStream attachmentPropertiesMemoryStream = new MemoryStream();

                //empty header
                byte[] attachmentPropertiesHeader = new byte[8];
                attachmentPropertiesMemoryStream.Write(attachmentPropertiesHeader, 0, 8);

                //PR_ATTACH_NUM
                Property attachmentIdProperty = new Property();
                attachmentIdProperty.Tag = 0x0E210003;
                attachmentIdProperty.Type = PropertyType.Integer32;
                attachmentIdProperty.Value = BitConverter.GetBytes(i);
                attachmentIdProperty.IsReadable = true;
                attachmentIdProperty.IsWriteable = false;

                attachmentPropertiesMemoryStream.Write(attachmentIdProperty.ToBytes(), 0, 16);

                //PR_ATTACHMENT_LINKID
                Property attachmentLinkIdProperty = new Property();
                attachmentLinkIdProperty.Tag = 0x7FFA0003;
                attachmentLinkIdProperty.Type = PropertyType.Integer32;
                attachmentLinkIdProperty.Value = BitConverter.GetBytes(i);
                attachmentLinkIdProperty.IsReadable = true;
                attachmentLinkIdProperty.IsWriteable = true;

                attachmentPropertiesMemoryStream.Write(attachmentLinkIdProperty.ToBytes(), 0, 16);

                //PR_RECORD_KEY
                attachment.RecordKey = BitConverter.GetBytes(i);
                Independentsoft.IO.StructuredStorage.Stream recordKeyStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_0FF90102", attachment.RecordKey);
                attachmentStorage.DirectoryEntries.Add(recordKeyStream);

                Property recordKeyProperty = new Property();
                recordKeyProperty.Tag = 0x0FF90102;
                recordKeyProperty.Type = PropertyType.Binary;
                recordKeyProperty.Size = (uint)attachment.RecordKey.Length;
                recordKeyProperty.IsReadable = true;
                recordKeyProperty.IsWriteable = true;

                attachmentPropertiesMemoryStream.Write(recordKeyProperty.ToBytes(), 0, 16);

                if (attachment.AdditionalInfo != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream additionalInfoStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_370F0102", attachment.AdditionalInfo);
                    attachmentStorage.DirectoryEntries.Add(additionalInfoStream);

                    Property additionalInfoProperty = new Property();
                    additionalInfoProperty.Tag = 0x370F0102;
                    additionalInfoProperty.Type = PropertyType.Binary;
                    additionalInfoProperty.Size = (uint)attachment.AdditionalInfo.Length;
                    additionalInfoProperty.IsReadable = true;
                    additionalInfoProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(additionalInfoProperty.ToBytes(), 0, 16);
                }

                if (attachment.ContentBase != null)
                {
                    byte[] contentBaseBuffer = encoding.GetBytes(attachment.ContentBase);
                    Independentsoft.IO.StructuredStorage.Stream contentBaseStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3711" + stringTypeMask, contentBaseBuffer);
                    attachmentStorage.DirectoryEntries.Add(contentBaseStream);

                    Property contentBaseProperty = new Property();
                    contentBaseProperty.Tag = (uint)0x3711 << 16 | stringTypeHexMask;
                    contentBaseProperty.Type = PropertyType.String8;
                    contentBaseProperty.Size = (uint)(contentBaseBuffer.Length + encoding.GetBytes("\0").Length);
                    contentBaseProperty.IsReadable = true;
                    contentBaseProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(contentBaseProperty.ToBytes(), 0, 16);
                }

                if (attachment.ContentId != null)
                {
                    byte[] contentIdBuffer = encoding.GetBytes(attachment.ContentId);
                    Independentsoft.IO.StructuredStorage.Stream contentIdStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3712" + stringTypeMask, contentIdBuffer);
                    attachmentStorage.DirectoryEntries.Add(contentIdStream);

                    Property contentIdProperty = new Property();
                    contentIdProperty.Tag = (uint)0x3712 << 16 | stringTypeHexMask;
                    contentIdProperty.Type = PropertyType.String8;
                    contentIdProperty.Size = (uint)(contentIdBuffer.Length + encoding.GetBytes("\0").Length);
                    contentIdProperty.IsReadable = true;
                    contentIdProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(contentIdProperty.ToBytes(), 0, 16);
                }

                if (attachment.ContentLocation != null)
                {
                    byte[] contentLocationBuffer = encoding.GetBytes(attachment.ContentLocation);
                    Independentsoft.IO.StructuredStorage.Stream contentLocationStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3713" + stringTypeMask, contentLocationBuffer);
                    attachmentStorage.DirectoryEntries.Add(contentLocationStream);

                    Property contentLocationProperty = new Property();
                    contentLocationProperty.Tag = (uint)0x3713 << 16 | stringTypeHexMask;
                    contentLocationProperty.Type = PropertyType.String8;
                    contentLocationProperty.Size = (uint)(contentLocationBuffer.Length + encoding.GetBytes("\0").Length);
                    contentLocationProperty.IsReadable = true;
                    contentLocationProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(contentLocationProperty.ToBytes(), 0, 16);
                }

                if (attachment.ContentDisposition != null)
                {
                    byte[] contentDispositionBuffer = encoding.GetBytes(attachment.ContentDisposition);
                    Independentsoft.IO.StructuredStorage.Stream contentDispositionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3716" + stringTypeMask, contentDispositionBuffer);
                    attachmentStorage.DirectoryEntries.Add(contentDispositionStream);

                    Property contentDispositionProperty = new Property();
                    contentDispositionProperty.Tag = (uint)0x3716 << 16 | stringTypeHexMask;
                    contentDispositionProperty.Type = PropertyType.String8;
                    contentDispositionProperty.Size = (uint)(contentDispositionBuffer.Length + encoding.GetBytes("\0").Length);
                    contentDispositionProperty.IsReadable = true;
                    contentDispositionProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(contentDispositionProperty.ToBytes(), 0, 16);
                }

                if (attachment.Data != null)
                {
                    if (attachment.Method == AttachmentMethod.Ole)
                    {
                        Independentsoft.IO.StructuredStorage.Storage dataObjectStorage = new Independentsoft.IO.StructuredStorage.Storage("__substg1.0_3701000D");
                        attachmentStorage.DirectoryEntries.Add(dataObjectStorage);

                        MemoryStream oleStorage = new MemoryStream(attachment.Data);
                        CompoundFile oleCompoundFile = new CompoundFile(oleStorage);

                        dataObjectStorage.classId = oleCompoundFile.Root.ClassId;
                        for (int j = 0; j < oleCompoundFile.Root.DirectoryEntries.Count; ++j)
                        {
                            dataObjectStorage.DirectoryEntries.Add(oleCompoundFile.Root.DirectoryEntries[j]);
                        }

                        Property dataProperty = new Property();
                        dataProperty.Tag = 0x3701000D;
                        dataProperty.Type = PropertyType.Object;
                        dataProperty.Size = (uint)UInt32.MaxValue;
                        dataProperty.IsReadable = true;
                        dataProperty.IsWriteable = true;
                        byte[] tempDataProperty = dataProperty.ToBytes();
                        tempDataProperty[12] = 4;

                        attachmentPropertiesMemoryStream.Write(tempDataProperty, 0, 16);
                    }
                    else
                    {

                        Independentsoft.IO.StructuredStorage.Stream dataStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_37010102", attachment.Data);
                        attachmentStorage.DirectoryEntries.Add(dataStream);

                        Property dataProperty = new Property();
                        dataProperty.Tag = 0x37010102;
                        dataProperty.Type = PropertyType.Binary;
                        dataProperty.Size = (uint)attachment.Data.Length;
                        dataProperty.IsReadable = true;
                        dataProperty.IsWriteable = true;

                        attachmentPropertiesMemoryStream.Write(dataProperty.ToBytes(), 0, 16);
                    }
                }

                if (attachment.Encoding != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream encodingStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_37020102", attachment.Encoding);
                    attachmentStorage.DirectoryEntries.Add(encodingStream);

                    Property encodingProperty = new Property();
                    encodingProperty.Tag = 0x37020102;
                    encodingProperty.Type = PropertyType.Binary;
                    encodingProperty.Size = (uint)attachment.Encoding.Length;
                    encodingProperty.IsReadable = true;
                    encodingProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(encodingProperty.ToBytes(), 0, 16);
                }

                if (attachment.Extension != null)
                {
                    byte[] extensionBuffer = encoding.GetBytes(attachment.Extension);
                    Independentsoft.IO.StructuredStorage.Stream extensionStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3703" + stringTypeMask, extensionBuffer);
                    attachmentStorage.DirectoryEntries.Add(extensionStream);

                    Property extensionProperty = new Property();
                    extensionProperty.Tag = (uint)0x3703 << 16 | stringTypeHexMask;
                    extensionProperty.Type = PropertyType.String8;
                    extensionProperty.Size = (uint)(extensionBuffer.Length + encoding.GetBytes("\0").Length);
                    extensionProperty.IsReadable = true;
                    extensionProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(extensionProperty.ToBytes(), 0, 16);
                }

                if (attachment.FileName != null)
                {
                    byte[] fileNameBuffer = encoding.GetBytes(attachment.FileName);
                    Independentsoft.IO.StructuredStorage.Stream fileNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3704" + stringTypeMask, fileNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(fileNameStream);

                    Property fileNameProperty = new Property();
                    fileNameProperty.Tag = (uint)0x3704 << 16 | stringTypeHexMask;
                    fileNameProperty.Type = PropertyType.String8;
                    fileNameProperty.Size = (uint)(fileNameBuffer.Length + encoding.GetBytes("\0").Length);
                    fileNameProperty.IsReadable = true;
                    fileNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(fileNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.Flags != AttachmentFlags.None)
                {
                    Property flagsProperty = new Property();
                    flagsProperty.Tag = 0x37140003;
                    flagsProperty.Type = PropertyType.Integer32;
                    flagsProperty.Value = BitConverter.GetBytes(EnumUtil.ParseAttachmentFlags(attachment.Flags));
                    flagsProperty.IsReadable = true;
                    flagsProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(flagsProperty.ToBytes(), 0, 16);
                }

                if (attachment.LongFileName != null)
                {
                    byte[] longFileNameBuffer = encoding.GetBytes(attachment.LongFileName);
                    Independentsoft.IO.StructuredStorage.Stream longFileNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3707" + stringTypeMask, longFileNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(longFileNameStream);

                    Property longFileNameProperty = new Property();
                    longFileNameProperty.Tag = (uint)0x3707 << 16 | stringTypeHexMask;
                    longFileNameProperty.Type = PropertyType.String8;
                    longFileNameProperty.Size = (uint)(longFileNameBuffer.Length + encoding.GetBytes("\0").Length);
                    longFileNameProperty.IsReadable = true;
                    longFileNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(longFileNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.LongPathName != null)
                {
                    byte[] longPathNameBuffer = encoding.GetBytes(attachment.LongPathName);
                    Independentsoft.IO.StructuredStorage.Stream longPathNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_370D" + stringTypeMask, longPathNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(longPathNameStream);

                    Property longPathNameProperty = new Property();
                    longPathNameProperty.Tag = (uint)0x370D << 16 | stringTypeHexMask;
                    longPathNameProperty.Type = PropertyType.String8;
                    longPathNameProperty.Size = (uint)(longPathNameBuffer.Length + encoding.GetBytes("\0").Length);
                    longPathNameProperty.IsReadable = true;
                    longPathNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(longPathNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.Method != AttachmentMethod.None)
                {
                    Property methodProperty = new Property();
                    methodProperty.Tag = 0x37050003;
                    methodProperty.Type = PropertyType.Integer32;
                    methodProperty.Value = BitConverter.GetBytes(EnumUtil.ParseAttachmentMethod(attachment.Method));
                    methodProperty.IsReadable = true;
                    methodProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(methodProperty.ToBytes(), 0, 16);
                }

                if (attachment.MimeSequence > 0)
                {
                    Property mimeSequenceProperty = new Property();
                    mimeSequenceProperty.Tag = 0x37100003;
                    mimeSequenceProperty.Type = PropertyType.Integer32;
                    mimeSequenceProperty.Value = BitConverter.GetBytes(attachment.MimeSequence);
                    mimeSequenceProperty.IsReadable = true;
                    mimeSequenceProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(mimeSequenceProperty.ToBytes(), 0, 16);
                }

                if (attachment.MimeTag != null)
                {
                    byte[] mimeTagBuffer = encoding.GetBytes(attachment.MimeTag);
                    Independentsoft.IO.StructuredStorage.Stream mimeTagStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_370E" + stringTypeMask, mimeTagBuffer);
                    attachmentStorage.DirectoryEntries.Add(mimeTagStream);

                    Property mimeTagProperty = new Property();
                    mimeTagProperty.Tag = (uint)0x370E << 16 | stringTypeHexMask;
                    mimeTagProperty.Type = PropertyType.String8;
                    mimeTagProperty.Size = (uint)(mimeTagBuffer.Length + encoding.GetBytes("\0").Length);
                    mimeTagProperty.IsReadable = true;
                    mimeTagProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(mimeTagProperty.ToBytes(), 0, 16);
                }

                if (attachment.PathName != null)
                {
                    byte[] pathNameBuffer = encoding.GetBytes(attachment.PathName);
                    Independentsoft.IO.StructuredStorage.Stream pathNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3708" + stringTypeMask, pathNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(pathNameStream);

                    Property pathNameProperty = new Property();
                    pathNameProperty.Tag = (uint)0x3708 << 16 | stringTypeHexMask;
                    pathNameProperty.Type = PropertyType.String8;
                    pathNameProperty.Size = (uint)(pathNameBuffer.Length + encoding.GetBytes("\0").Length);
                    pathNameProperty.IsReadable = true;
                    pathNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(pathNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.Rendering != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream renderingStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_37090102", attachment.Rendering);
                    attachmentStorage.DirectoryEntries.Add(renderingStream);

                    Property renderingProperty = new Property();
                    renderingProperty.Tag = 0x37090102;
                    renderingProperty.Type = PropertyType.Binary;
                    renderingProperty.Size = (uint)attachment.Rendering.Length;
                    renderingProperty.IsReadable = true;
                    renderingProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(renderingProperty.ToBytes(), 0, 16);
                }

                if (attachment.RenderingPosition > 0)
                {
                    Property renderingPositionProperty = new Property();
                    renderingPositionProperty.Tag = 0x370b0003;
                    renderingPositionProperty.Type = PropertyType.Integer32;
                    renderingPositionProperty.Value = BitConverter.GetBytes(attachment.RenderingPosition);
                    renderingPositionProperty.IsReadable = true;
                    renderingPositionProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(renderingPositionProperty.ToBytes(), 0, 16);
                }

                if (attachment.Size > 0)
                {
                    Property sizeProperty = new Property();
                    sizeProperty.Tag = 0x0E200003;
                    sizeProperty.Type = PropertyType.Integer32;
                    sizeProperty.Value = BitConverter.GetBytes(attachment.Size);
                    sizeProperty.IsReadable = true;
                    sizeProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(sizeProperty.ToBytes(), 0, 16);
                }

                if (attachment.Tag != null)
                {
                    Independentsoft.IO.StructuredStorage.Stream tagStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_370A0102", attachment.Tag);
                    attachmentStorage.DirectoryEntries.Add(tagStream);

                    Property tagProperty = new Property();
                    tagProperty.Tag = 0x370A0102;
                    tagProperty.Type = PropertyType.Binary;
                    tagProperty.Size = (uint)attachment.Tag.Length;
                    tagProperty.IsReadable = true;
                    tagProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(tagProperty.ToBytes(), 0, 16);
                }

                if (attachment.TransportName != null)
                {
                    byte[] transportNameBuffer = encoding.GetBytes(attachment.TransportName);
                    Independentsoft.IO.StructuredStorage.Stream transportNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_370C" + stringTypeMask, transportNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(transportNameStream);

                    Property transportNameProperty = new Property();
                    transportNameProperty.Tag = (uint)0x370C << 16 | stringTypeHexMask;
                    transportNameProperty.Type = PropertyType.String8;
                    transportNameProperty.Size = (uint)(transportNameBuffer.Length + encoding.GetBytes("\0").Length);
                    transportNameProperty.IsReadable = true;
                    transportNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(transportNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.DisplayName != null)
                {
                    byte[] displayNameBuffer = encoding.GetBytes(attachment.DisplayName);
                    Independentsoft.IO.StructuredStorage.Stream displayNameStream = new Independentsoft.IO.StructuredStorage.Stream("__substg1.0_3001" + stringTypeMask, displayNameBuffer);
                    attachmentStorage.DirectoryEntries.Add(displayNameStream);

                    Property displayNameProperty = new Property();
                    displayNameProperty.Tag = (uint)0x3001 << 16 | stringTypeHexMask;
                    displayNameProperty.Type = PropertyType.String8;
                    displayNameProperty.Size = (uint)(displayNameBuffer.Length + encoding.GetBytes("\0").Length);
                    displayNameProperty.IsReadable = true;
                    displayNameProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(displayNameProperty.ToBytes(), 0, 16);
                }

                if (attachment.EmbeddedMessage != null && attachment.Method != AttachmentMethod.Ole)
                {
                    DirectoryEntryList embeddedMessagePropertiesDirectoryEntries = attachment.EmbeddedMessage.CreateMessageProperties(ref namedProperties);

                    Independentsoft.IO.StructuredStorage.Storage embeddedMessageStorage = new Independentsoft.IO.StructuredStorage.Storage("__substg1.0_3701000D");

                    for (int s = 0; s < embeddedMessagePropertiesDirectoryEntries.Count; s++)
                    {
                        embeddedMessageStorage.DirectoryEntries.Add(embeddedMessagePropertiesDirectoryEntries[s]);
                    }

                    attachmentStorage.DirectoryEntries.Add(embeddedMessageStorage);

                    Property embeddedMessageProperty = new Property();
                    embeddedMessageProperty.Tag = (uint)0x3701000D;
                    embeddedMessageProperty.Type = PropertyType.Object;
                    embeddedMessageProperty.Size = UInt32.MaxValue;
                    embeddedMessageProperty.IsReadable = true;
                    embeddedMessageProperty.IsWriteable = true;

                    byte[] tempEmbeddedMessageProperty = embeddedMessageProperty.ToBytes();
                    tempEmbeddedMessageProperty[12] = 1;

                    attachmentPropertiesMemoryStream.Write(tempEmbeddedMessageProperty, 0, 16);
                }

                if (attachment.ObjectType != ObjectType.None)
                {
                    Property objectTypeProperty = new Property();
                    objectTypeProperty.Tag = 0x0FFE0003;
                    objectTypeProperty.Type = PropertyType.Integer32;
                    objectTypeProperty.Value = BitConverter.GetBytes(EnumUtil.ParseObjectType(attachment.ObjectType));
                    objectTypeProperty.IsReadable = true;
                    objectTypeProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(objectTypeProperty.ToBytes(), 0, 16);
                }

                if (attachment.IsHidden)
                {
                    Property isHiddenProperty = new Property();
                    isHiddenProperty.Tag = 0x7FFE000B;
                    isHiddenProperty.Type = PropertyType.Boolean;
                    isHiddenProperty.Value = BitConverter.GetBytes(1);
                    isHiddenProperty.IsReadable = true;
                    isHiddenProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(isHiddenProperty.ToBytes(), 0, 16);
                }

                if (attachment.IsContactPhoto)
                {
                    Property isContactPhotoProperty = new Property();
                    isContactPhotoProperty.Tag = 0x7FFF000B;
                    isContactPhotoProperty.Type = PropertyType.Boolean;
                    isContactPhotoProperty.Value = BitConverter.GetBytes(1);
                    isContactPhotoProperty.IsReadable = true;
                    isContactPhotoProperty.IsWriteable = true;

                    attachmentPropertiesMemoryStream.Write(isContactPhotoProperty.ToBytes(), 0, 16);
                }

                if (attachment.CreationTime.CompareTo(DateTime.MinValue) > 0)
                {
                    DateTime year1601 = new DateTime(1601, 1, 1);
                    TimeSpan timeSpan = attachment.CreationTime.ToUniversalTime().Subtract(year1601);

                    long ticks = timeSpan.Ticks;

                    byte[] ticksBytes = BitConverter.GetBytes(ticks);

                    Property creationTimeProperty = new Property();
                    creationTimeProperty.Tag = 0x30070040;
                    creationTimeProperty.Type = PropertyType.Time;
                    creationTimeProperty.Value = ticksBytes;
                    creationTimeProperty.IsReadable = true;
                    creationTimeProperty.IsWriteable = false;

                    attachmentPropertiesMemoryStream.Write(creationTimeProperty.ToBytes(), 0, 16);
                }

                if (attachment.LastModificationTime.CompareTo(DateTime.MinValue) > 0)
                {
                    DateTime year1601 = new DateTime(1601, 1, 1);
                    TimeSpan timeSpan = attachment.LastModificationTime.ToUniversalTime().Subtract(year1601);

                    long ticks = timeSpan.Ticks;

                    byte[] ticksBytes = BitConverter.GetBytes(ticks);

                    Property lastModificationTimeProperty = new Property();
                    lastModificationTimeProperty.Tag = 0x30080040;
                    lastModificationTimeProperty.Type = PropertyType.Time;
                    lastModificationTimeProperty.Value = ticksBytes;
                    lastModificationTimeProperty.IsReadable = true;
                    lastModificationTimeProperty.IsWriteable = false;

                    attachmentPropertiesMemoryStream.Write(lastModificationTimeProperty.ToBytes(), 0, 16);
                }

                if (attachment.DataObjectStorage != null && attachment.Method == AttachmentMethod.Ole)
                {
                    attachmentStorage.DirectoryEntries.Add(attachment.DataObjectStorage);
                    attachmentStorage.DirectoryEntries.Add(attachment.PropertiesStream);
                }
                else
                {
                    Independentsoft.IO.StructuredStorage.Stream attachmentPropertiesStream = new Independentsoft.IO.StructuredStorage.Stream("__properties_version1.0", attachmentPropertiesMemoryStream.ToArray());
                    attachmentStorage.DirectoryEntries.Add(attachmentPropertiesStream);
                }

                directoryEntries.Add(attachmentStorage);
            }

            return directoryEntries;
        }

        private static int GetGuidIndex(byte[] guid1, IList<byte[]> guidList)
        {
            if (guid1 != null)
            {
                for (int l = 0; l < guidList.Count; l++)
                {
                    byte[] guid2 = guidList[l];
                    bool equal = true;

                    for (int i = 0; i < 16; i++)
                    {
                        if (guid1[i] != guid2[i])
                        {
                            equal = false;
                            break;
                        }
                    }

                    if (equal)
                    {
                        return l;
                    }
                }
            }

            return -1;
        }

        /// <summary>
        /// Converts to MIME message.
        /// </summary>
        /// <returns>Independentsoft.Email.Mime.Message.</returns>
        public Independentsoft.Email.Mime.Message ConvertToMimeMessage()
        {
            return ConvertToMimeMessageImplementation();
        }

        private Independentsoft.Email.Mime.Message ConvertToMimeMessageImplementation()
        {
            Independentsoft.Email.Mime.Message mimeMessage = new Independentsoft.Email.Mime.Message();

            if (this.TransportMessageHeaders != null)
            {
                string transportMessageHeaders = this.TransportMessageHeaders;

                if (!transportMessageHeaders.EndsWith("\r\n\r\n"))
                {
                    transportMessageHeaders += "\r\n\r\n";
                }

                System.Text.Encoding codePageEncoding = Util.GetEncoding(this.InternetCodePage);

                byte[] transportMessageHeadersBuffer = codePageEncoding.GetBytes(transportMessageHeaders);
                mimeMessage = new Independentsoft.Email.Mime.Message(transportMessageHeadersBuffer);

                mimeMessage.BodyParts.Clear();
                mimeMessage.To.Clear();
                mimeMessage.Cc.Clear();
                mimeMessage.Bcc.Clear();
            }

            for (int i = 0; i < this.Recipients.Count; i++)
            {
                Independentsoft.Email.Mime.Mailbox mimeMailbox = new Independentsoft.Email.Mime.Mailbox();

                if (this.Recipients[i].DisplayName != null && this.Recipients[i].DisplayName.Length > 0)
                {
                    mimeMailbox.Name = this.Recipients[i].DisplayName;
                }
                else if (this.Recipients[i].SmtpAddress != null)
                {
                    mimeMailbox.Name = this.Recipients[i].SmtpAddress;
                }
                else
                {
                    mimeMailbox.Name = this.Recipients[i].EmailAddress;
                }

                if (this.Recipients[i].SmtpAddress != null)
                {
                    mimeMailbox.EmailAddress = this.Recipients[i].SmtpAddress;
                }
                else
                {
                    mimeMailbox.EmailAddress = this.Recipients[i].EmailAddress;
                }


                if (this.Recipients[i].RecipientType == RecipientType.To)
                {
                    mimeMessage.To.Add(mimeMailbox);
                }
                else if (this.Recipients[i].RecipientType == RecipientType.Cc)
                {
                    mimeMessage.Cc.Add(mimeMailbox);
                }
                else if (this.Recipients[i].RecipientType == RecipientType.Bcc)
                {
                    mimeMessage.Bcc.Add(mimeMailbox);
                }
            }

            if (mimeMessage.From == null) //not set in TransportMessageHeaders
            {
                Independentsoft.Email.Mime.Mailbox fromMailbox = new Independentsoft.Email.Mime.Mailbox();

                if (this.SenderName != null && this.SenderName.Length > 0)
                {
                    fromMailbox.Name = this.SenderName;
                }
                else
                {
                    fromMailbox.Name = this.SenderEmailAddress;
                }

                if (this.InternetAccountName != null && this.InternetAccountName.Contains("@"))
                {
                    fromMailbox.EmailAddress = this.InternetAccountName;
                }
                else
                {
                    fromMailbox.EmailAddress = this.SenderEmailAddress;
                }

                mimeMessage.From = fromMailbox;
            }

            mimeMessage.Subject = this.Subject;
            mimeMessage.Date = this.ClientSubmitTime;
            mimeMessage.ContentType = new Independentsoft.Email.Mime.ContentType("multipart", "mixed");

            Independentsoft.Email.Mime.BodyPart alternativeBodyPart = new Independentsoft.Email.Mime.BodyPart();
            alternativeBodyPart.ContentType = new Independentsoft.Email.Mime.ContentType("multipart", "alternative");

            Independentsoft.Email.Mime.BodyPart relatedBodyPart = new Independentsoft.Email.Mime.BodyPart();
            relatedBodyPart.ContentType = new Independentsoft.Email.Mime.ContentType("multipart", "related");

            if (this.Body != null && this.Body.Length > 0)
            {
                Independentsoft.Email.Mime.BodyPart plainBodyPart = new Independentsoft.Email.Mime.BodyPart();
                plainBodyPart.ContentType = new Independentsoft.Email.Mime.ContentType("text", "plain", "utf-8");
                plainBodyPart.ContentTransferEncoding = Independentsoft.Email.Mime.ContentTransferEncoding.QuotedPrintable;
                plainBodyPart.Body = this.Body;

                alternativeBodyPart.BodyParts.Add(plainBodyPart);
            }

            string bodyHtmlText = this.BodyHtmlText;

            if (bodyHtmlText != null && bodyHtmlText.Length > 0)
            {
                System.Text.Encoding htmlEncoding = System.Text.Encoding.UTF7;
                string htmlCharset = "utf-8";

                int htmlCharsetIndex = bodyHtmlText.IndexOf("charset=");

                if (htmlCharsetIndex > 0)
                {
                    int htmlCharsetEndIndex = bodyHtmlText.IndexOf("\"", htmlCharsetIndex);

                    if (htmlCharsetEndIndex != -1)
                    {
                        htmlCharset = bodyHtmlText.Substring(htmlCharsetIndex + 8, htmlCharsetEndIndex - htmlCharsetIndex - 8);

                        try
                        {
							// PRGX: Scrub chracter set for proper names
                            switch (htmlCharset.ToLower())
                            {
                              case "utf8": 
                                htmlCharset = "UTF-8";
                                break;
                            }
                            if (htmlCharset.Length > 0)
                            {
                              htmlEncoding = System.Text.Encoding.GetEncoding(htmlCharset);
                            }
                        }
                        catch
                        {
                        }
                    }
                }

                Independentsoft.Email.Mime.BodyPart htmlBodyPart = new Independentsoft.Email.Mime.BodyPart();

                byte[] bodyHtml = this.BodyHtml;

                if (this.Encoding == System.Text.Encoding.Unicode && bodyHtml != null)
                {
                    if (Utf8Util.IsUtf8(bodyHtml, bodyHtml.Length))
                    {
                        htmlEncoding = System.Text.Encoding.UTF8;
                        htmlCharset = "utf-8";
                    }

                    htmlBodyPart.Body = htmlEncoding.GetString(bodyHtml, 0, bodyHtml.Length);
                }
                else
                {
                    htmlBodyPart.Body = bodyHtmlText;
                }

                htmlBodyPart.ContentType = new Independentsoft.Email.Mime.ContentType("text", "html", htmlCharset);
                htmlBodyPart.ContentTransferEncoding = Independentsoft.Email.Mime.ContentTransferEncoding.QuotedPrintable;

                alternativeBodyPart.BodyParts.Add(htmlBodyPart);
            }

            if (alternativeBodyPart.BodyParts.Count > 0)
            {
                relatedBodyPart.BodyParts.Add(alternativeBodyPart);
            }

            IList<Independentsoft.Email.Mime.Attachment> attachmentList = new List<Independentsoft.Email.Mime.Attachment>();

            for (int i = 0; i < this.Attachments.Count; i++)
            {
                if (this.Attachments[i].EmbeddedMessage != null)
                {
                    Independentsoft.Email.Mime.ContentType mainContentType = mimeMessage.ContentType;

                    mainContentType.Type = "multipart";
                    mainContentType.SubType = "mixed";

                    mimeMessage.ContentType = mainContentType;

                    Independentsoft.Email.Mime.Message emlAttachment = this.Attachments[i].EmbeddedMessage.ConvertToMimeMessage();

                    Independentsoft.Email.Mime.BodyPart mimeAttachment = new Independentsoft.Email.Mime.BodyPart();
                    mimeAttachment.Body = emlAttachment.ToString();

                    Independentsoft.Email.Mime.ContentType mimeAttachmentContentType = new Email.Mime.ContentType("message", "rfc822");
                    mimeAttachmentContentType.Parameters.Add(new Independentsoft.Email.Mime.Parameter("name", "\"" + emlAttachment.GetFileName() + "\""));

                    mimeAttachment.ContentType = mimeAttachmentContentType;

                    Independentsoft.Email.Mime.ContentDisposition mimeAttachmentContentDisposition = new Email.Mime.ContentDisposition(Independentsoft.Email.Mime.ContentDispositionType.Attachment);
                    mimeAttachmentContentDisposition.Parameters.Add(new Independentsoft.Email.Mime.Parameter("filename", "\"" + emlAttachment.GetFileName() + "\""));

                    mimeAttachment.ContentDisposition = mimeAttachmentContentDisposition;

                    mimeMessage.BodyParts.Add(mimeAttachment);
                }
                else
                {
                    Independentsoft.Email.Mime.Attachment mimeAttachment = new Independentsoft.Email.Mime.Attachment(this.Attachments[i].GetBytes());

                    if (this.Attachments[i].MimeTag != null)
                    {
                        mimeAttachment.ContentType = new Independentsoft.Email.Mime.ContentType(this.Attachments[i].MimeTag);
                    }

                    if (this.Attachments[i].ContentId != null)
                    {
                        if (!this.Attachments[i].ContentId.StartsWith("<"))
                        {
                            this.Attachments[i].ContentId = "<" + this.Attachments[i].ContentId;
                        }

                        if (!this.Attachments[i].ContentId.EndsWith(">"))
                        {
                            this.Attachments[i].ContentId += ">";
                        }

                        mimeAttachment.ContentID = this.Attachments[i].ContentId;
                    }

                    mimeAttachment.ContentLocation = this.Attachments[i].ContentLocation;

                    if (this.Attachments[i].LongFileName != null)
                    {
                        mimeAttachment.Name = this.Attachments[i].LongFileName;
                    }
                    else if (this.Attachments[i].DisplayName != null)
                    {
                        mimeAttachment.Name = this.Attachments[i].DisplayName;
                    }
                    else if (this.Attachments[i].FileName != null)
                    {
                        mimeAttachment.Name = this.Attachments[i].FileName;
                    }

                    if (mimeAttachment.ContentID != null || mimeAttachment.ContentLocation != null || (mimeAttachment.ContentDisposition != null && mimeAttachment.ContentDisposition.Type == Email.Mime.ContentDispositionType.Inline)) //embedded
                    {
                        relatedBodyPart.BodyParts.Add(mimeAttachment);
                    }
                    else //normal attachements
                    {
                        attachmentList.Add(mimeAttachment);
                    }
                }
            }

            if (relatedBodyPart.BodyParts.Count > 0)
            {
                mimeMessage.BodyParts.Add(relatedBodyPart);
            }

            for (int i = 0; i < attachmentList.Count; i++)
            {
                mimeMessage.BodyParts.Add(attachmentList[i]);
            }

            return mimeMessage;
        }

        private void AddMailboxCollections(Independentsoft.Email.Mime.Message mimeMessage)
        {
            string toString = "";
            string ccString = "";
            string bccString = "";
            string replyToString = "";

            for (int i = 0; i < mimeMessage.To.Count; i++)
            {
                if (mimeMessage.To[i] != null)
                {
                    toString += mimeMessage.To[i].ToString();
                }

                if (i < mimeMessage.To.Count - 1)
                {
                    toString += ", ";
                }
            }

            for (int i = 0; i < mimeMessage.Cc.Count; i++)
            {
                if (mimeMessage.Cc[i] != null)
                {
                    ccString += mimeMessage.Cc[i].ToString();
                }

                if (i < mimeMessage.Cc.Count - 1)
                {
                    ccString += ", ";
                }
            }

            for (int i = 0; i < mimeMessage.Bcc.Count; i++)
            {
                if (mimeMessage.Bcc[i] != null)
                {
                    bccString += mimeMessage.Bcc[i].ToString();
                }

                if (i < mimeMessage.Bcc.Count - 1)
                {
                    bccString += ", ";
                }
            }

            for (int i = 0; i < mimeMessage.ReplyTo.Count; i++)
            {
                if (mimeMessage.ReplyTo[i] != null)
                {
                    replyToString += mimeMessage.ReplyTo[i].ToString();
                }

                if (i < mimeMessage.ReplyTo.Count - 1)
                {
                    replyToString += ", ";
                }
            }

            if (toString.Length > 0)
            {
                Independentsoft.Email.Mime.Header toHeader = new Independentsoft.Email.Mime.Header(StandardHeader.To, toString);

                mimeMessage.Headers.Remove(StandardHeader.To);
                mimeMessage.Headers.Add(toHeader);
            }

            if (ccString.Length > 0)
            {
                Independentsoft.Email.Mime.Header ccHeader = new Independentsoft.Email.Mime.Header(StandardHeader.Cc, ccString);

                mimeMessage.Headers.Remove(StandardHeader.Cc);
                mimeMessage.Headers.Add(ccHeader);
            }

            if (bccString.Length > 0)
            {
                Independentsoft.Email.Mime.Header bccHeader = new Independentsoft.Email.Mime.Header(StandardHeader.Bcc, bccString);

                mimeMessage.Headers.Remove(StandardHeader.Bcc);
                mimeMessage.Headers.Add(bccHeader);
            }

            if (replyToString.Length > 0)
            {
                Independentsoft.Email.Mime.Header replyToHeader = new Independentsoft.Email.Mime.Header(StandardHeader.ReplyTo, replyToString);

                mimeMessage.Headers.Remove(StandardHeader.ReplyTo);
                mimeMessage.Headers.Add(replyToHeader);
            }
        }

        private void ConvertFromMimeMessage(Independentsoft.Email.Mime.Message mimeMessage)
        {
            if (mimeMessage != null)
            {
                this.transportMessageHeaders = "";

                //add missing headers removed during mimeMessage parse
                AddMailboxCollections(mimeMessage);

                for (int i = 0; i < mimeMessage.Headers.Count; i++)
                {
                    this.transportMessageHeaders += mimeMessage.Headers[i].ToString() + "\r\n";
                }

                for (int i = 0; i < mimeMessage.To.Count; i++)
                {
                    Independentsoft.Email.Mime.Mailbox mailbox = mimeMessage.To[i];

                    Recipient recipient = new Recipient();
                    recipient.AddressType = "SMTP";
                    recipient.RecipientType = RecipientType.To;
                    recipient.EmailAddress = mailbox.EmailAddress;
                    recipient.EntryId = CreateRecipientEntryID(mailbox.EmailAddress);

                    if (mailbox.Name != null && mailbox.Name.Length > 0)
                    {
                        recipient.DisplayName = mailbox.Name;
                        this.DisplayTo += mailbox.Name;
                    }
                    else
                    {
                        recipient.DisplayName = mailbox.EmailAddress;
                        this.DisplayTo += mailbox.EmailAddress;
                    }

                    if (i < mimeMessage.To.Count - 1)
                    {
                        this.DisplayTo += "; ";
                    }

                    this.Recipients.Add(recipient);
                }

                for (int i = 0; i < mimeMessage.Cc.Count; i++)
                {
                    Independentsoft.Email.Mime.Mailbox mailbox = mimeMessage.Cc[i];

                    Recipient recipient = new Recipient();
                    recipient.AddressType = "SMTP";
                    recipient.RecipientType = RecipientType.Cc;
                    recipient.EmailAddress = mailbox.EmailAddress;
                    recipient.EntryId = CreateRecipientEntryID(mailbox.EmailAddress);

                    if (mailbox.Name != null && mailbox.Name.Length > 0)
                    {
                        recipient.DisplayName = mailbox.Name;
                        this.DisplayCc += mailbox.Name;
                    }
                    else
                    {
                        recipient.DisplayName = mailbox.EmailAddress;
                        this.DisplayCc += mailbox.EmailAddress;
                    }

                    if (i < mimeMessage.Cc.Count - 1)
                    {
                        this.DisplayCc += "; ";
                    }

                    this.Recipients.Add(recipient);
                }

                for (int i = 0; i < mimeMessage.Bcc.Count; i++)
                {
                    Independentsoft.Email.Mime.Mailbox mailbox = mimeMessage.Bcc[i];

                    Recipient recipient = new Recipient();
                    recipient.AddressType = "SMTP";
                    recipient.RecipientType = RecipientType.Bcc;
                    recipient.EmailAddress = mailbox.EmailAddress;
                    recipient.EntryId = CreateRecipientEntryID(mailbox.EmailAddress);

                    if (mailbox.Name != null && mailbox.Name.Length > 0)
                    {
                        recipient.DisplayName = mailbox.Name;
                        this.DisplayBcc += mailbox.Name;
                    }
                    else
                    {
                        recipient.DisplayName = mailbox.EmailAddress;
                        this.DisplayBcc += mailbox.EmailAddress;
                    }

                    if (i < mimeMessage.Bcc.Count - 1)
                    {
                        this.DisplayBcc += "; ";
                    }

                    this.Recipients.Add(recipient);
                }

                this.CreationTime = mimeMessage.Date;
                this.MessageDeliveryTime = mimeMessage.Date;
                this.ClientSubmitTime = mimeMessage.Date;
                this.InternetMessageId = mimeMessage.MessageID;

                if (mimeMessage.Headers["Priority"] != null && mimeMessage.Headers["Priority"].Value != null && mimeMessage.Headers["Priority"].Value.ToLower() == "urgent")
                {
                    this.Priority = Priority.High;
                    this.Importance = Importance.High;
                }
                else if (mimeMessage.Headers["Priority"] != null && mimeMessage.Headers["Priority"].Value != null && mimeMessage.Headers["Priority"].Value.ToLower() == "non-urgent")
                {
                    this.Priority = Priority.Low;
                    this.Importance = Importance.Low;
                }
                else if (mimeMessage.Headers["X-Priority"] != null && mimeMessage.Headers["X-Priority"].Value != null && mimeMessage.Headers["X-Priority"].Value == "1")
                {
                    this.Priority = Priority.High;
                    this.Importance = Importance.High;
                }
                else if (mimeMessage.Headers["X-Priority"] != null && mimeMessage.Headers["X-Priority"].Value != null && mimeMessage.Headers["X-Priority"].Value == "-1")
                {
                    this.Priority = Priority.Low;
                    this.Importance = Importance.Low;
                }
                else if (mimeMessage.Headers["Importance"] != null && mimeMessage.Headers["Importance"].Value != null && mimeMessage.Headers["Importance"].Value.ToLower() == "high")
                {
                    this.Priority = Priority.High;
                    this.Importance = Importance.High;
                }
                else if (mimeMessage.Headers["Importance"] != null && mimeMessage.Headers["Importance"].Value != null && mimeMessage.Headers["Importance"].Value.ToLower() == "low")
                {
                    this.Priority = Priority.Low;
                    this.Importance = Importance.Low;
                }

                if (mimeMessage.Sender != null)
                {
                    if (mimeMessage.Sender.Name != null && mimeMessage.Sender.Name.Length > 0)
                    {
                        this.SenderName = mimeMessage.Sender.Name;
                    }
                    else
                    {
                        this.SenderName = mimeMessage.Sender.EmailAddress;
                    }

                    this.SenderEmailAddress = mimeMessage.Sender.EmailAddress;
                    this.SenderAddressType = "SMTP";

                    if (mimeMessage.Sender.EmailAddress != null)
                    {
                        this.SenderEntryId = CreateRecipientEntryID(mimeMessage.Sender.EmailAddress);
                    }
                }
                else if (mimeMessage.From != null)
                {
                    if (mimeMessage.From.Name != null && mimeMessage.From.Name.Length > 0)
                    {
                        this.SenderName = mimeMessage.From.Name;
                    }
                    else
                    {
                        this.SenderName = mimeMessage.From.EmailAddress;
                    }

                    this.SenderEmailAddress = mimeMessage.From.EmailAddress;
                    this.SenderAddressType = "SMTP";

                    if (mimeMessage.From.EmailAddress != null)
                    {
                        this.SenderEntryId = CreateRecipientEntryID(mimeMessage.From.EmailAddress);
                    }
                }

                if (mimeMessage.From != null)
                {
                    this.sentRepresentingEmailAddress = mimeMessage.From.EmailAddress;
                    this.sentRepresentingName = mimeMessage.From.Name;
                    this.sentRepresentingAddressType = "SMTP";
                }

                this.Subject = mimeMessage.Subject;

                //Body
                BodyPartList allBodyParts = new BodyPartList();

                if (mimeMessage.ContentType == null)
                {
                    this.Body = mimeMessage.Body;
                }
                else if (mimeMessage.ContentType != null && mimeMessage.ContentType.Type != null && mimeMessage.ContentType.Type.ToLower() == "text" && mimeMessage.ContentType.SubType != null && mimeMessage.ContentType.SubType.ToLower() == "plain")
                {
                    this.Body = mimeMessage.Body;
                }
                else if (mimeMessage.ContentType != null && mimeMessage.ContentType.Type != null && mimeMessage.ContentType.Type.ToLower() == "text" && mimeMessage.ContentType.SubType != null && mimeMessage.ContentType.SubType.ToLower() == "html")
                {
                    if (mimeMessage.ContentType.CharSet != null)
                    {
                        System.Text.Encoding encoding = Independentsoft.Email.Mime.Util.GetEncoding(mimeMessage.ContentType.CharSet);
                        this.InternetCodePage = encoding.CodePage;

                        if (mimeMessage.Body != null)
                        {
                            this.BodyHtml = Independentsoft.Email.Mime.Util.GetEncoding(mimeMessage.ContentType.CharSet).GetBytes(mimeMessage.Body);
                        }
                    }
                    else
                    {
                        this.InternetCodePage = 65001; //utf8

                        this.BodyHtmlText = mimeMessage.Body;
                    }
                }
                else
                {
                    BodyPartList children = GetChildren(mimeMessage.BodyParts);

                    foreach (Independentsoft.Email.Mime.BodyPart child in children)
                    {
                        allBodyParts.Add(child);
                    }

                    foreach (Independentsoft.Email.Mime.BodyPart bodyPart in allBodyParts)
                    {
                        if (bodyPart.ContentType != null && bodyPart.ContentType.Type != null && bodyPart.ContentType.Type.ToLower() == "text" && bodyPart.ContentType.SubType != null && bodyPart.ContentType.SubType.ToLower() == "plain" && (bodyPart.ContentDisposition == null || bodyPart.ContentDisposition.Type != Independentsoft.Email.Mime.ContentDispositionType.Attachment))
                        {
                            if (this.Body == null)
                            {
                                if (bodyPart.ContentType.CharSet != null)
                                {
                                    System.Text.Encoding encoding = Independentsoft.Email.Mime.Util.GetEncoding(bodyPart.ContentType.CharSet);
                                    this.InternetCodePage = encoding.CodePage;

                                    this.Body = bodyPart.Body;
                                }
                                else
                                {
                                    this.InternetCodePage = 65001; //utf8

                                    this.Body = bodyPart.Body;
                                }
                            }
                        }
                        else if (bodyPart.ContentType != null && bodyPart.ContentType.Type != null && bodyPart.ContentType.Type.ToLower() == "text" && bodyPart.ContentType.SubType != null && bodyPart.ContentType.SubType.ToLower() == "html" && (bodyPart.ContentDisposition == null || bodyPart.ContentDisposition.Type != Independentsoft.Email.Mime.ContentDispositionType.Attachment))
                        {
                            if (bodyPart.ContentType.CharSet != null)
                            {
                                System.Text.Encoding encoding = Independentsoft.Email.Mime.Util.GetEncoding(bodyPart.ContentType.CharSet);
                                this.InternetCodePage = encoding.CodePage;

                                if (bodyPart.Body != null)
                                {
                                    this.BodyHtml = Independentsoft.Email.Mime.Util.GetEncoding(bodyPart.ContentType.CharSet).GetBytes(bodyPart.Body);
                                    //this.BodyRtf = System.Text.Encoding.GetEncoding(bodyPart.ContentType.CharSet).GetBytes(Html2Rtf(bodyPart.Body, 65001));
                                }
                            }
                            else
                            {
                                this.InternetCodePage = 65001; //utf8

                                this.BodyHtmlText = bodyPart.Body;
                            }
                        }
                    }
                }

                //attachments
                Independentsoft.Email.Mime.Attachment[] mimeAttachments = mimeMessage.GetAttachments(true);

                for (int i = 0; i < mimeAttachments.Length; i++)
                {
                    Independentsoft.Msg.Attachment msgAttachment = new Attachment(mimeAttachments[i].Name, mimeAttachments[i].GetBytes());
                    msgAttachment.ContentId = mimeAttachments[i].ContentID;
                    msgAttachment.ContentLocation = mimeAttachments[i].ContentLocation;

                    if (msgAttachment.FileName == null && msgAttachment.DisplayName == null && mimeAttachments[i].ContentDescription != null)
                    {
                        msgAttachment.FileName = mimeAttachments[i].ContentDescription;
                        msgAttachment.DisplayName = mimeAttachments[i].ContentDescription;
                    }

                    this.Attachments.Add(msgAttachment);
                }

                this.Encoding = System.Text.Encoding.Unicode;
            }
        }

        private static BodyPartList GetChildren(BodyPartList bodyParts)
        {
            BodyPartList children = new BodyPartList();

            for (int i = 0; i < bodyParts.Count; i++)
            {
                BodyPartList parts = GetChildren(bodyParts[i].BodyParts);

                foreach(Independentsoft.Email.Mime.BodyPart part in parts)
                {
                    children.Add(part);
                }
                
                children.Add(bodyParts[i]);
            }

            return children;
        }

        private static byte[] CreateRecipientEntryID(string emailAddress)
        {
            byte[] prefix1 = new byte[] { 0x00, 0x00, 0x00, 0x00, 0x81, 0x2B, 0x1F, 0xA4, 0xBE, 0xA3, 0x10, 0x19, 0x9D, 0x6E, 0x00, 0xDD, 0x01, 0x0F, 0x54, 0x02, 0x00, 0x00, 0x01, 0x80 };
            byte[] prefix2 = new byte[] { };

            string sufixString = emailAddress + "\0SMTP\0" + emailAddress + "\0";
            byte[] suffix = System.Text.Encoding.Unicode.GetBytes(sufixString);

            byte[] entryID = new byte[prefix1.Length + suffix.Length];

            System.Array.Copy(prefix1, 0, entryID, 0, prefix1.Length);
            System.Array.Copy(suffix, 0, entryID, prefix1.Length, suffix.Length);

            return entryID;
        }

        #region Properties

        /// <summary>
        /// Gets or sets message encoding. Default is UTF8 encoding.
        /// </summary>
        /// <value>The encoding.</value>
        /// <example>
        /// In order to save message as Unicode use:
        ///   <code>
        /// message.Encoding = System.Text.Encoding.Unicode;
        ///   </code>
        ///   </example>
        public System.Text.Encoding Encoding
        {
            get
            {
                return encoding;
            }
            set
            {
                encoding = value;
            }
        }

        /// <summary>
        /// Contains a text string that identifies the sender-defined message class, such as IPM.Note.
        /// </summary>
        /// <value>The message class.</value>
        public string MessageClass
        {
            get
            {
                return messageClass;
            }
            set
            {
                messageClass = value;
            }
        }

        /// <summary>
        /// Contains the full subject of a message.
        /// </summary>
        /// <value>The subject.</value>
        public string Subject
        {
            get
            {
                return subject;
            }
            set
            {
                subject = value;
            }
        }

        /// <summary>
        /// Contains a subject prefix that typically indicates some action on a message, such as "FW: " for forwarding.
        /// </summary>
        /// <value>The subject prefix.</value>
        public string SubjectPrefix
        {
            get
            {
                return subjectPrefix;
            }
            set
            {
                subjectPrefix = value;
            }
        }

        /// <summary>
        /// Contains the topic of the first message in a conversation thread.
        /// </summary>
        /// <value>The conversation topic.</value>
        /// <remarks>A conversation thread represents a series of messages and replies. These properties are set for the first message in a thread, usually to the <see cref="NormalizedSubject" /> property. Subsequent messages in the thread should use the same topic without modification.</remarks>
        public string ConversationTopic
        {
            get
            {
                return conversationTopic;
            }
            set
            {
                conversationTopic = value;
            }
        }

        /// <summary>
        /// Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;).
        /// </summary>
        /// <value>The display BCC.</value>
        public string DisplayBcc
        {
            get
            {
                return displayBcc;
            }
            set
            {
                displayBcc = value;
            }
        }

        /// <summary>
        /// Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;).
        /// </summary>
        /// <value>The display cc.</value>
        public string DisplayCc
        {
            get
            {
                return displayCc;
            }
            set
            {
                displayCc = value;
            }
        }

        /// <summary>
        /// Contains a list of the display names of the primary (To) message recipients, separated by semicolons (;).
        /// </summary>
        /// <value>The display to.</value>
        public string DisplayTo
        {
            get
            {
                return displayTo;
            }
            set
            {
                displayTo = value;
            }
        }

        /// <summary>
        /// Contains a list of the display names of the primary (To) message recipients, separated by semicolons (;).
        /// </summary>
        /// <value>The original display to.</value>
        public string OriginalDisplayTo
        {
            get
            {
                return originalDisplayTo;
            }
            set
            {
                originalDisplayTo = value;
            }
        }

        /// <summary>
        /// Contains reply to email address.
        /// </summary>
        /// <value>The reply to.</value>
        public string ReplyTo
        {
            get
            {
                return replyTo;
            }
            set
            {
                replyTo = value;
            }
        }

        /// <summary>
        /// Contains the message subject with any prefix removed.
        /// </summary>
        /// <value>The normalized subject.</value>
        public string NormalizedSubject
        {
            get
            {
                return normalizedSubject;
            }
            set
            {
                normalizedSubject = value;
            }
        }

        /// <summary>
        /// Contains the message text.
        /// </summary>
        /// <value>The body.</value>
        public string Body
        {
            get
            {
                return body;
            }
            set
            {
                body = value;
            }
        }

        /// <summary>
        /// Contains the Rich Text Format (RTF) version of the message text.
        /// </summary>
        /// <value>The body RTF.</value>
        public byte[] BodyRtf
        {
            get
            {
                if (rtfCompressed != null && rtfCompressed.Length > 0)
                {
                    return Util.DecompressRtfBody(rtfCompressed);
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
                    using (MemoryStream memoryStream = new MemoryStream(value.Length))
                    {
                        BinaryWriter writer = new BinaryWriter(memoryStream);

                        writer.Write((uint)value.Length + 12);
                        writer.Write((uint)value.Length);
                        writer.Write(0x414c454d);  //magic number

                        Crc crc = new Crc();
                        crc.Update(value, 0, value.Length);

                        writer.Write(crc.Value);
                        writer.Write(value);

                        this.rtfCompressed = new byte[memoryStream.Length];

                        System.Array.Copy(memoryStream.ToArray(), 0, rtfCompressed, 0, rtfCompressed.Length);

                        this.RtfInSync = true;
                    }
                }
            }
        }

        /// <summary>
        /// Contains the Rich Text Format (RTF) version of the message text, usually in compressed form.
        /// </summary>
        /// <value>The RTF compressed.</value>
        public byte[] RtfCompressed
        {
            get
            {
                return rtfCompressed;
            }
            set
            {
                rtfCompressed = value;
            }
        }

        /// <summary>
        /// Contains a binary-comparable key that identifies correlated objects for a search.
        /// </summary>
        /// <value>The search key.</value>
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
        /// Contains a change key of a message.
        /// </summary>
        /// <value>The change key.</value>
        public byte[] ChangeKey
        {
            get
            {
                return changeKey;
            }
            set
            {
                changeKey = value;
            }
        }

        /// <summary>
        /// Contains a MAPI entry identifier used to open and edit properties of a particular MAPI object.
        /// </summary>
        /// <value>The entry identifier.</value>
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
        /// Gets or sets the read receipt entry identifier.
        /// </summary>
        /// <value>The read receipt entry identifier.</value>
        public byte[] ReadReceiptEntryId
        {
            get
            {
                return readReceiptEntryId;
            }
            set
            {
                readReceiptEntryId = value;
            }
        }

        /// <summary>
        /// Gets or sets the read receipt search key.
        /// </summary>
        /// <value>The read receipt search key.</value>
        public byte[] ReadReceiptSearchKey
        {
            get
            {
                return readReceiptSearchKey;
            }
            set
            {
                readReceiptSearchKey = value;
            }
        }

        /// <summary>
        /// Contains the creation date and time of the message.
        /// </summary>
        /// <value>The creation time.</value>
        /// <remarks>A message store sets this property for each message that it creates.</remarks>
        public DateTime CreationTime
        {
            get
            {
                return creationTime;
            }
            set
            {
                creationTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time when the message was last modified.
        /// </summary>
        /// <value>The last modification time.</value>
        /// <remarks>This property is initially set to the same value as the <see cref="CreationTime" /> property.</remarks>
        public DateTime LastModificationTime
        {
            get
            {
                return lastModificationTime;
            }
            set
            {
                lastModificationTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time when a message was delivered.
        /// </summary>
        /// <value>The message delivery time.</value>
        /// <remarks>This property describes the time the message was stored at the server, rather than the download time when the transport provider copied the message from the server to the local store.</remarks>
        public DateTime MessageDeliveryTime
        {
            get
            {
                return messageDeliveryTime;
            }
            set
            {
                messageDeliveryTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time the message sender submitted a message.
        /// </summary>
        /// <value>The client submit time.</value>
        public DateTime ClientSubmitTime
        {
            get
            {
                return clientSubmitTime;
            }
            set
            {
                clientSubmitTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time when a message sender wants a message delivered.
        /// </summary>
        /// <value>The deferred delivery time.</value>
        public DateTime DeferredDeliveryTime
        {
            get
            {
                return deferredDeliveryTime;
            }
            set
            {
                deferredDeliveryTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time the mail provider submitted a message.
        /// </summary>
        /// <value>The provider submit time.</value>
        public DateTime ProviderSubmitTime
        {
            get
            {
                return providerSubmitTime;
            }
            set
            {
                providerSubmitTime = value;
            }
        }

        /// <summary>
        /// Contains the report date and time.
        /// </summary>
        /// <value>The report time.</value>
        public DateTime ReportTime
        {
            get
            {
                return reportTime;
            }
            set
            {
                reportTime = value;
            }
        }

        /// <summary>
        /// Contains the time when the last verb was executed.
        /// </summary>
        /// <value>The last verb execution time.</value>
        public DateTime LastVerbExecutionTime
        {
            get
            {
                return lastVerbExecutionTime;
            }
            set
            {
                lastVerbExecutionTime = value;
            }
        }

        /// <summary>
        /// Contains report text.
        /// </summary>
        /// <value>The report text.</value>
        public string ReportText
        {
            get
            {
                return reportText;
            }
            set
            {
                reportText = value;
            }
        }

        /// <summary>
        /// Contains name of the person who created message.
        /// </summary>
        /// <value>The name of the creator.</value>
        public string CreatorName
        {
            get
            {
                return creatorName;
            }
            set
            {
                creatorName = value;
            }
        }

        /// <summary>
        /// Contains name of the person who modified message.
        /// </summary>
        /// <value>The last name of the modifier.</value>
        public string LastModifierName
        {
            get
            {
                return lastModifierName;
            }
            set
            {
                lastModifierName = value;
            }
        }

        /// <summary>
        /// Contains unique ID for the message.
        /// </summary>
        /// <value>The internet message identifier.</value>
        public string InternetMessageId
        {
            get
            {
                return internetMessageId;
            }
            set
            {
                internetMessageId = value;
            }
        }

        /// <summary>
        /// Contains the identifier of the message to which this message is a reply.
        /// </summary>
        /// <value>The in reply to.</value>
        public string InReplyTo
        {
            get
            {
                return inReplyTo;
            }
            set
            {
                inReplyTo = value;
            }
        }

        /// <summary>
        /// Contains Internet reference ID for the message.
        /// </summary>
        /// <value>The internet references.</value>
        public string InternetReferences
        {
            get
            {
                return internetReferences;
            }
            set
            {
                internetReferences = value;
            }
        }

        /// <summary>
        /// Contains the code page that is used for the message.
        /// </summary>
        /// <value>The message code page.</value>
        public long MessageCodePage
        {
            get
            {
                return messageCodePage;
            }
            set
            {
                messageCodePage = (uint)value;
            }
        }

        /// <summary>
        /// Contains a number that indicates which icon to use when you display a group of e-mail objects.
        /// </summary>
        /// <value>The index of the icon.</value>
        public long IconIndex
        {
            get
            {
                return iconIndex;
            }
            set
            {
                iconIndex = (uint)value;
            }
        }

        /// <summary>
        /// Contains the size of the body, subject, sender, and attachments.
        /// </summary>
        /// <value>The size.</value>
        public long Size
        {
            get
            {
                return messageSize;
            }
            set
            {
                messageSize = (uint)value;
            }
        }

        /// <summary>
        /// Indicates the code page used for the <see cref="Body" /> or the <see cref="BodyHtml" /> properties.
        /// </summary>
        /// <value>The internet code page.</value>
        public long InternetCodePage
        {
            get
            {
                return internetCodePage;
            }
            set
            {
                internetCodePage = (uint)value;
            }
        }

        /// <summary>
        /// Contains a binary value that indicates the relative position of this message within a conversation thread.
        /// </summary>
        /// <value>The index of the conversation.</value>
        public byte[] ConversationIndex
        {
            get
            {
                return conversationIndex;
            }
            set
            {
                conversationIndex = value;
            }
        }

        /// <summary>
        /// Contains true if message is invisible.
        /// </summary>
        /// <value><c>true</c> if this instance is hidden; otherwise, <c>false</c>.</value>
        public bool IsHidden
        {
            get
            {
                return isHidden;
            }
            set
            {
                isHidden = value;
            }
        }

        /// <summary>
        /// Contains true if message is read only.
        /// </summary>
        /// <value><c>true</c> if this instance is read only; otherwise, <c>false</c>.</value>
        public bool IsReadOnly
        {
            get
            {
                return isReadOnly;
            }
            set
            {
                isReadOnly = value;
            }
        }

        /// <summary>
        /// Contains true if message is system message.
        /// </summary>
        /// <value><c>true</c> if this instance is system; otherwise, <c>false</c>.</value>
        public bool IsSystem
        {
            get
            {
                return isSystem;
            }
            set
            {
                isSystem = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [disable full fidelity].
        /// </summary>
        /// <value><c>true</c> if [disable full fidelity]; otherwise, <c>false</c>.</value>
        public bool DisableFullFidelity
        {
            get
            {
                return disableFullFidelity;
            }
            set
            {
                disableFullFidelity = value;
            }
        }

        /// <summary>
        /// Contains true if a message contains at least one attachment.
        /// </summary>
        /// <value><c>true</c> if this instance has attachment; otherwise, <c>false</c>.</value>
        public bool HasAttachment
        {
            get
            {
                return hasAttachment;
            }
            set
            {
                hasAttachment = value;
            }
        }

        /// <summary>
        /// Contains true if the <see cref="RtfCompressed" /> property has the same text content as the <see cref="Body" /> property for this message.
        /// </summary>
        /// <value><c>true</c> if [RTF in synchronize]; otherwise, <c>false</c>.</value>
        public bool RtfInSync
        {
            get
            {
                return rtfInSync;
            }
            set
            {
                rtfInSync = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [read receipt requested].
        /// </summary>
        /// <value><c>true</c> if [read receipt requested]; otherwise, <c>false</c>.</value>
        public bool ReadReceiptRequested
        {
            get
            {
                return readReceiptRequested;
            }
            set
            {
                readReceiptRequested = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [delivery report requested].
        /// </summary>
        /// <value><c>true</c> if [delivery report requested]; otherwise, <c>false</c>.</value>
        public bool DeliveryReportRequested
        {
            get
            {
                return deliveryReportRequested;
            }
            set
            {
                deliveryReportRequested = value;
            }
        }

        /// <summary>
        /// Contains the Hypertext Markup Language (HTML) version of the message text.
        /// </summary>
        /// <value>The body HTML.</value>
        public byte[] BodyHtml
        {
            get
            {
                if (bodyHtml == null && BodyRtf != null) //large RTF body is slow on .net 4.0
                {
                    HtmlText htmlText = Util.ConvertRtfToHtml(BodyRtf);

                    if (htmlText != null && htmlText.Text != null)
                    {
                        return System.Text.Encoding.UTF8.GetBytes(htmlText.Text);
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return bodyHtml;
                }
            }
            set
            {
                bodyHtml = value;
            }
        }

        /// <summary>
        /// Contains the Hypertext Markup Language (HTML) version of the message text.
        /// </summary>
        /// <value>The body HTML text.</value>
        public string BodyHtmlText
        {
            get
            {
                if (bodyHtml != null)
                {
                    System.Text.Encoding internetCodePageEncoding = null;

                    try
                    {
                        internetCodePageEncoding = System.Text.Encoding.GetEncoding((int)this.internetCodePage);
                    }
                    catch //ignore error like non-existing code page
                    {
                    }


                    if (this.internetCodePage > 0 && internetCodePageEncoding != null)
                    {
                        string bodyHtmlText = internetCodePageEncoding.GetString(bodyHtml, 0, bodyHtml.Length);

                        return bodyHtmlText;
                    }
                    else if (Utf8Util.IsUtf8(bodyHtml, bodyHtml.Length))
                    {
                        System.Text.Encoding htmlEncoding = System.Text.Encoding.UTF8;

                        string bodyHtmlText = htmlEncoding.GetString(bodyHtml, 0, bodyHtml.Length);

                        return bodyHtmlText;
                    }
                    else
                    {
                        string bodyHtmlText = encoding.GetString(bodyHtml, 0, bodyHtml.Length);

                        System.Text.Encoding htmlEncoding = System.Text.Encoding.UTF7;

                        int htmlCharsetIndex = bodyHtmlText.IndexOf("charset=");

                        if (htmlCharsetIndex > 0)
                        {
                            int htmlCharsetEndIndex = bodyHtmlText.IndexOf("\"", htmlCharsetIndex);

                            if (htmlCharsetEndIndex != -1)
                            {
                                string htmlCharset = bodyHtmlText.Substring(htmlCharsetIndex + 8, htmlCharsetEndIndex - htmlCharsetIndex - 8);
                                htmlEncoding = System.Text.Encoding.GetEncoding(htmlCharset);
                            }
                        }

                        bodyHtmlText = htmlEncoding.GetString(bodyHtml, 0, bodyHtml.Length);

                        return bodyHtmlText;
                    }
                }
                else
                {
                    byte[] bodyRtf = BodyRtf;

                    if (bodyRtf != null && bodyRtf.Length > 0) //large RTF body is slow on .net 4.0
                    {
                        HtmlText htmlText = Util.ConvertRtfToHtml(bodyRtf);

                        if (htmlText != null)
                        {
                            return htmlText.Text;
                        }
                    }

                    return null;
                }
            }
            set
            {
                if (value != null)
                {
                    bodyHtml = System.Text.Encoding.UTF8.GetBytes(value);
                }
            }
        }

        /// <summary>
        /// Contains a value that indicates the message sender's opinion of the sensitivity of a message.
        /// </summary>
        /// <value>The sensitivity.</value>
        public Sensitivity Sensitivity
        {
            get
            {
                return sensitivity;
            }
            set
            {
                sensitivity = value;
            }
        }

        /// <summary>
        /// Contains the last verb executed.
        /// </summary>
        /// <value>The last verb executed.</value>
        public LastVerbExecuted LastVerbExecuted
        {
            get
            {
                return lastVerbExecuted;
            }
            set
            {
                lastVerbExecuted = value;
            }
        }

        /// <summary>
        /// Contains a value that indicates the message sender's opinion of the importance of a message.
        /// </summary>
        /// <value>The importance.</value>
        public Importance Importance
        {
            get
            {
                return importance;
            }
            set
            {
                importance = value;
            }
        }

        /// <summary>
        /// Contains the relative priority of a message.
        /// </summary>
        /// <value>The priority.</value>
        /// <remarks>This property and the <see cref="Importance" /> property should not be confused. Importance indicates a value to users, while priority indicates the order or speed at which the message should be sent by the messaging system software. Higher priority usually indicates a higher cost. Higher importance usually is associated with a different display by the user interface.</remarks>
        public Priority Priority
        {
            get
            {
                return priority;
            }
            set
            {
                priority = value;
            }
        }

        /// <summary>
        /// Specifies the flag icon of the message object.
        /// </summary>
        /// <value>The flag icon.</value>
        public FlagIcon FlagIcon
        {
            get
            {
                return flagIcon;
            }
            set
            {
                flagIcon = value;
            }
        }

        /// <summary>
        /// Specifies the flag state of the message object.
        /// </summary>
        /// <value>The flag status.</value>
        public FlagStatus FlagStatus
        {
            get
            {
                return flagStatus;
            }
            set
            {
                flagStatus = value;
            }
        }

        /// <summary>
        /// Contains the type of an object.
        /// </summary>
        /// <value>The type of the object.</value>
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
        /// Contains the address type for the messaging user who is represented by the user actually receiving the message.
        /// </summary>
        /// <value>The type of the received representing address.</value>
        public string ReceivedRepresentingAddressType
        {
            get
            {
                return receivedRepresentingAddressType;
            }
            set
            {
                receivedRepresentingAddressType = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address for the messaging user who is represented by the receiving user.
        /// </summary>
        /// <value>The received representing email address.</value>
        public string ReceivedRepresentingEmailAddress
        {
            get
            {
                return receivedRepresentingEmailAddress;
            }
            set
            {
                receivedRepresentingEmailAddress = value;
            }
        }

        /// <summary>
        /// Contains the entry identifier for the messaging user who is represented by the receiving user.
        /// </summary>
        /// <value>The received representing entry identifier.</value>
        public byte[] ReceivedRepresentingEntryId
        {
            get
            {
                return receivedRepresentingEntryId;
            }
            set
            {
                receivedRepresentingEntryId = value;
            }
        }

        /// <summary>
        /// Contains the display name for the messaging user who is represented by the receiving user.
        /// </summary>
        /// <value>The name of the received representing.</value>
        public string ReceivedRepresentingName
        {
            get
            {
                return receivedRepresentingName;
            }
            set
            {
                receivedRepresentingName = value;
            }
        }

        /// <summary>
        /// Contains the search key for the messaging user represented by the receiving user.
        /// </summary>
        /// <value>The received representing search key.</value>
        public byte[] ReceivedRepresentingSearchKey
        {
            get
            {
                return receivedRepresentingSearchKey;
            }
            set
            {
                receivedRepresentingSearchKey = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address type, such as SMTP, for the messaging user who actually receives the message.
        /// </summary>
        /// <value>The type of the received by address.</value>
        public string ReceivedByAddressType
        {
            get
            {
                return receivedByAddressType;
            }
            set
            {
                receivedByAddressType = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address for the messaging user who receives the message.
        /// </summary>
        /// <value>The received by email address.</value>
        public string ReceivedByEmailAddress
        {
            get
            {
                return receivedByEmailAddress;
            }
            set
            {
                receivedByEmailAddress = value;
            }
        }

        /// <summary>
        /// Contains the entry identifier of the messaging user who actually receives the message.
        /// </summary>
        /// <value>The received by entry identifier.</value>
        public byte[] ReceivedByEntryId
        {
            get
            {
                return receivedByEntryId;
            }
            set
            {
                receivedByEntryId = value;
            }
        }

        /// <summary>
        /// Contains the display name of the messaging user who receives the message.
        /// </summary>
        /// <value>The name of the received by.</value>
        public string ReceivedByName
        {
            get
            {
                return receivedByName;
            }
            set
            {
                receivedByName = value;
            }
        }

        /// <summary>
        /// Contains the search key of the messaging user who receives the message.
        /// </summary>
        /// <value>The received by search key.</value>
        public byte[] ReceivedBySearchKey
        {
            get
            {
                return receivedBySearchKey;
            }
            set
            {
                receivedBySearchKey = value;
            }
        }

        /// <summary>
        /// Contains the message sender's e-mail address type.
        /// </summary>
        /// <value>The type of the sender address.</value>
        public string SenderAddressType
        {
            get
            {
                return senderAddressType;
            }
            set
            {
                senderAddressType = value;
            }
        }

        /// <summary>
        /// Contains the message sender's e-mail address.
        /// </summary>
        /// <value>The sender email address.</value>
        public string SenderEmailAddress
        {
            get
            {
                return senderEmailAddress;
            }
            set
            {
                senderEmailAddress = value;
            }
        }

        /// <summary>
        /// Contains the message sender's e-mail address.
        /// </summary>
        /// <value>The sender email address.</value>
        public string SenderSmtpAddress
        {
            get
            {
                return senderSmtpAddress;
            }
            set
            {
                senderSmtpAddress = value;
            }
        }

        /// <summary>
        /// Contains the message sender's entry identifier.
        /// </summary>
        /// <value>The sender entry identifier.</value>
        public byte[] SenderEntryId
        {
            get
            {
                return senderEntryId;
            }
            set
            {
                senderEntryId = value;
            }
        }

        /// <summary>
        /// Contains the message sender's display name.
        /// </summary>
        /// <value>The name of the sender.</value>
        public string SenderName
        {
            get
            {
                return senderName;
            }
            set
            {
                senderName = value;
            }
        }

        /// <summary>
        /// Contains the message sender's search key.
        /// </summary>
        /// <value>The sender search key.</value>
        public byte[] SenderSearchKey
        {
            get
            {
                return senderSearchKey;
            }
            set
            {
                senderSearchKey = value;
            }
        }

        /// <summary>
        /// Contains the address type for the messaging user who is represented by the sender.
        /// </summary>
        /// <value>The type of the sent representing address.</value>
        public string SentRepresentingAddressType
        {
            get
            {
                return sentRepresentingAddressType;
            }
            set
            {
                sentRepresentingAddressType = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address for the messaging user who is represented by the sender.
        /// </summary>
        /// <value>The sent representing email address.</value>
        public string SentRepresentingEmailAddress
        {
            get
            {
                return sentRepresentingEmailAddress;
            }
            set
            {
                sentRepresentingEmailAddress = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address for the messaging user who is represented by the sender.
        /// </summary>
        /// <value>The sent representing email address.</value>
        public string SentRepresentingSmtpAddress
        {
            get
            {
                return sentRepresentingSmtpAddress;
            }
            set
            {
                sentRepresentingSmtpAddress = value;
            }
        }

        /// <summary>
        /// Contains the entry identifier for the messaging user represented by the sender.
        /// </summary>
        /// <value>The sent representing entry identifier.</value>
        public byte[] SentRepresentingEntryId
        {
            get
            {
                return sentRepresentingEntryId;
            }
            set
            {
                sentRepresentingEntryId = value;
            }
        }

        /// <summary>
        /// Contains the display name for the messaging user represented by the sender.
        /// </summary>
        /// <value>The name of the sent representing.</value>
        public string SentRepresentingName
        {
            get
            {
                return sentRepresentingName;
            }
            set
            {
                sentRepresentingName = value;
            }
        }

        /// <summary>
        /// Contains the search key for the messaging user represented by the sender.
        /// </summary>
        /// <value>The sent representing search key.</value>
        public byte[] SentRepresentingSearchKey
        {
            get
            {
                return sentRepresentingSearchKey;
            }
            set
            {
                sentRepresentingSearchKey = value;
            }
        }

        /// <summary>
        /// Contains transport-specific message envelope information.
        /// </summary>
        /// <value>The transport message headers.</value>
        public string TransportMessageHeaders
        {
            get
            {
                return transportMessageHeaders;
            }
            set
            {
                transportMessageHeaders = value;
            }
        }

        /// <summary>
        /// Contains a bitmask of flags that indicate the origin and current state of a message.
        /// </summary>
        /// <value>The message flags.</value>
        /// <remarks>This property is a nontransmittable message property exposed at both the sending and receiving ends of a transmission, with different values depending upon the client application or store provider involved. This property is initialized by the client or message store provider when a message is created and saved for the first time and then updated periodically by the message store provider, a transport provider, and the MAPI spooler as the message is processed and its state changes. This property exists on a message both before and after submission, and on all copies of the received message. Although it is not a recipient property, it is exposed differently to each recipient according to whether it has been read or modified by that recipient.</remarks>
        public IList<MessageFlag> MessageFlags
        {
            get
            {
                return messageFlags;
            }
        }

        /// <summary>
        /// Contains values that client applications should query to determine the characteristics of a message store.
        /// </summary>
        /// <value>The store support masks.</value>
        public IList<StoreSupportMask> StoreSupportMasks
        {
            get
            {
                return storeSupportMasks;
            }
        }

        /// <summary>
        /// Contains version number of Microsoft Office Outlook client.
        /// </summary>
        /// <value>The outlook version.</value>
        public string OutlookVersion
        {
            get
            {
                return outlookVersion;
            }
            set
            {
                outlookVersion = value;
            }
        }

        /// <summary>
        /// Contains internal version number of Microsoft Office Outlook client.
        /// </summary>
        /// <value>The outlook internal version.</value>
        public long OutlookInternalVersion
        {
            get
            {
                return outlookInternalVersion;
            }
            set
            {
                outlookInternalVersion = (uint)value;
            }
        }

        /// <summary>
        /// Contains the start date and time of a message.
        /// </summary>
        /// <value>The common start time.</value>
        public DateTime CommonStartTime
        {
            get
            {
                return commonStartTime;
            }
            set
            {
                commonStartTime = value;
            }
        }

        /// <summary>
        /// Contains the end date and time of a message.
        /// </summary>
        /// <value>The common end time.</value>
        public DateTime CommonEndTime
        {
            get
            {
                return commonEndTime;
            }
            set
            {
                commonEndTime = value;
            }
        }

        /// <summary>
        /// Contains the date and time specifying the date by which an e-mail message is due.
        /// </summary>
        /// <value>The flag due by.</value>
        public DateTime FlagDueBy
        {
            get
            {
                return flagDueBy;
            }
            set
            {
                flagDueBy = value;
            }
        }

        /// <summary>
        /// Contains the names of the companies associated with the contact item.
        /// </summary>
        /// <value>The companies.</value>
        public IList<string> Companies
        {
            get
            {
                return companies;
            }
        }

        /// <summary>
        /// Contains the names of the contacts associated with the item.
        /// </summary>
        /// <value>The contact names.</value>
        public IList<string> ContactNames
        {
            get
            {
                return contactNames;
            }
        }

        /// <summary>
        /// Contains the categories associated with a message.
        /// </summary>
        /// <value>The keywords.</value>
        public IList<string> Keywords
        {
            get
            {
                return keywords;
            }
        }

        /// <summary>
        /// Contains the categories associated with a message.
        /// </summary>
        /// <value>The categories.</value>
        public IList<string> Categories
        {
            get
            {
                return keywords;
            }
        }

        /// <summary>
        /// Contains the billing information associated with a message.
        /// </summary>
        /// <value>The billing information.</value>
        public string BillingInformation
        {
            get
            {
                return billingInformation;
            }
            set
            {
                billingInformation = value;
            }
        }

        /// <summary>
        /// Contains free-form string value and can be used to store mileage information associated with the message.
        /// </summary>
        /// <value>The mileage.</value>
        public string Mileage
        {
            get
            {
                return mileage;
            }
            set
            {
                mileage = value;
            }
        }

        /// <summary>
        /// Contains account name or email address.
        /// </summary>
        /// <value>The name of the internet account.</value>
        public string InternetAccountName
        {
            get
            {
                return internetAccountName;
            }
            set
            {
                internetAccountName = value;
            }
        }

        /// <summary>
        /// Contains the path and file name of the sound file to play when the reminder occurs for the appointment, mail message, or task.
        /// </summary>
        /// <value>The reminder sound file.</value>
        public string ReminderSoundFile
        {
            get
            {
                return reminderSoundFile;
            }
            set
            {
                reminderSoundFile = value;
            }
        }

        /// <summary>
        /// Contains true if message is marked as private.
        /// </summary>
        /// <value><c>true</c> if this instance is private; otherwise, <c>false</c>.</value>
        public bool IsPrivate
        {
            get
            {
                return isPrivate;
            }
            set
            {
                isPrivate = value;
            }
        }

        /// <summary>
        /// Contains true if the reminder overrides the default reminder behavior for the appointment, mail item, or task.
        /// </summary>
        /// <value><c>true</c> if [reminder override default]; otherwise, <c>false</c>.</value>
        public bool ReminderOverrideDefault
        {
            get
            {
                return reminderOverrideDefault;
            }
            set
            {
                reminderOverrideDefault = value;
            }
        }

        /// <summary>
        /// Contains true if the reminder should play a sound when it occurs for this appointment or task.
        /// </summary>
        /// <value><c>true</c> if [reminder play sound]; otherwise, <c>false</c>.</value>
        public bool ReminderPlaySound
        {
            get
            {
                return reminderPlaySound;
            }
            set
            {
                reminderPlaySound = value;
            }
        }

        /// <summary>
        /// Contains appointment's the start date and time.
        /// </summary>
        /// <value>The appointment start time.</value>
        public DateTime AppointmentStartTime
        {
            get
            {
                return appointmentStartTime;
            }
            set
            {
                appointmentStartTime = value;
            }
        }

        /// <summary>
        /// Contains appointment's the end date and time.
        /// </summary>
        /// <value>The appointment end time.</value>
        public DateTime AppointmentEndTime
        {
            get
            {
                return appointmentEndTime;
            }
            set
            {
                appointmentEndTime = value;
            }
        }

        /// <summary>
        /// Contains appointment's location.
        /// </summary>
        /// <value>The location.</value>
        public string Location
        {
            get
            {
                return location;
            }
            set
            {
                location = value;
            }
        }

        /// <summary>
        /// Contains appointment's message class.
        /// </summary>
        /// <value>The appointment message class.</value>
        public string AppointmentMessageClass
        {
            get
            {
                return appointmentMessageClass;
            }
            set
            {
                appointmentMessageClass = value;
            }
        }

        /// <summary>
        /// Contains appointment's time zone.
        /// </summary>
        /// <value>The time zone.</value>
        public string TimeZone
        {
            get
            {
                return timeZone;
            }
            set
            {
                timeZone = value;
            }
        }

        /// <summary>
        /// Contains recurring pattern description.
        /// </summary>
        /// <value>The recurrence pattern description.</value>
        public string RecurrencePatternDescription
        {
            get
            {
                return recurrencePatternDescription;
            }
            set
            {
                recurrencePatternDescription = value;
            }
        }

        /// <summary>
        /// Contains appoinmtment or task recurring pattern.
        /// </summary>
        /// <value>The recurrence pattern.</value>
        public RecurrencePattern RecurrencePattern
        {
            get
            {
                return recurrencePattern;
            }
        }


        /// <summary>
        /// Contains message's global unique id.
        /// </summary>
        /// <value>The unique identifier.</value>
        public byte[] Guid
        {
            get
            {
                return guid;
            }
            set
            {
                guid = value;
            }
        }

        /// <summary>
        /// Contains appointment's label color.
        /// </summary>
        /// <value>The label.</value>
        public long Label
        {
            get
            {
                return label;
            }
            set
            {
                label = (int)value;
            }
        }

        /// <summary>
        /// Contains appointment's duration in minutes.
        /// </summary>
        /// <value>The duration.</value>
        public long Duration
        {
            get
            {
                return duration;
            }
            set
            {
                duration = (uint)value;
            }
        }

        /// <summary>
        /// Contains appointment's busy status.
        /// </summary>
        /// <value>The busy status.</value>
        public BusyStatus BusyStatus
        {
            get
            {
                return busyStatus;
            }
            set
            {
                busyStatus = value;
            }
        }

        /// <summary>
        /// Contains the status of the meeting.
        /// </summary>
        /// <value>The meeting status.</value>
        public MeetingStatus MeetingStatus
        {
            get
            {
                return meetingStatus;
            }
            set
            {
                meetingStatus = value;
            }
        }

        /// <summary>
        /// Contains the response to a meeting request.
        /// </summary>
        /// <value>The response status.</value>
        public ResponseStatus ResponseStatus
        {
            get
            {
                return responseStatus;
            }
            set
            {
                responseStatus = value;
            }
        }

        /// <summary>
        /// Contains the recurrence pattern type.
        /// </summary>
        /// <value>The type of the recurrence.</value>
        public RecurrenceType RecurrenceType
        {
            get
            {
                return recurrenceType;
            }
            set
            {
                recurrenceType = value;
            }
        }

        /// <summary>
        /// Contains task's owner name.
        /// </summary>
        /// <value>The owner.</value>
        public string Owner
        {
            get
            {
                return owner;
            }
            set
            {
                owner = value;
            }
        }

        /// <summary>
        /// Contains task's delegator name.
        /// </summary>
        /// <value>The delegator.</value>
        public string Delegator
        {
            get
            {
                return delegator;
            }
            set
            {
                delegator = value;
            }
        }

        /// <summary>
        /// Contains the percentage of the task completed at the current date and time.
        /// </summary>
        /// <value>The percent complete.</value>
        public double PercentComplete
        {
            get
            {
                return percentComplete;
            }
            set
            {
                percentComplete = value;
            }
        }

        /// <summary>
        /// Contains the actual effort (in minutes) spent on the task.
        /// </summary>
        /// <value>The actual work.</value>
        public long ActualWork
        {
            get
            {
                return actualWork;
            }
            set
            {
                actualWork = (uint)value;
            }
        }

        /// <summary>
        /// Contains the total work for the task.
        /// </summary>
        /// <value>The total work.</value>
        public long TotalWork
        {
            get
            {
                return totalWork;
            }
            set
            {
                totalWork = (uint)value;
            }
        }

        /// <summary>
        /// Contains true if the task is a team task.
        /// </summary>
        /// <value><c>true</c> if this instance is team task; otherwise, <c>false</c>.</value>
        public bool IsTeamTask
        {
            get
            {
                return isTeamTask;
            }
            set
            {
                isTeamTask = value;
            }
        }

        /// <summary>
        /// Contains true if the task is complete.
        /// </summary>
        /// <value><c>true</c> if this instance is complete; otherwise, <c>false</c>.</value>
        public bool IsComplete
        {
            get
            {
                return isComplete;
            }
            set
            {
                isComplete = value;
            }
        }

        /// <summary>
        /// Contains true if the task or appointment is recurring.
        /// </summary>
        /// <value><c>true</c> if this instance is recurring; otherwise, <c>false</c>.</value>
        public bool IsRecurring
        {
            get
            {
                return isRecurring;
            }
            set
            {
                isRecurring = value;
            }
        }

        /// <summary>
        /// Contains true if the appointment is all day event.
        /// </summary>
        /// <value><c>true</c> if this instance is all day event; otherwise, <c>false</c>.</value>
        public bool IsAllDayEvent
        {
            get
            {
                return isAllDayEvent;
            }
            set
            {
                isAllDayEvent = value;
            }
        }

        /// <summary>
        /// Contains true if a reminder has been set for this appointment, e-mail item, or task.
        /// </summary>
        /// <value><c>true</c> if this instance is reminder set; otherwise, <c>false</c>.</value>
        public bool IsReminderSet
        {
            get
            {
                return isReminderSet;
            }
            set
            {
                isReminderSet = value;
            }
        }

        /// <summary>
        /// Contains the date and time at which the reminder should occur for the specified item.
        /// </summary>
        /// <value>The reminder time.</value>
        public DateTime ReminderTime
        {
            get
            {
                return reminderTime;
            }
            set
            {
                reminderTime = value;
            }
        }

        /// <summary>
        /// Contains the number of minutes the reminder should occur prior to the start of the appointment.
        /// </summary>
        /// <value>The reminder minutes before start.</value>
        public long ReminderMinutesBeforeStart
        {
            get
            {
                return reminderMinutesBeforeStart;
            }
            set
            {
                reminderMinutesBeforeStart = (uint)value;
            }
        }

        /// <summary>
        /// Contains task's the start date and time.
        /// </summary>
        /// <value>The task start date.</value>
        public DateTime TaskStartDate
        {
            get
            {
                return taskStartDate;
            }
            set
            {
                taskStartDate = value;
            }
        }

        /// <summary>
        /// Contains task's the due date and time.
        /// </summary>
        /// <value>The task due date.</value>
        public DateTime TaskDueDate
        {
            get
            {
                return taskDueDate;
            }
            set
            {
                taskDueDate = value;
            }
        }

        /// <summary>
        /// Contains the completion date of the task.
        /// </summary>
        /// <value>The date completed.</value>
        public DateTime DateCompleted
        {
            get
            {
                return dateCompleted;
            }
            set
            {
                dateCompleted = value;
            }
        }

        /// <summary>
        /// Contains the status of the task.
        /// </summary>
        /// <value>The task status.</value>
        public TaskStatus TaskStatus
        {
            get
            {
                return taskStatus;
            }
            set
            {
                taskStatus = value;
            }
        }

        /// <summary>
        /// Contains the ownership state of the task.
        /// </summary>
        /// <value>The task ownership.</value>
        public TaskOwnership TaskOwnership
        {
            get
            {
                return taskOwnership;
            }
            set
            {
                taskOwnership = value;
            }
        }

        /// <summary>
        /// Contains the delegation state of a task.
        /// </summary>
        /// <value>The state of the task delegation.</value>
        public TaskDelegationState TaskDelegationState
        {
            get
            {
                return taskDelegationState;
            }
            set
            {
                taskDelegationState = value;
            }
        }

        /// <summary>
        /// Contains height of the note item.
        /// </summary>
        /// <value>The height of the note.</value>
        public long NoteHeight
        {
            get
            {
                return noteHeight;
            }
            set
            {
                noteHeight = (uint)value;
            }
        }

        /// <summary>
        /// Contains width of the note item.
        /// </summary>
        /// <value>The width of the note.</value>
        public long NoteWidth
        {
            get
            {
                return noteWidth;
            }
            set
            {
                noteWidth = (uint)value;
            }
        }

        /// <summary>
        /// Contains top position of the note item.
        /// </summary>
        /// <value>The note top.</value>
        public long NoteTop
        {
            get
            {
                return noteTop;
            }
            set
            {
                noteTop = (uint)value;
            }
        }

        /// <summary>
        /// Contains left position of the note item.
        /// </summary>
        /// <value>The note left.</value>
        public long NoteLeft
        {
            get
            {
                return noteLeft;
            }
            set
            {
                noteLeft = (uint)value;
            }
        }

        /// <summary>
        /// Contains background color of the note item.
        /// </summary>
        /// <value>The color of the note.</value>
        public NoteColor NoteColor
        {
            get
            {
                return noteColor;
            }
            set
            {
                noteColor = value;
            }
        }

        /// <summary>
        /// Contains journal's the start date and time.
        /// </summary>
        /// <value>The journal start time.</value>
        public DateTime JournalStartTime
        {
            get
            {
                return journalStartTime;
            }
            set
            {
                journalStartTime = value;
            }
        }

        /// <summary>
        /// Contains journal's the end date and time.
        /// </summary>
        /// <value>The journal end time.</value>
        public DateTime JournalEndTime
        {
            get
            {
                return journalEndTime;
            }
            set
            {
                journalEndTime = value;
            }
        }

        /// <summary>
        /// Contains the type of the journal item.
        /// </summary>
        /// <value>The type of the journal.</value>
        public string JournalType
        {
            get
            {
                return journalType;
            }
            set
            {
                journalType = value;
            }
        }

        /// <summary>
        /// Contains the type description of the journal item.
        /// </summary>
        /// <value>The journal type description.</value>
        public string JournalTypeDescription
        {
            get
            {
                return journalTypeDescription;
            }
            set
            {
                journalTypeDescription = value;
            }
        }

        /// <summary>
        /// Contains journal's the duration in minutes.
        /// </summary>
        /// <value>The duration of the journal.</value>
        public long JournalDuration
        {
            get
            {
                return journalDuration;
            }
            set
            {
                journalDuration = (uint)value;
            }
        }

        /// <summary>
        /// Contains the birthday date for the contact.
        /// </summary>
        /// <value>The birthday.</value>
        public DateTime Birthday
        {
            get
            {
                return birthday;
            }
            set
            {
                birthday = value;
            }
        }

        /// <summary>
        /// Contains the names of the children of the contact.
        /// </summary>
        /// <value>The children names.</value>
        public IList<string> ChildrenNames
        {
            get
            {
                return childrenNames;
            }
        }

        /// <summary>
        /// Contains the name of assistent of the contact.
        /// </summary>
        /// <value>The name of the assistent.</value>
        public string AssistentName
        {
            get
            {
                return assistentName;
            }
            set
            {
                assistentName = value;
            }
        }

        /// <summary>
        /// Contains assistent's phone number of the contact.
        /// </summary>
        /// <value>The assistent phone.</value>
        public string AssistentPhone
        {
            get
            {
                return assistentPhone;
            }
            set
            {
                assistentPhone = value;
            }
        }

        /// <summary>
        /// Contains the first business telephone number for the contact.
        /// </summary>
        /// <value>The business phone.</value>
        public string BusinessPhone
        {
            get
            {
                return businessPhone;
            }
            set
            {
                businessPhone = value;
            }
        }

        /// <summary>
        /// Contains the business fax number for the contact.
        /// </summary>
        /// <value>The business fax.</value>
        public string BusinessFax
        {
            get
            {
                return businessFax;
            }
            set
            {
                businessFax = value;
            }
        }

        /// <summary>
        /// Contains the url of the business Web page for the contact.
        /// </summary>
        /// <value>The business home page.</value>
        public string BusinessHomePage
        {
            get
            {
                return businessHomePage;
            }
            set
            {
                businessHomePage = value;
            }
        }

        /// <summary>
        /// Contains the callback telephone number for the contact.
        /// </summary>
        /// <value>The callback phone.</value>
        public string CallbackPhone
        {
            get
            {
                return callbackPhone;
            }
            set
            {
                callbackPhone = value;
            }
        }

        /// <summary>
        /// Contains the car telephone number for the contact.
        /// </summary>
        /// <value>The car phone.</value>
        public string CarPhone
        {
            get
            {
                return carPhone;
            }
            set
            {
                carPhone = value;
            }
        }

        /// <summary>
        /// Contains the mobile telephone number for the contact.
        /// </summary>
        /// <value>The cellular phone.</value>
        public string CellularPhone
        {
            get
            {
                return cellularPhone;
            }
            set
            {
                cellularPhone = value;
            }
        }

        /// <summary>
        /// Contains the company main telephone number for the contact.
        /// </summary>
        /// <value>The company main phone.</value>
        public string CompanyMainPhone
        {
            get
            {
                return companyMainPhone;
            }
            set
            {
                companyMainPhone = value;
            }
        }

        /// <summary>
        /// Contains the company name for the contact.
        /// </summary>
        /// <value>The name of the company.</value>
        public string CompanyName
        {
            get
            {
                return companyName;
            }
            set
            {
                companyName = value;
            }
        }

        /// <summary>
        /// Contains the name of the computer network for the contact.
        /// </summary>
        /// <value>The name of the computer network.</value>
        public string ComputerNetworkName
        {
            get
            {
                return computerNetworkName;
            }
            set
            {
                computerNetworkName = value;
            }
        }

        /// <summary>
        /// Contains the country/region code portion of the business address for the contact.
        /// </summary>
        /// <value>The business address country.</value>
        public string BusinessAddressCountry
        {
            get
            {
                return businessAddressCountry;
            }
            set
            {
                businessAddressCountry = value;
            }
        }

        /// <summary>
        /// Contains the customer ID for the contact.
        /// </summary>
        /// <value>The customer identifier.</value>
        public string CustomerId
        {
            get
            {
                return customerId;
            }
            set
            {
                customerId = value;
            }
        }

        /// <summary>
        /// Contains the department name for the contact.
        /// </summary>
        /// <value>The name of the department.</value>
        public string DepartmentName
        {
            get
            {
                return departmentName;
            }
            set
            {
                departmentName = value;
            }
        }

        /// <summary>
        /// Contains display name.
        /// </summary>
        /// <value>The display name.</value>
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
        /// Contains display name prefix.
        /// </summary>
        /// <value>The display name prefix.</value>
        public string DisplayNamePrefix
        {
            get
            {
                return displayNamePrefix;
            }
            set
            {
                displayNamePrefix = value;
            }
        }

        /// <summary>
        /// Contains the FTP site entry for the contact.
        /// </summary>
        /// <value>The FTP site.</value>
        public string FtpSite
        {
            get
            {
                return ftpSite;
            }
            set
            {
                ftpSite = value;
            }
        }

        /// <summary>
        /// Contains the generation for the contact.
        /// </summary>
        /// <value>The generation.</value>
        public string Generation
        {
            get
            {
                return generation;
            }
            set
            {
                generation = value;
            }
        }

        /// <summary>
        /// Contains the given name for the contact.
        /// </summary>
        /// <value>The name of the given.</value>
        public string GivenName
        {
            get
            {
                return givenName;
            }
            set
            {
                givenName = value;
            }
        }

        /// <summary>
        /// Contains the government ID number for the contact.
        /// </summary>
        /// <value>The government identifier.</value>
        public string GovernmentId
        {
            get
            {
                return governmentId;
            }
            set
            {
                governmentId = value;
            }
        }

        /// <summary>
        /// Contains the hobby names for the contact.
        /// </summary>
        /// <value>The hobbies.</value>
        public string Hobbies
        {
            get
            {
                return hobbies;
            }
            set
            {
                hobbies = value;
            }
        }

        /// <summary>
        /// Contains the second home telephone number for the contact.
        /// </summary>
        /// <value>The home phone2.</value>
        public string HomePhone2
        {
            get
            {
                return homePhone2;
            }
            set
            {
                homePhone2 = value;
            }
        }

        /// <summary>
        /// Contains the city portion of the home address for the contact.
        /// </summary>
        /// <value>The home address city.</value>
        public string HomeAddressCity
        {
            get
            {
                return homeAddressCity;
            }
            set
            {
                homeAddressCity = value;
            }
        }

        /// <summary>
        /// Contains the country/region portion of the home address for the contact.
        /// </summary>
        /// <value>The home address country.</value>
        public string HomeAddressCountry
        {
            get
            {
                return homeAddressCountry;
            }
            set
            {
                homeAddressCountry = value;
            }
        }

        /// <summary>
        /// Contains the postal code portion of the home address for the contact.
        /// </summary>
        /// <value>The home address postal code.</value>
        public string HomeAddressPostalCode
        {
            get
            {
                return homeAddressPostalCode;
            }
            set
            {
                homeAddressPostalCode = value;
            }
        }

        /// <summary>
        /// Contains the post office box number portion of the home address for the contact.
        /// </summary>
        /// <value>The home address post office box.</value>
        public string HomeAddressPostOfficeBox
        {
            get
            {
                return homeAddressPostOfficeBox;
            }
            set
            {
                homeAddressPostOfficeBox = value;
            }
        }

        /// <summary>
        /// Contains the state portion of the home address for the contact.
        /// </summary>
        /// <value>The state of the home address.</value>
        public string HomeAddressState
        {
            get
            {
                return homeAddressState;
            }
            set
            {
                homeAddressState = value;
            }
        }

        /// <summary>
        /// Contains the street portion of the home address for the contact.
        /// </summary>
        /// <value>The home address street.</value>
        public string HomeAddressStreet
        {
            get
            {
                return homeAddressStreet;
            }
            set
            {
                homeAddressStreet = value;
            }
        }

        /// <summary>
        /// Contains the home fax number for the contact.
        /// </summary>
        /// <value>The home fax.</value>
        public string HomeFax
        {
            get
            {
                return homeFax;
            }
            set
            {
                homeFax = value;
            }
        }

        /// <summary>
        /// Contains the first home telephone number for the contact.
        /// </summary>
        /// <value>The home phone.</value>
        public string HomePhone
        {
            get
            {
                return homePhone;
            }
            set
            {
                homePhone = value;
            }
        }

        /// <summary>
        /// Contains the initials for the contact.
        /// </summary>
        /// <value>The initials.</value>
        public string Initials
        {
            get
            {
                return initials;
            }
            set
            {
                initials = value;
            }
        }

        /// <summary>
        /// Contains the ISDN number for the contact.
        /// </summary>
        /// <value>The isdn.</value>
        public string Isdn
        {
            get
            {
                return isdn;
            }
            set
            {
                isdn = value;
            }
        }

        /// <summary>
        /// Contains the city name portion of the business address for the contact.
        /// </summary>
        /// <value>The business address city.</value>
        public string BusinessAddressCity
        {
            get
            {
                return businessAddressCity;
            }
            set
            {
                businessAddressCity = value;
            }
        }

        /// <summary>
        /// Contains the manager name for the contact.
        /// </summary>
        /// <value>The name of the manager.</value>
        public string ManagerName
        {
            get
            {
                return managerName;
            }
            set
            {
                managerName = value;
            }
        }

        /// <summary>
        /// Contains the middle name for the contact.
        /// </summary>
        /// <value>The name of the middle.</value>
        public string MiddleName
        {
            get
            {
                return middleName;
            }
            set
            {
                middleName = value;
            }
        }

        /// <summary>
        /// Contains the nickname for the contact.
        /// </summary>
        /// <value>The nickname.</value>
        public string Nickname
        {
            get
            {
                return nickname;
            }
            set
            {
                nickname = value;
            }
        }

        /// <summary>
        /// Contains the specific office location for the contact.
        /// </summary>
        /// <value>The office location.</value>
        public string OfficeLocation
        {
            get
            {
                return officeLocation;
            }
            set
            {
                officeLocation = value;
            }
        }

        /// <summary>
        /// Contains the second business telephone number for the contact.
        /// </summary>
        /// <value>The business phone2.</value>
        public string BusinessPhone2
        {
            get
            {
                return businessPhone2;
            }
            set
            {
                businessPhone2 = value;
            }
        }

        /// <summary>
        /// Contains the city portion of the other address for the contact.
        /// </summary>
        /// <value>The other address city.</value>
        public string OtherAddressCity
        {
            get
            {
                return otherAddressCity;
            }
            set
            {
                otherAddressCity = value;
            }
        }

        /// <summary>
        /// Contains the country/region portion of the other address for the contact.
        /// </summary>
        /// <value>The other address country.</value>
        public string OtherAddressCountry
        {
            get
            {
                return otherAddressCountry;
            }
            set
            {
                otherAddressCountry = value;
            }
        }

        /// <summary>
        /// Contains the postal code portion of the other address for the contact.
        /// </summary>
        /// <value>The other address postal code.</value>
        public string OtherAddressPostalCode
        {
            get
            {
                return otherAddressPostalCode;
            }
            set
            {
                otherAddressPostalCode = value;
            }
        }

        /// <summary>
        /// Contains the state portion of the other address for the contact.
        /// </summary>
        /// <value>The state of the other address.</value>
        public string OtherAddressState
        {
            get
            {
                return otherAddressState;
            }
            set
            {
                otherAddressState = value;
            }
        }

        /// <summary>
        /// Contains the street portion of the other address for the contact.
        /// </summary>
        /// <value>The other address street.</value>
        public string OtherAddressStreet
        {
            get
            {
                return otherAddressStreet;
            }
            set
            {
                otherAddressStreet = value;
            }
        }

        /// <summary>
        /// Contains the other telephone number for the contact.
        /// </summary>
        /// <value>The other phone.</value>
        public string OtherPhone
        {
            get
            {
                return otherPhone;
            }
            set
            {
                otherPhone = value;
            }
        }

        /// <summary>
        /// Contains the pager number for the contact.
        /// </summary>
        /// <value>The pager.</value>
        public string Pager
        {
            get
            {
                return pager;
            }
            set
            {
                pager = value;
            }
        }

        /// <summary>
        /// Contains the url of the personal Web page for the contact.
        /// </summary>
        /// <value>The personal home page.</value>
        public string PersonalHomePage
        {
            get
            {
                return personalHomePage;
            }
            set
            {
                personalHomePage = value;
            }
        }

        /// <summary>
        /// Contains the postal address for the contact.
        /// </summary>
        /// <value>The postal address.</value>
        public string PostalAddress
        {
            get
            {
                return postalAddress;
            }
            set
            {
                postalAddress = value;
            }
        }

        /// <summary>
        /// Contains the postal code (zip code) portion of the business address for the contact.
        /// </summary>
        /// <value>The business address postal code.</value>
        public string BusinessAddressPostalCode
        {
            get
            {
                return businessAddressPostalCode;
            }
            set
            {
                businessAddressPostalCode = value;
            }
        }

        /// <summary>
        /// Contains the post office box number portion of the business address for the contact.
        /// </summary>
        /// <value>The business address post office box.</value>
        public string BusinessAddressPostOfficeBox
        {
            get
            {
                return businessAddressPostOfficeBox;
            }
            set
            {
                businessAddressPostOfficeBox = value;
            }
        }

        /// <summary>
        /// Contains the state code portion of the business address for the contact.
        /// </summary>
        /// <value>The state of the business address.</value>
        public string BusinessAddressState
        {
            get
            {
                return businessAddressState;
            }
            set
            {
                businessAddressState = value;
            }
        }

        /// <summary>
        /// Contains the street address portion of the business address for the contact.
        /// </summary>
        /// <value>The business address street.</value>
        public string BusinessAddressStreet
        {
            get
            {
                return businessAddressStreet;
            }
            set
            {
                businessAddressStreet = value;
            }
        }

        /// <summary>
        /// Contains the primary fax number for the contact.
        /// </summary>
        /// <value>The primary fax.</value>
        public string PrimaryFax
        {
            get
            {
                return primaryFax;
            }
            set
            {
                primaryFax = value;
            }
        }

        /// <summary>
        /// Contains the primary telephone number for the contact.
        /// </summary>
        /// <value>The primary phone.</value>
        public string PrimaryPhone
        {
            get
            {
                return primaryPhone;
            }
            set
            {
                primaryPhone = value;
            }
        }

        /// <summary>
        /// Contains the profession for the contact.
        /// </summary>
        /// <value>The profession.</value>
        public string Profession
        {
            get
            {
                return profession;
            }
            set
            {
                profession = value;
            }
        }

        /// <summary>
        /// Contains the radio telephone number for the contact.
        /// </summary>
        /// <value>The radio phone.</value>
        public string RadioPhone
        {
            get
            {
                return radioPhone;
            }
            set
            {
                radioPhone = value;
            }
        }

        /// <summary>
        /// Contains the spouse name for the contact.
        /// </summary>
        /// <value>The name of the spouse.</value>
        public string SpouseName
        {
            get
            {
                return spouseName;
            }
            set
            {
                spouseName = value;
            }
        }

        /// <summary>
        /// Contains the last name for the contact.
        /// </summary>
        /// <value>The surname.</value>
        public string Surname
        {
            get
            {
                return surname;
            }
            set
            {
                surname = value;
            }
        }

        /// <summary>
        /// Contains the telex number for the contact.
        /// </summary>
        /// <value>The telex.</value>
        public string Telex
        {
            get
            {
                return telex;
            }
            set
            {
                telex = value;
            }
        }

        /// <summary>
        /// Contains the title for the contact.
        /// </summary>
        /// <value>The title.</value>
        public string Title
        {
            get
            {
                return title;
            }
            set
            {
                title = value;
            }
        }

        /// <summary>
        /// Contains the TTY/TDD telephone number for the contact.
        /// </summary>
        /// <value>The tty TDD phone.</value>
        public string TtyTddPhone
        {
            get
            {
                return ttyTddPhone;
            }
            set
            {
                ttyTddPhone = value;
            }
        }

        /// <summary>
        /// Contains the wedding anniversary date for the contact.
        /// </summary>
        /// <value>The wedding anniversary.</value>
        public DateTime WeddingAnniversary
        {
            get
            {
                return weddingAnniversary;
            }
            set
            {
                weddingAnniversary = value;
            }
        }

        /// <summary>
        /// Contains the gender of the contact.
        /// </summary>
        /// <value>The gender.</value>
        public Gender Gender
        {
            get
            {
                return gender;
            }
            set
            {
                gender = value;
            }
        }

        /// <summary>
        /// Contains the type of the mailing address for the contact.
        /// </summary>
        /// <value>The selected mailing address.</value>
        public SelectedMailingAddress SelectedMailingAddress
        {
            get
            {
                return selectedMailingAddress;
            }
            set
            {
                selectedMailingAddress = value;
            }
        }

        /// <summary>
        /// Contains true if the contact has picture.
        /// </summary>
        /// <value><c>true</c> if [contact has picture]; otherwise, <c>false</c>.</value>
        public bool ContactHasPicture
        {
            get
            {
                return contactHasPicture;
            }
            set
            {
                contactHasPicture = value;
            }
        }

        /// <summary>
        /// Contains the default keyword string assigned to the contact when it is filed.
        /// </summary>
        /// <value>The file as.</value>
        public string FileAs
        {
            get
            {
                return fileAs;
            }
            set
            {
                fileAs = value;
            }
        }

        /// <summary>
        /// Contains the instant messenger address for the contact.
        /// </summary>
        /// <value>The instant messenger address.</value>
        public string InstantMessengerAddress
        {
            get
            {
                return instantMessengerAddress;
            }
            set
            {
                instantMessengerAddress = value;
            }
        }

        /// <summary>
        /// Contains the url location of the user's free-busy information in vCard Free-Busy standard format.
        /// </summary>
        /// <value>The internet free busy address.</value>
        public string InternetFreeBusyAddress
        {
            get
            {
                return internetFreeBusyAddress;
            }
            set
            {
                internetFreeBusyAddress = value;
            }
        }

        /// <summary>
        /// Contains the whole, unparsed business address for the contact.
        /// </summary>
        /// <value>The business address.</value>
        public string BusinessAddress
        {
            get
            {
                return businessAddress;
            }
            set
            {
                businessAddress = value;
            }
        }

        /// <summary>
        /// Contains the whole, unparsed home address for the contact.
        /// </summary>
        /// <value>The home address.</value>
        public string HomeAddress
        {
            get
            {
                return homeAddress;
            }
            set
            {
                homeAddress = value;
            }
        }

        /// <summary>
        /// Contains the whole, unparsed other address for the contact.
        /// </summary>
        /// <value>The other address.</value>
        public string OtherAddress
        {
            get
            {
                return otherAddress;
            }
            set
            {
                otherAddress = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address of the first e-mail entry for the contact.
        /// </summary>
        /// <value>The email1 address.</value>
        public string Email1Address
        {
            get
            {
                return email1Address;
            }
            set
            {
                email1Address = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address of the second e-mail entry for the contact.
        /// </summary>
        /// <value>The email2 address.</value>
        public string Email2Address
        {
            get
            {
                return email2Address;
            }
            set
            {
                email2Address = value;
            }
        }

        /// <summary>
        /// Contains the e-mail address of the third e-mail entry for the contact.
        /// </summary>
        /// <value>The email3 address.</value>
        public string Email3Address
        {
            get
            {
                return email3Address;
            }
            set
            {
                email3Address = value;
            }
        }

        /// <summary>
        /// Contains the display name of the first e-mail address for the contact.
        /// </summary>
        /// <value>The display name of the email1.</value>
        public string Email1DisplayName
        {
            get
            {
                return email1DisplayName;
            }
            set
            {
                email1DisplayName = value;
            }
        }

        /// <summary>
        /// Contains the display name of the second e-mail address for the contact.
        /// </summary>
        /// <value>The display name of the email2.</value>
        public string Email2DisplayName
        {
            get
            {
                return email2DisplayName;
            }
            set
            {
                email2DisplayName = value;
            }
        }

        /// <summary>
        /// Contains the display name of the third e-mail address for the contact.
        /// </summary>
        /// <value>The display name of the email3.</value>
        public string Email3DisplayName
        {
            get
            {
                return email3DisplayName;
            }
            set
            {
                email3DisplayName = value;
            }
        }

        /// <summary>
        /// Contains the display as name of the first e-mail address for the contact.
        /// </summary>
        /// <value>The email1 display as.</value>
        public string Email1DisplayAs
        {
            get
            {
                return email1DisplayAs;
            }
            set
            {
                email1DisplayAs = value;
            }
        }

        /// <summary>
        /// Contains the display as name of the second e-mail address for the contact.
        /// </summary>
        /// <value>The email2 display as.</value>
        public string Email2DisplayAs
        {
            get
            {
                return email2DisplayAs;
            }
            set
            {
                email2DisplayAs = value;
            }
        }

        /// <summary>
        /// Contains the display as name of the third e-mail address for the contact.
        /// </summary>
        /// <value>The email3 display as.</value>
        public string Email3DisplayAs
        {
            get
            {
                return email3DisplayAs;
            }
            set
            {
                email3DisplayAs = value;
            }
        }

        /// <summary>
        /// Contains the type of the first e-mail address for the contact.
        /// </summary>
        /// <value>The type of the email1.</value>
        public string Email1Type
        {
            get
            {
                return email1Type;
            }
            set
            {
                email1Type = value;
            }
        }

        /// <summary>
        /// Contains the type of the second e-mail address for the contact.
        /// </summary>
        /// <value>The type of the email2.</value>
        public string Email2Type
        {
            get
            {
                return email2Type;
            }
            set
            {
                email2Type = value;
            }
        }

        /// <summary>
        /// Contains the type of the third e-mail address for the contact.
        /// </summary>
        /// <value>The type of the email3.</value>
        public string Email3Type
        {
            get
            {
                return email3Type;
            }
            set
            {
                email3Type = value;
            }
        }

        /// <summary>
        /// Contains the entry ID of the first e-mail address for the contact.
        /// </summary>
        /// <value>The email1 entry identifier.</value>
        public byte[] Email1EntryId
        {
            get
            {
                return email1EntryId;
            }
            set
            {
                email1EntryId = value;
            }
        }

        /// <summary>
        /// Contains the entry ID of the second e-mail address for the contact.
        /// </summary>
        /// <value>The email2 entry identifier.</value>
        public byte[] Email2EntryId
        {
            get
            {
                return email2EntryId;
            }
            set
            {
                email2EntryId = value;
            }
        }

        /// <summary>
        /// Contains the entry ID of the third e-mail address for the contact.
        /// </summary>
        /// <value>The email3 entry identifier.</value>
        public byte[] Email3EntryId
        {
            get
            {
                return email3EntryId;
            }
            set
            {
                email3EntryId = value;
            }
        }

        /// <summary>
        /// Contains collection of recipients.
        /// </summary>
        /// <value>The recipients.</value>
        public IList<Recipient> Recipients
        {
            get
            {
                return recipients;
            }
        }

        /// <summary>
        /// Contains collection of attachments.
        /// </summary>
        /// <value>The attachments.</value>
        public IList<Attachment> Attachments
        {
            get
            {
                return attachments;
            }
        }

        /// <summary>
        /// Contains collection of extended (custom) properties.
        /// </summary>
        /// <value>The extended properties.</value>
        public ExtendedPropertyList ExtendedProperties
        {
            get
            {
                return extendedProperties;
            }
        }

        /// <summary>
        /// Contains collection of named properties definition. 
        /// </summary>
        internal IList<NamedProperty> NamedProperties
        {
            get
            {
                return namedProperties;
            }
        }

        /// <summary>
        /// Contains true if Message is embedded into another message object.
        /// </summary>
        /// <value><c>true</c> if this instance is embedded; otherwise, <c>false</c>.</value>
        public bool IsEmbedded
        {
            get
            {
                return isEmbedded;
            }
            set
            {
                isEmbedded = value;
            }
        }

        #endregion
    }
}
