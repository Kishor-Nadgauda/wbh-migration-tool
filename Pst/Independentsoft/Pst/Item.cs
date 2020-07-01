using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Item.
    /// </summary>
    public class Item
    {
        internal PstFileReader reader;
        internal ulong id;
        internal ulong parentId;
        internal Table table;

        internal string messageClass;
        internal string subject;
        internal string subjectPrefix;
        internal string conversationTopic;
        internal string displayBcc;
        internal string displayCc;
        internal string displayTo;
        internal string originalDisplayTo;
        internal string replyTo;
        internal string normalizedSubject;
        internal string body;
        internal byte[] bodyHtml;
        internal byte[] rtfCompressed;
        internal byte[] searchKey;
        internal byte[] entryId;
        internal DateTime creationTime;
        internal DateTime lastModificationTime;
        internal DateTime messageDeliveryTime;
        internal DateTime clientSubmitTime;
        internal DateTime providerSubmitTime;
        internal DateTime reportTime;
        internal string reportText;
        internal string creatorName;
        internal string lastModifierName;
        internal string internetMessageId;
        internal string inReplyTo;
        internal string internetReferences;
        internal int messageCodePage;
        internal int iconIndex;
        internal int messageSize;
        internal int internetCodePage;
        internal byte[] conversationIndex;
        internal bool isHidden;
        internal bool isReadOnly;
        internal bool isSystem;
        internal bool disableFullFidelity;
        internal bool hasAttachment;
        internal bool rtfInSync;
        internal Sensitivity sensitivity = Sensitivity.None;
        internal Importance importance = Importance.None;
        internal Priority priority = Priority.None;
        internal FlagIcon flagIcon = FlagIcon.None;
        internal FlagStatus flagStatus = FlagStatus.None;
        internal ObjectType objectType = ObjectType.None;
        internal string receivedRepresentingAddressType;
        internal string receivedRepresentingEmailAddress;
        internal byte[] receivedRepresentingEntryId;
        internal string receivedRepresentingName;
        internal byte[] receivedRepresentingSearchKey;
        internal string receivedByAddressType;
        internal string receivedByEmailAddress;
        internal byte[] receivedByEntryId;
        internal string receivedByName;
        internal byte[] receivedBySearchKey;
        internal string senderAddressType;
        internal string senderEmailAddress;
        internal byte[] senderEntryId;
        internal string senderName;
        internal byte[] senderSearchKey;
        internal string sentRepresentingAddressType;
        internal string sentRepresentingEmailAddress;
        internal byte[] sentRepresentingEntryId;
        internal string sentRepresentingName;
        internal byte[] sentRepresentingSearchKey;
        internal string transportMessageHeaders;
        internal DateTime lastVerbExecutionTime;
        internal LastVerbExecuted lastVerbExecuted = LastVerbExecuted.None;
        internal IList<MessageFlag> messageFlags = new List<MessageFlag>();

        //contacts properties
        internal DateTime birthday;
        internal IList<string> childrenNames = new List<string>();
        internal string assistentName;
        internal string assistentPhone;
        internal string businessPhone;
        internal string businessPhone2;
        internal string businessFax;
        internal string businessHomePage;
        internal string callbackPhone;
        internal string carPhone;
        internal string cellularPhone;
        internal string companyMainPhone;
        internal string companyName;
        internal string computerNetworkName;
        internal string customerId;
        internal string departmentName;
        internal string displayName;
        internal string displayNamePrefix;
        internal string ftpSite;
        internal string generation;
        internal string givenName;
        internal string governmentId;
        internal string hobbies;
        internal string homePhone2;
        internal string homeAddressCity;
        internal string homeAddressCountry;
        internal string homeAddressPostalCode;
        internal string homeAddressPostOfficeBox;
        internal string homeAddressState;
        internal string homeAddressStreet;
        internal string homeFax;
        internal string homePhone;
        internal string initials;
        internal string isdn;
        internal string managerName;
        internal string middleName;
        internal string nickname;
        internal string officeLocation;
        internal string otherAddressCity;
        internal string otherAddressCountry;
        internal string otherAddressPostalCode;
        internal string otherAddressState;
        internal string otherAddressStreet;
        internal string otherPhone;
        internal string pager;
        internal string personalHomePage;
        internal string postalAddress;
        internal string businessAddressCountry;
        internal string businessAddressCity;
        internal string businessAddressPostalCode;
        internal string businessAddressPostOfficeBox;
        internal string businessAddressState;
        internal string businessAddressStreet;
        internal string primaryFax;
        internal string primaryPhone;
        internal string profession;
        internal string radioPhone;
        internal string spouseName;
        internal string surname;
        internal string telex;
        internal string title;
        internal string ttyTddPhone;
        internal DateTime weddingAnniversary;
        internal Gender gender = Gender.None;

        //general named properties
        internal string outlookVersion;
        internal int outlookInternalVersion;
        internal DateTime commonStartTime;
        internal DateTime commonEndTime;
        internal DateTime flagDueBy;
        internal bool isRecurring;
        internal DateTime reminderTime;
        internal int reminderMinutesBeforeStart;
        internal IList<string> companies = new List<string>();
        internal IList<string> keywords = new List<string>();
        internal IList<string> contactNames = new List<string>();
        internal string billingInformation;
        internal string mileage;
        internal string reminderSoundFile;
        internal bool isPrivate;
        internal bool isReminderSet;
        internal bool reminderOverrideDefault;
        internal bool reminderPlaySound;
        internal Independentsoft.Msg.RecurrencePattern recurrencePattern;

        //appointments named properties
        internal DateTime appointmentStartTime;
        internal DateTime appointmentEndTime;
        internal bool isAllDayEvent;
        internal string location;
        internal BusyStatus busyStatus = BusyStatus.None;
        internal MeetingStatus meetingStatus = MeetingStatus.None;
        internal ResponseStatus responseStatus = ResponseStatus.None;
        internal RecurrenceType recurrenceType = RecurrenceType.None;
        internal string appointmentMessageClass;
        internal string timeZone;
        internal string recurrencePatternDescription;
        internal byte[] guid;
        internal int label = -1;
        internal int duration;

        //tasks named properties
        internal DateTime taskStartDate;
        internal DateTime taskDueDate;
        internal string owner;
        internal string delegator;
        internal double percentComplete;
        internal int actualWork;
        internal int totalWork;
        internal bool isTeamTask;
        internal bool isComplete;
        internal DateTime dateCompleted;
        internal TaskStatus taskStatus = TaskStatus.None;
        internal TaskOwnership taskOwnership = TaskOwnership.None;
        internal TaskDelegationState taskDelegationState = TaskDelegationState.None;

        //notes named properties
        internal int noteWidth;
        internal int noteHeight;
        internal int noteLeft;
        internal int noteTop;
        internal NoteColor noteColor = NoteColor.None;

        //journals named properties
        internal DateTime journalStartTime;
        internal DateTime journalEndTime;
        internal string journalType;
        internal string journalTypeDescription;
        internal int journalDuration;

        //contacts named properties
        internal SelectedMailingAddress selectedMailingAddress = SelectedMailingAddress.None;
        internal bool contactHasPicture;
        internal string fileAs;
        internal string instantMessengerAddress;
        internal string internetFreeBusyAddress;
        internal string businessAddress;
        internal string homeAddress;
        internal string otherAddress;
        internal string email1Address;
        internal string email2Address;
        internal string email3Address;
        internal string email1DisplayName;
        internal string email2DisplayName;
        internal string email3DisplayName;
        internal string email1DisplayAs;
        internal string email2DisplayAs;
        internal string email3DisplayAs;
        internal string email1Type;
        internal string email2Type;
        internal string email3Type;
        internal byte[] email1EntryId;
        internal byte[] email2EntryId;
        internal byte[] email3EntryId;

        internal System.Text.Encoding encoding = System.Text.Encoding.UTF8;
        internal bool isEmbedded;

        private IList<Recipient> recipients = new List<Recipient>();
        private IList<Attachment> attachments = new List<Attachment>();
        private ExtendedPropertyList extendedProperties = new ExtendedPropertyList();

        internal Item()
        {
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false, this.encoding);
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="encoding">The encoding.</param>
        public void Save(string filePath, System.Text.Encoding encoding)
        {
            Save(filePath, false, encoding);
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        public void Save(string filePath, bool overwrite)
        {
            Save(filePath, overwrite, this.encoding);
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        /// <param name="encoding">The encoding.</param>
        public void Save(string filePath, bool overwrite, System.Text.Encoding encoding)
        {
            Independentsoft.Msg.Message msgFile = CreateMessage(encoding);
            msgFile.Save(filePath, overwrite);
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public void Save(System.IO.Stream stream)
        {
            Save(stream, this.encoding);
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="encoding">The encoding.</param>
        public void Save(System.IO.Stream stream, System.Text.Encoding encoding)
        {
            Independentsoft.Msg.Message msgFile = CreateMessage(encoding);
            msgFile.Save(stream);
        }

        /// <summary>
        /// Gets the message file.
        /// </summary>
        /// <returns>Independentsoft.Msg.Message.</returns>
        public Independentsoft.Msg.Message GetMessageFile()
        {
            return GetMessageFile(this.encoding);
        }

        /// <summary>
        /// Gets the message file.
        /// </summary>
        /// <param name="encoding">The encoding.</param>
        /// <returns>Independentsoft.Msg.Message.</returns>
        public Independentsoft.Msg.Message GetMessageFile(System.Text.Encoding encoding)
        {
            return CreateMessage(encoding);
        }

        private Independentsoft.Msg.Message CreateMessage(System.Text.Encoding encoding)
        {
            Independentsoft.Msg.Message msgFile = new Independentsoft.Msg.Message();
            msgFile.Encoding = encoding;

            msgFile.MessageClass = this.messageClass;
            msgFile.Subject = this.subject;
            msgFile.SubjectPrefix = this.subjectPrefix;
            msgFile.ConversationTopic = this.conversationTopic;
            msgFile.DisplayBcc = this.displayBcc;
            msgFile.DisplayCc = this.displayCc;
            msgFile.DisplayTo = this.displayTo;
            msgFile.OriginalDisplayTo = this.originalDisplayTo;
            msgFile.ReplyTo = this.replyTo;
            msgFile.NormalizedSubject = this.normalizedSubject;
            msgFile.Body = this.body;
            msgFile.BodyHtml = this.bodyHtml;
            msgFile.RtfCompressed = this.rtfCompressed;
            msgFile.SearchKey = this.searchKey;
            msgFile.EntryId = this.entryId;
            msgFile.CreationTime = this.creationTime;
            msgFile.LastModificationTime = this.lastModificationTime;
            msgFile.MessageDeliveryTime = this.messageDeliveryTime;
            msgFile.ClientSubmitTime = this.clientSubmitTime;
            msgFile.ProviderSubmitTime = this.providerSubmitTime;
            msgFile.ReportTime = this.reportTime;
            msgFile.ReportText = this.reportText;
            msgFile.CreatorName = this.creatorName;
            msgFile.LastModifierName = this.lastModifierName;
            msgFile.InternetMessageId = this.internetMessageId;
            msgFile.InReplyTo = this.inReplyTo;
            msgFile.InternetReferences = this.internetReferences;
            msgFile.MessageCodePage = this.messageCodePage;
            msgFile.IconIndex = this.iconIndex;
            msgFile.Size = this.messageSize;
            msgFile.InternetCodePage = this.internetCodePage;
            msgFile.ConversationIndex = this.conversationIndex;
            msgFile.IsHidden = this.isHidden;
            msgFile.IsReadOnly = this.isReadOnly;
            msgFile.IsSystem = this.isSystem;
            msgFile.DisableFullFidelity = this.disableFullFidelity;
            msgFile.HasAttachment = this.hasAttachment;
            msgFile.RtfInSync = this.rtfInSync;
            msgFile.Sensitivity = Independentsoft.Msg.EnumUtil.ParseSensitivity(EnumUtil.ParseSensitivity(this.sensitivity));
            msgFile.Importance = Independentsoft.Msg.EnumUtil.ParseImportance(EnumUtil.ParseImportance(this.importance));
            msgFile.Priority = Independentsoft.Msg.EnumUtil.ParsePriority(EnumUtil.ParsePriority(this.priority));
            msgFile.FlagIcon = Independentsoft.Msg.EnumUtil.ParseFlagIcon(EnumUtil.ParseFlagIcon(this.flagIcon));
            msgFile.FlagStatus = Independentsoft.Msg.EnumUtil.ParseFlagStatus(EnumUtil.ParseFlagStatus(this.flagStatus));
            msgFile.ObjectType = Independentsoft.Msg.EnumUtil.ParseObjectType(EnumUtil.ParseObjectType(this.objectType));
            msgFile.ReceivedRepresentingAddressType = this.receivedRepresentingAddressType;
            msgFile.ReceivedRepresentingEmailAddress = this.receivedRepresentingEmailAddress;
            msgFile.ReceivedRepresentingEntryId = this.receivedRepresentingEntryId;
            msgFile.ReceivedRepresentingName = this.receivedRepresentingName;
            msgFile.ReceivedRepresentingSearchKey = this.receivedRepresentingSearchKey;
            msgFile.ReceivedByAddressType = this.receivedByAddressType;
            msgFile.ReceivedByEmailAddress = this.receivedByEmailAddress;
            msgFile.ReceivedByEntryId = this.receivedByEntryId;
            msgFile.ReceivedByName = this.receivedByName;
            msgFile.ReceivedBySearchKey = this.receivedBySearchKey;
            msgFile.SenderAddressType = this.senderAddressType;
            msgFile.SenderEmailAddress = this.senderEmailAddress;
            msgFile.SenderEntryId = this.senderEntryId;
            msgFile.SenderName = this.senderName;
            msgFile.SenderSearchKey = this.senderSearchKey;
            msgFile.SentRepresentingAddressType = this.sentRepresentingAddressType;
            msgFile.SentRepresentingEmailAddress = this.sentRepresentingEmailAddress;
            msgFile.SentRepresentingEntryId = this.sentRepresentingEntryId;
            msgFile.SentRepresentingName = this.sentRepresentingName;
            msgFile.SentRepresentingSearchKey = this.sentRepresentingSearchKey;
            msgFile.TransportMessageHeaders = this.transportMessageHeaders;
            msgFile.LastVerbExecutionTime = this.lastVerbExecutionTime;
            msgFile.LastVerbExecuted = Independentsoft.Msg.EnumUtil.ParseLastVerbExecuted(EnumUtil.ParseLastVerbExecuted(this.lastVerbExecuted));

            foreach (Independentsoft.Msg.MessageFlag flag in EnumUtil.ConvertMessageFlags(this.messageFlags))
            {
                msgFile.MessageFlags.Add(flag);
            }

            msgFile.Birthday = this.birthday;
            
            foreach (string childName in this.childrenNames)
            {
                msgFile.ChildrenNames.Add(childName);
            }

            msgFile.AssistentName = this.assistentName;
            msgFile.AssistentPhone = this.assistentPhone;
            msgFile.BusinessPhone = this.businessPhone;
            msgFile.BusinessPhone2 = this.businessPhone2;
            msgFile.BusinessFax = this.businessFax;
            msgFile.BusinessHomePage = this.businessHomePage;
            msgFile.CallbackPhone = this.callbackPhone;
            msgFile.CarPhone = this.carPhone;
            msgFile.CellularPhone = this.cellularPhone;
            msgFile.CompanyMainPhone = this.companyMainPhone;
            msgFile.CompanyName = this.companyName;
            msgFile.ComputerNetworkName = this.computerNetworkName;
            msgFile.CustomerId = this.customerId;
            msgFile.DepartmentName = this.departmentName;
            msgFile.DisplayName = this.displayName;
            msgFile.DisplayNamePrefix = this.displayNamePrefix;
            msgFile.FtpSite = this.ftpSite;
            msgFile.Generation = this.generation;
            msgFile.GivenName = this.givenName;
            msgFile.GovernmentId = this.governmentId;
            msgFile.Hobbies = this.hobbies;
            msgFile.HomePhone2 = this.homePhone2;
            msgFile.HomeAddressCity = this.homeAddressCity;
            msgFile.HomeAddressCountry = this.homeAddressCountry;
            msgFile.HomeAddressPostalCode = this.homeAddressPostalCode;
            msgFile.HomeAddressPostOfficeBox = this.homeAddressPostOfficeBox;
            msgFile.HomeAddressState = this.homeAddressState;
            msgFile.HomeAddressStreet = this.homeAddressStreet;
            msgFile.HomeFax = this.homeFax;
            msgFile.HomePhone = this.homePhone;
            msgFile.Initials = this.initials;
            msgFile.Isdn = this.isdn;
            msgFile.ManagerName = this.managerName;
            msgFile.MiddleName = this.middleName;
            msgFile.Nickname = this.nickname;
            msgFile.OfficeLocation = this.officeLocation;
            msgFile.OtherAddressCity = this.otherAddressCity;
            msgFile.OtherAddressCountry = this.otherAddressCountry;
            msgFile.OtherAddressPostalCode = this.otherAddressPostalCode;
            msgFile.OtherAddressState = this.otherAddressState;
            msgFile.OtherAddressStreet = this.otherAddressStreet;
            msgFile.OtherPhone = this.otherPhone;
            msgFile.Pager = this.pager;
            msgFile.PersonalHomePage = this.personalHomePage;
            msgFile.PostalAddress = this.postalAddress;
            msgFile.BusinessAddressCountry = this.businessAddressCountry;
            msgFile.BusinessAddressCity = this.businessAddressCity;
            msgFile.BusinessAddressPostalCode = this.businessAddressPostalCode;
            msgFile.BusinessAddressPostOfficeBox = this.businessAddressPostOfficeBox;
            msgFile.BusinessAddressState = this.businessAddressState;
            msgFile.BusinessAddressStreet = this.businessAddressStreet;
            msgFile.PrimaryFax = this.primaryFax;
            msgFile.PrimaryPhone = this.primaryPhone;
            msgFile.Profession = this.profession;
            msgFile.RadioPhone = this.radioPhone;
            msgFile.SpouseName = this.spouseName;
            msgFile.Surname = this.surname;
            msgFile.Telex = this.telex;
            msgFile.Title = this.title;
            msgFile.TtyTddPhone = this.ttyTddPhone;
            msgFile.WeddingAnniversary = this.weddingAnniversary;
            msgFile.OutlookVersion = this.outlookVersion;
            msgFile.OutlookInternalVersion = this.outlookInternalVersion;
            msgFile.CommonStartTime = this.commonStartTime;
            msgFile.CommonEndTime = this.commonEndTime;
            msgFile.FlagDueBy = this.flagDueBy;
            msgFile.IsRecurring = this.isRecurring;
            msgFile.ReminderTime = this.reminderTime;
            msgFile.ReminderMinutesBeforeStart = this.reminderMinutesBeforeStart;

            foreach (string company in this.companies)
            {
                msgFile.Companies.Add(company);
            }

            foreach (string contactName in this.contactNames)
            {
                msgFile.ContactNames.Add(contactName);
            }

            foreach (string keyword in this.keywords)
            {
                msgFile.Keywords.Add(keyword);
            }

            msgFile.BillingInformation = this.billingInformation;
            msgFile.Mileage = this.mileage;
            msgFile.ReminderSoundFile = this.reminderSoundFile;
            msgFile.IsPrivate = this.isPrivate;
            msgFile.IsReminderSet = this.isReminderSet;
            msgFile.ReminderOverrideDefault = this.reminderOverrideDefault;
            msgFile.ReminderPlaySound = this.reminderPlaySound;
            msgFile.AppointmentStartTime = this.appointmentStartTime;
            msgFile.AppointmentEndTime = this.appointmentEndTime;
            msgFile.IsAllDayEvent = this.isAllDayEvent;
            msgFile.Location = this.location;
            msgFile.BusyStatus = Independentsoft.Msg.EnumUtil.ParseBusyStatus(EnumUtil.ParseBusyStatus(this.busyStatus));
            msgFile.MeetingStatus = Independentsoft.Msg.EnumUtil.ParseMeetingStatus(EnumUtil.ParseMeetingStatus(this.meetingStatus));
            msgFile.ResponseStatus = Independentsoft.Msg.EnumUtil.ParseResponseStatus(EnumUtil.ParseResponseStatus(this.responseStatus));
            msgFile.RecurrenceType = Independentsoft.Msg.EnumUtil.ParseRecurrenceType(EnumUtil.ParseRecurrenceType(this.recurrenceType));
            msgFile.Gender = Independentsoft.Msg.EnumUtil.ParseGender(EnumUtil.ParseGender(this.gender));
            msgFile.AppointmentMessageClass = this.appointmentMessageClass;
            msgFile.TimeZone = this.timeZone;
            msgFile.RecurrencePatternDescription = this.recurrencePatternDescription;
            msgFile.recurrencePattern = this.recurrencePattern;
            msgFile.Guid = this.guid;
            msgFile.Label = this.label;
            msgFile.Duration = this.duration;
            msgFile.TaskStartDate = this.taskStartDate;
            msgFile.TaskDueDate = this.taskDueDate;
            msgFile.Owner = this.owner;
            msgFile.Delegator = this.delegator;
            msgFile.PercentComplete = this.percentComplete;
            msgFile.ActualWork = this.actualWork;
            msgFile.TotalWork = this.totalWork;
            msgFile.IsTeamTask = this.isTeamTask;
            msgFile.IsComplete = this.isComplete;
            msgFile.DateCompleted = this.dateCompleted;
            msgFile.TaskStatus = Independentsoft.Msg.EnumUtil.ParseTaskStatus(EnumUtil.ParseTaskStatus(this.taskStatus));
            msgFile.TaskOwnership = Independentsoft.Msg.EnumUtil.ParseTaskOwnership(EnumUtil.ParseTaskOwnership(this.taskOwnership));
            msgFile.TaskDelegationState = Independentsoft.Msg.EnumUtil.ParseTaskDelegationState(EnumUtil.ParseTaskDelegationState(this.taskDelegationState));
            msgFile.NoteWidth = this.noteWidth;
            msgFile.NoteHeight = this.noteHeight;
            msgFile.NoteLeft = this.noteLeft;
            msgFile.NoteTop = this.noteTop;
            msgFile.NoteColor = Independentsoft.Msg.EnumUtil.ParseNoteColor(EnumUtil.ParseNoteColor(this.noteColor));
            msgFile.JournalStartTime = this.journalStartTime;
            msgFile.JournalEndTime = this.journalEndTime;
            msgFile.JournalType = this.journalType;
            msgFile.JournalTypeDescription = this.journalTypeDescription;
            msgFile.JournalDuration = this.journalDuration;
            msgFile.SelectedMailingAddress = Independentsoft.Msg.EnumUtil.ParseSelectedMailingAddress(EnumUtil.ParseSelectedMailingAddress(this.selectedMailingAddress));
            msgFile.ContactHasPicture = this.contactHasPicture;
            msgFile.FileAs = this.fileAs;
            msgFile.InstantMessengerAddress = this.instantMessengerAddress;
            msgFile.InternetFreeBusyAddress = this.internetFreeBusyAddress;
            msgFile.BusinessAddress = this.businessAddress;
            msgFile.HomeAddress = this.homeAddress;
            msgFile.OtherAddress = this.otherAddress;
            msgFile.Email1Address = this.email1Address;
            msgFile.Email2Address = this.email2Address;
            msgFile.Email3Address = this.email3Address;
            msgFile.Email1DisplayName = this.email1DisplayName;
            msgFile.Email2DisplayName = this.email2DisplayName;
            msgFile.Email3DisplayName = this.email3DisplayName;
            msgFile.Email1DisplayAs = this.email1DisplayAs;
            msgFile.Email2DisplayAs = this.email2DisplayAs;
            msgFile.Email3DisplayAs = this.email3DisplayAs;
            msgFile.Email1Type = this.email1Type;
            msgFile.Email2Type = this.email2Type;
            msgFile.Email3Type = this.email3Type;
            msgFile.Email1EntryId = this.email1EntryId;
            msgFile.Email2EntryId = this.email2EntryId;
            msgFile.Email3EntryId = this.email3EntryId;

            for (int e = 0; e < this.ExtendedProperties.Count; e++)
            {
                Independentsoft.Pst.ExtendedProperty pstExtendedProperty = this.ExtendedProperties[e];

                if (pstExtendedProperty.Tag != null && pstExtendedProperty.Tag.Guid != null)
                {
                    Independentsoft.Msg.ExtendedProperty msgExtendedProperty = new Independentsoft.Msg.ExtendedProperty();

                    if (pstExtendedProperty.Tag is Independentsoft.Pst.ExtendedPropertyId)
                    {
                        Independentsoft.Pst.ExtendedPropertyId pstExtendedPropertyId = (Independentsoft.Pst.ExtendedPropertyId)pstExtendedProperty.Tag;

                        Independentsoft.Msg.ExtendedPropertyId msgExtendedPropertyId = new Independentsoft.Msg.ExtendedPropertyId();

                        msgExtendedPropertyId.Id = pstExtendedPropertyId.Id;
                        msgExtendedPropertyId.Guid = pstExtendedPropertyId.Guid;

                        msgExtendedProperty.Tag = msgExtendedPropertyId;
                    }
                    else
                    {
                        Independentsoft.Pst.ExtendedPropertyName pstExtendedPropertyName = (Independentsoft.Pst.ExtendedPropertyName)pstExtendedProperty.Tag;

                        Independentsoft.Msg.ExtendedPropertyName msgExtendedPropertyName = new Independentsoft.Msg.ExtendedPropertyName();

                        msgExtendedPropertyName.Name = pstExtendedPropertyName.Name;
                        msgExtendedPropertyName.Guid = pstExtendedPropertyName.Guid;

                        msgExtendedProperty.Tag = msgExtendedPropertyName;
                    }

                    if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.ApplicationTime)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Time;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.ApplicationTimeArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleTime;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Binary)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Binary;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.BinaryArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleBinary;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Boolean)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Boolean;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Currency)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Currency;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.CurrencyArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleCurrency;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Double)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Floating64;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.DoubleArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleFloating64;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Error)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.ErrorCode;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Float)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Floating32;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.FloatArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleFloating32;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Guid)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Guid;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.GuidArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleGuid;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Integer)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Integer32;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.IntegerArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleInteger32;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Long)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Integer64;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.LongArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleInteger64;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Object)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Object;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.Short)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Integer16;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.ShortArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleInteger16;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.String)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.String;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.StringArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleString;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.String8)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.String8;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.String8Array)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleString8;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.SystemTime)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Time;
                    }
                    else if (pstExtendedProperty.Tag.Type == Independentsoft.Pst.PropertyType.SystemTimeArray)
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.MultipleTime;
                    }
                    else
                    {
                        msgExtendedProperty.Tag.Type = Independentsoft.Msg.PropertyType.Binary;
                    }

                    msgExtendedProperty.Value = pstExtendedProperty.Value;

                    msgFile.ExtendedProperties.Add(msgExtendedProperty);
                }
            }

            for (int r = 0; r < this.Recipients.Count; r++)
            {
                Independentsoft.Msg.Recipient recipient = new Independentsoft.Msg.Recipient();

                recipient.AddressType = this.Recipients[r].AddressType;
                recipient.DisplayName = this.Recipients[r].DisplayName;
                recipient.DisplayType = Independentsoft.Msg.EnumUtil.ParseDisplayType(EnumUtil.ParseDisplayType(this.Recipients[r].DisplayType));
                recipient.EmailAddress = this.Recipients[r].EmailAddress;
                recipient.EntryId = this.Recipients[r].EntryId;
                recipient.InstanceKey = this.Recipients[r].InstanceKey;
                recipient.ObjectType = Independentsoft.Msg.EnumUtil.ParseObjectType(EnumUtil.ParseObjectType(this.Recipients[r].ObjectType));
                recipient.RecipientType = Independentsoft.Msg.EnumUtil.ParseRecipientType(EnumUtil.ParseRecipientType(this.Recipients[r].RecipientType));
                recipient.SearchKey = this.Recipients[r].SearchKey;
                recipient.Responsibility = this.Recipients[r].Responsibility;
                recipient.SmtpAddress = this.Recipients[r].SmtpAddress;
                recipient.DisplayName7Bit = this.Recipients[r].DisplayName7Bit;
                recipient.TransmitableDisplayName = this.Recipients[r].TransmitableDisplayName;
                recipient.SendRichInfo = this.Recipients[r].SendRichInfo;
                recipient.SendInternetEncoding = this.Recipients[r].SendInternetEncoding;
                recipient.OriginatingAddressType = this.Recipients[r].OriginatingAddressType;
                recipient.OriginatingEmailAddress = this.Recipients[r].OriginatingEmailAddress;

                msgFile.Recipients.Add(recipient);
            }

            for (int r = 0; r < this.Attachments.Count; r++)
            {
                Independentsoft.Msg.Attachment attachment = new Independentsoft.Msg.Attachment();

                attachment.AdditionalInfo = this.Attachments[r].AdditionalInfo;
                attachment.ContentBase = this.Attachments[r].ContentBase;
                attachment.ContentDisposition = this.Attachments[r].ContentDisposition;
                attachment.ContentId = this.Attachments[r].ContentId;
                attachment.ContentLocation = this.Attachments[r].ContentLocation;
                attachment.Data = this.Attachments[r].Data;
                attachment.DataObject = this.Attachments[r].DataObject;
                attachment.DisplayName = this.Attachments[r].DisplayName;
                attachment.Encoding = this.Attachments[r].Encoding;
                attachment.Extension = this.Attachments[r].Extension;
                attachment.FileName = this.Attachments[r].FileName;
                attachment.Flags = Independentsoft.Msg.EnumUtil.ParseAttachmentFlags(EnumUtil.ParseAttachmentFlags(this.Attachments[r].Flags));
                attachment.IsContactPhoto = this.Attachments[r].IsContactPhoto;
                attachment.IsHidden = this.Attachments[r].IsHidden;
                attachment.LongFileName = this.Attachments[r].LongFileName;
                attachment.LongPathName = this.Attachments[r].LongPathName;
                attachment.Method = Independentsoft.Msg.EnumUtil.ParseAttachmentMethod(EnumUtil.ParseAttachmentMethod(this.Attachments[r].Method));
                attachment.MimeSequence = this.Attachments[r].MimeSequence;
                attachment.MimeTag = this.Attachments[r].MimeTag;
                attachment.ObjectType = Independentsoft.Msg.EnumUtil.ParseObjectType(EnumUtil.ParseObjectType(this.Attachments[r].ObjectType));
                attachment.PathName = this.Attachments[r].PathName;
                attachment.Rendering = this.Attachments[r].Rendering;
                attachment.RenderingPosition = this.Attachments[r].RenderingPosition;
                attachment.Size = this.Attachments[r].Size;
                attachment.Tag = this.Attachments[r].Tag;
                attachment.TransportName = this.Attachments[r].TransportName;
                attachment.CreationTime = this.Attachments[r].CreationTime;
                attachment.LastModificationTime = this.Attachments[r].LastModificationTime;

                if (this.Attachments[r].EmbeddedItem != null)
                {
                    attachment.EmbeddedMessage = this.Attachments[r].EmbeddedItem.CreateMessage(encoding);
                }

                msgFile.Attachments.Add(attachment);
            }

            return msgFile;
        }

        #region Properties

        /// <summary>
        /// Gets the identifier.
        /// </summary>
        /// <value>The identifier.</value>
        public long Id
        {
            get
            {
                return (long)id;
            }
        }

        /// <summary>
        /// Gets the parent identifier.
        /// </summary>
        /// <value>The parent identifier.</value>
        public long ParentId
        {
            get
            {
                return (long)parentId;
            }
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <value>The table.</value>
        public Table Table
        {
            get
            {
                return table;
            }
        }

        /// <summary>
        /// Gets or sets message class, such as "IPM.Note". The message class specifies the type of the message. It determines the set of properties defined for the message, the kind of information the message conveys, and how to handle the message. This property represents MAPI property <a href="http://msdn.microsoft.com/en-us/library/ms531519(EXCHG.10).aspx">PR_MESSAGE_CLASS</a>.
        /// Here is a list of standard values <a href="http://msdn.microsoft.com/en-us/library/aa204771.aspx" />.
        /// </summary>
        /// <value>The message class.</value>
        public string MessageClass
        {
            get
            {
                return messageClass;
            }
        }

        /// <summary>
        /// Gets or set subject. This property represents MAPI property <a href="http://msdn.microsoft.com/en-us/library/ms527932(EXCHG.10).aspx">PR_SUBJECT</a>.
        /// </summary>
        /// <value>The subject.</value>
        public string Subject
        {
            get
            {
                return subject;
            }
        }

        /// <summary>
        /// Gets or sets subject prefix that typically indicates some action on a message, such as "FW: " for forwarding. This property represents MAPI property <a href="http://msdn.microsoft.com/en-us/library/ms526750(EXCHG.10).aspx">PR_SUBJECT_PREFIX</a>.
        /// </summary>
        /// <value>The subject prefix.</value>
        public string SubjectPrefix
        {
            get
            {
                return subjectPrefix;
            }
        }

        /// <summary>
        /// Gets the normalized subject.
        /// </summary>
        /// <value>The normalized subject.</value>
        public string NormalizedSubject
        {
            get
            {
                return normalizedSubject;
            }
        }

        /// <summary>
        /// Gets the body.
        /// </summary>
        /// <value>The body.</value>
        public string Body
        {
            get
            {
                return body;
            }
        }

        /// <summary>
        /// Gets the body HTML.
        /// </summary>
        /// <value>The body HTML.</value>
        public byte[] BodyHtml
        {
            get
            {
                return bodyHtml;
            }
        }

        /// <summary>
        /// Gets the body HTML text.
        /// </summary>
        /// <value>The body HTML text.</value>
        public string BodyHtmlText
        {
            get
            {
                if (bodyHtml != null)
                {
                    if (Independentsoft.Msg.Utf8Util.IsUtf8(bodyHtml, bodyHtml.Length))
                    {
                        return System.Text.Encoding.UTF8.GetString(bodyHtml, 0, bodyHtml.Length);
                    }
                    else
                    {
                        return System.Text.Encoding.UTF7.GetString(bodyHtml, 0, bodyHtml.Length);
                    }
                }
                else
                {
                    byte[] bodyRtf = BodyRtf;

                    if (bodyRtf != null && bodyRtf.Length > 0 && bodyRtf.Length < 16777216) //convert large RTF body is slow
                    {
                        Independentsoft.Msg.HtmlText htmlText = Independentsoft.Msg.Util.ConvertRtfToHtml(bodyRtf);

                        if (htmlText != null)
                        {
                            return htmlText.Text;
                        }
                    }

                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the body RTF.
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
        }

        /// <summary>
        /// Gets the RTF compressed.
        /// </summary>
        /// <value>The RTF compressed.</value>
        public byte[] RtfCompressed
        {
            get
            {
                return rtfCompressed;
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
        /// Gets the creation time.
        /// </summary>
        /// <value>The creation time.</value>
        public DateTime CreationTime
        {
            get
            {
                return creationTime;
            }
        }

        /// <summary>
        /// Gets the last modification time.
        /// </summary>
        /// <value>The last modification time.</value>
        public DateTime LastModificationTime
        {
            get
            {
                return lastModificationTime;
            }
        }

        /// <summary>
        /// Gets the report text.
        /// </summary>
        /// <value>The report text.</value>
        public string ReportText
        {
            get
            {
                return reportText;
            }
        }

        /// <summary>
        /// Gets the name of the creator.
        /// </summary>
        /// <value>The name of the creator.</value>
        public string CreatorName
        {
            get
            {
                return creatorName;
            }
        }

        /// <summary>
        /// Gets the last name of the modifier.
        /// </summary>
        /// <value>The last name of the modifier.</value>
        public string LastModifierName
        {
            get
            {
                return lastModifierName;
            }
        }

        /// <summary>
        /// Gets the index of the icon.
        /// </summary>
        /// <value>The index of the icon.</value>
        public int IconIndex
        {
            get
            {
                return iconIndex;
            }
        }

        /// <summary>
        /// Gets the size.
        /// </summary>
        /// <value>The size.</value>
        public int Size
        {
            get
            {
                return messageSize;
            }
        }

        /// <summary>
        /// Gets the internet code page.
        /// </summary>
        /// <value>The internet code page.</value>
        public int InternetCodePage
        {
            get
            {
                return internetCodePage;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is hidden.
        /// </summary>
        /// <value><c>true</c> if this instance is hidden; otherwise, <c>false</c>.</value>
        public bool IsHidden
        {
            get
            {
                return isHidden;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is read only.
        /// </summary>
        /// <value><c>true</c> if this instance is read only; otherwise, <c>false</c>.</value>
        public bool IsReadOnly
        {
            get
            {
                return isReadOnly;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is system.
        /// </summary>
        /// <value><c>true</c> if this instance is system; otherwise, <c>false</c>.</value>
        public bool IsSystem
        {
            get
            {
                return isSystem;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [disable full fidelity].
        /// </summary>
        /// <value><c>true</c> if [disable full fidelity]; otherwise, <c>false</c>.</value>
        public bool DisableFullFidelity
        {
            get
            {
                return disableFullFidelity;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has attachment.
        /// </summary>
        /// <value><c>true</c> if this instance has attachment; otherwise, <c>false</c>.</value>
        public bool HasAttachment
        {
            get
            {
                return hasAttachment;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [RTF in synchronize].
        /// </summary>
        /// <value><c>true</c> if [RTF in synchronize]; otherwise, <c>false</c>.</value>
        public bool RtfInSync
        {
            get
            {
                return rtfInSync;
            }
        }

        /// <summary>
        /// Gets the sensitivity.
        /// </summary>
        /// <value>The sensitivity.</value>
        public Sensitivity Sensitivity
        {
            get
            {
                return sensitivity;
            }
        }

        /// <summary>
        /// Gets the flag status.
        /// </summary>
        /// <value>The flag status.</value>
        public FlagStatus FlagStatus
        {
            get
            {
                return flagStatus;
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
        /// Gets the message flags.
        /// </summary>
        /// <value>The message flags.</value>
        public MessageFlag[] MessageFlags
        {
            get
            {
                MessageFlag[] messageFlagsArray = new MessageFlag[messageFlags.Count];

                for (int i = 0; i < messageFlagsArray.Length; i++)
                {
                    messageFlagsArray[i] = (MessageFlag)messageFlags[i];
                }

                return messageFlagsArray;
            }
        }

        /// <summary>
        /// Gets the outlook version.
        /// </summary>
        /// <value>The outlook version.</value>
        public string OutlookVersion
        {
            get
            {
                return outlookVersion;
            }
        }

        /// <summary>
        /// Gets the outlook internal version.
        /// </summary>
        /// <value>The outlook internal version.</value>
        public long OutlookInternalVersion
        {
            get
            {
                return outlookInternalVersion;
            }
        }

        /// <summary>
        /// Gets the keywords.
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
        /// Gets the categories.
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
        /// Gets the billing information.
        /// </summary>
        /// <value>The billing information.</value>
        public string BillingInformation
        {
            get
            {
                return billingInformation;
            }
        }

        /// <summary>
        /// Gets the mileage.
        /// </summary>
        /// <value>The mileage.</value>
        public string Mileage
        {
            get
            {
                return mileage;
            }
        }

        /// <summary>
        /// Gets the encoding.
        /// </summary>
        /// <value>The encoding.</value>
        public System.Text.Encoding Encoding
        {
            get
            {
                return encoding;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is embedded.
        /// </summary>
        /// <value><c>true</c> if this instance is embedded; otherwise, <c>false</c>.</value>
        public bool IsEmbedded
        {
            get
            {
                return isEmbedded;
            }
        }

        /// <summary>
        /// Gets the recipients.
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
        /// Gets the attachments.
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
        /// Gets the extended properties.
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
        /// Gets the type of the received representing address.
        /// </summary>
        /// <value>The type of the received representing address.</value>
        public string ReceivedRepresentingAddressType
        {
            get
            {
                return receivedRepresentingAddressType;
            }
        }

        /// <summary>
        /// Gets the received representing email address.
        /// </summary>
        /// <value>The received representing email address.</value>
        public string ReceivedRepresentingEmailAddress
        {
            get
            {
                return receivedRepresentingEmailAddress;
            }
        }

        /// <summary>
        /// Gets the received representing entry identifier.
        /// </summary>
        /// <value>The received representing entry identifier.</value>
        public byte[] ReceivedRepresentingEntryId
        {
            get
            {
                return receivedRepresentingEntryId;
            }
        }

        /// <summary>
        /// Gets the name of the received representing.
        /// </summary>
        /// <value>The name of the received representing.</value>
        public string ReceivedRepresentingName
        {
            get
            {
                return receivedRepresentingName;
            }
        }

        /// <summary>
        /// Gets the received representing search key.
        /// </summary>
        /// <value>The received representing search key.</value>
        public byte[] ReceivedRepresentingSearchKey
        {
            get
            {
                return receivedRepresentingSearchKey;
            }
        }

        /// <summary>
        /// Gets the type of the received by address.
        /// </summary>
        /// <value>The type of the received by address.</value>
        public string ReceivedByAddressType
        {
            get
            {
                return receivedByAddressType;
            }
        }

        /// <summary>
        /// Gets the received by email address.
        /// </summary>
        /// <value>The received by email address.</value>
        public string ReceivedByEmailAddress
        {
            get
            {
                return receivedByEmailAddress;
            }
        }

        /// <summary>
        /// Gets the received by entry identifier.
        /// </summary>
        /// <value>The received by entry identifier.</value>
        public byte[] ReceivedByEntryId
        {
            get
            {
                return receivedByEntryId;
            }
        }

        /// <summary>
        /// Gets the name of the received by.
        /// </summary>
        /// <value>The name of the received by.</value>
        public string ReceivedByName
        {
            get
            {
                return receivedByName;
            }
        }

        /// <summary>
        /// Gets the received by search key.
        /// </summary>
        /// <value>The received by search key.</value>
        public byte[] ReceivedBySearchKey
        {
            get
            {
                return receivedBySearchKey;
            }
        }

        /// <summary>
        /// Gets the type of the sender address.
        /// </summary>
        /// <value>The type of the sender address.</value>
        public string SenderAddressType
        {
            get
            {
                return senderAddressType;
            }
        }

        /// <summary>
        /// Gets the sender email address.
        /// </summary>
        /// <value>The sender email address.</value>
        public string SenderEmailAddress
        {
            get
            {
                return senderEmailAddress;
            }
        }

        /// <summary>
        /// Gets the sender entry identifier.
        /// </summary>
        /// <value>The sender entry identifier.</value>
        public byte[] SenderEntryId
        {
            get
            {
                return senderEntryId;
            }
        }

        /// <summary>
        /// Gets the name of the sender.
        /// </summary>
        /// <value>The name of the sender.</value>
        public string SenderName
        {
            get
            {
                return senderName;
            }
        }

        /// <summary>
        /// Gets the sender search key.
        /// </summary>
        /// <value>The sender search key.</value>
        public byte[] SenderSearchKey
        {
            get
            {
                return senderSearchKey;
            }
        }

        /// <summary>
        /// Gets the type of the sent representing address.
        /// </summary>
        /// <value>The type of the sent representing address.</value>
        public string SentRepresentingAddressType
        {
            get
            {
                return sentRepresentingAddressType;
            }
        }

        /// <summary>
        /// Gets the sent representing email address.
        /// </summary>
        /// <value>The sent representing email address.</value>
        public string SentRepresentingEmailAddress
        {
            get
            {
                return sentRepresentingEmailAddress;
            }
        }

        /// <summary>
        /// Gets the sent representing entry identifier.
        /// </summary>
        /// <value>The sent representing entry identifier.</value>
        public byte[] SentRepresentingEntryId
        {
            get
            {
                return sentRepresentingEntryId;
            }
        }

        /// <summary>
        /// Gets the name of the sent representing.
        /// </summary>
        /// <value>The name of the sent representing.</value>
        public string SentRepresentingName
        {
            get
            {
                return sentRepresentingName;
            }
        }

        /// <summary>
        /// Gets the sent representing search key.
        /// </summary>
        /// <value>The sent representing search key.</value>
        public byte[] SentRepresentingSearchKey
        {
            get
            {
                return sentRepresentingSearchKey;
            }
        }

        /// <summary>
        /// Gets the transport message headers.
        /// </summary>
        /// <value>The transport message headers.</value>
        public string TransportMessageHeaders
        {
            get
            {
                return transportMessageHeaders;
            }
        }

        /// <summary>
        /// Gets the last verb executed.
        /// </summary>
        /// <value>The last verb executed.</value>
        public LastVerbExecuted LastVerbExecuted
        {
            get
            {
                return lastVerbExecuted;
            }
        }

        /// <summary>
        /// Gets the importance.
        /// </summary>
        /// <value>The importance.</value>
        public Importance Importance
        {
            get
            {
                return importance;
            }
        }

        /// <summary>
        /// Gets the priority.
        /// </summary>
        /// <value>The priority.</value>
        public Priority Priority
        {
            get
            {
                return priority;
            }
        }

        /// <summary>
        /// Gets the flag icon.
        /// </summary>
        /// <value>The flag icon.</value>
        public FlagIcon FlagIcon
        {
            get
            {
                return flagIcon;
            }
        }

        /// <summary>
        /// Gets the conversation topic.
        /// </summary>
        /// <value>The conversation topic.</value>
        public string ConversationTopic
        {
            get
            {
                return conversationTopic;
            }
        }

        /// <summary>
        /// Gets the display BCC.
        /// </summary>
        /// <value>The display BCC.</value>
        public string DisplayBcc
        {
            get
            {
                return displayBcc;
            }
        }

        /// <summary>
        /// Gets the display cc.
        /// </summary>
        /// <value>The display cc.</value>
        public string DisplayCc
        {
            get
            {
                return displayCc;
            }
        }

        /// <summary>
        /// Gets the display to.
        /// </summary>
        /// <value>The display to.</value>
        public string DisplayTo
        {
            get
            {
                return displayTo;
            }
        }

        /// <summary>
        /// Gets the original display to.
        /// </summary>
        /// <value>The original display to.</value>
        public string OriginalDisplayTo
        {
            get
            {
                return originalDisplayTo;
            }
        }

        /// <summary>
        /// Gets the reply to.
        /// </summary>
        /// <value>The reply to.</value>
        public string ReplyTo
        {
            get
            {
                return replyTo;
            }
        }

        /// <summary>
        /// Gets the index of the conversation.
        /// </summary>
        /// <value>The index of the conversation.</value>
        public byte[] ConversationIndex
        {
            get
            {
                return conversationIndex;
            }
        }

        /// <summary>
        /// Gets the internet message identifier.
        /// </summary>
        /// <value>The internet message identifier.</value>
        public string InternetMessageId
        {
            get
            {
                return internetMessageId;
            }
        }

        /// <summary>
        /// Gets the in reply to.
        /// </summary>
        /// <value>The in reply to.</value>
        public string InReplyTo
        {
            get
            {
                return inReplyTo;
            }
        }

        /// <summary>
        /// Gets the internet references.
        /// </summary>
        /// <value>The internet references.</value>
        public string InternetReferences
        {
            get
            {
                return internetReferences;
            }
        }

        /// <summary>
        /// Gets the code page.
        /// </summary>
        /// <value>The code page.</value>
        public int CodePage
        {
            get
            {
                return messageCodePage;
            }
        }


        /// <summary>
        /// Gets the message delivery time.
        /// </summary>
        /// <value>The message delivery time.</value>
        public DateTime MessageDeliveryTime
        {
            get
            {
                return messageDeliveryTime;
            }
        }

        /// <summary>
        /// Gets the client submit time.
        /// </summary>
        /// <value>The client submit time.</value>
        public DateTime ClientSubmitTime
        {
            get
            {
                return clientSubmitTime;
            }
        }

        /// <summary>
        /// Gets the provider submit time.
        /// </summary>
        /// <value>The provider submit time.</value>
        public DateTime ProviderSubmitTime
        {
            get
            {
                return providerSubmitTime;
            }
        }

        /// <summary>
        /// Gets the report time.
        /// </summary>
        /// <value>The report time.</value>
        public DateTime ReportTime
        {
            get
            {
                return reportTime;
            }
        }

        /// <summary>
        /// Gets the last verb execution time.
        /// </summary>
        /// <value>The last verb execution time.</value>
        public DateTime LastVerbExecutionTime
        {
            get
            {
                return lastVerbExecutionTime;
            }
        }

        #endregion
    }
}
