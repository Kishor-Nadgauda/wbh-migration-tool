using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class EnumUtil
    {
        internal static Gender ParseGender(short gender)
        {
            if (gender == 1)
            {
                return Gender.Female;
            }
            else if (gender == 2)
            {
                return Gender.Male;
            }
            else
            {
                return Gender.None;
            }
        }

        internal static short ParseGender(Gender gender)
        {
            if (gender == Gender.Female)
            {
                return 1;
            }
            else if (gender == Gender.Male)
            {
                return 2;
            }
            else
            {
                return 0;
            }
        }

        internal static DisplayType ParseDisplayType(uint type)
        {
            if (type == 0)
            {
                return DisplayType.MailUser;
            }
            else if (type == 1)
            {
                return DisplayType.DistributionList;
            }
            else if (type == 2)
            {
                return DisplayType.Forum;
            }
            else if (type == 3)
            {
                return DisplayType.Agent;
            }
            else if (type == 4)
            {
                return DisplayType.Organization;
            }
            else if (type == 5)
            {
                return DisplayType.PrivateDistributionList;
            }
            else if (type == 6)
            {
                return DisplayType.RemoteMailUser;
            }
            else if (type == 0x01000000)
            {
                return DisplayType.Folder;
            }
            else if (type == 0x02000000)
            {
                return DisplayType.FolderLink;
            }
            else if (type == 0x04000000)
            {
                return DisplayType.FolderSpecial;
            }
            else if (type == 0x00010000)
            {
                return DisplayType.Modifiable;
            }
            else if (type == 0x00020000)
            {
                return DisplayType.GlobalAddressBook;
            }
            else if (type == 0x00030000)
            {
                return DisplayType.LocalAddressBook;
            }
            else if (type == 0x00040000)
            {
                return DisplayType.WideAreaNetworkAddressBook;
            }
            else if (type == 0x00050000)
            {
                return DisplayType.NotSpecific;
            }
            else
            {
                return DisplayType.None;
            }
        }

        internal static uint ParseDisplayType(DisplayType type)
        {
            if (type == DisplayType.MailUser)
            {
                return 0;
            }
            else if (type == DisplayType.DistributionList)
            {
                return 1;
            }
            else if (type == DisplayType.Forum)
            {
                return 2;
            }
            else if (type == DisplayType.Agent)
            {
                return 3;
            }
            else if (type == DisplayType.Organization)
            {
                return 4;
            }
            else if (type == DisplayType.PrivateDistributionList)
            {
                return 5;
            }
            else if (type == DisplayType.RemoteMailUser)
            {
                return 6;
            }
            else if (type == DisplayType.Folder)
            {
                return 0x01000000;
            }
            else if (type == DisplayType.FolderLink)
            {
                return 0x02000000;
            }
            else if (type == DisplayType.FolderSpecial)
            {
                return 0x04000000;
            }
            else if (type == DisplayType.Modifiable)
            {
                return 0x00010000;
            }
            else if (type == DisplayType.GlobalAddressBook)
            {
                return 0x00020000;
            }
            else if (type == DisplayType.LocalAddressBook)
            {
                return 0x00030000;
            }
            else if (type == DisplayType.WideAreaNetworkAddressBook)
            {
                return 0x00040000;
            }
            else if (type == DisplayType.NotSpecific)
            {
                return 0x00050000;
            }
            else
            {
                return 0;
            }
        }

        internal static RecipientType ParseRecipientType(uint type)
        {
            if (type == 1)
            {
                return RecipientType.To;
            }
            else if (type == 2)
            {
                return RecipientType.Cc;
            }
            else if (type == 3)
            {
                return RecipientType.Bcc;
            }
            else if (type == 0x10000000)
            {
                return RecipientType.P1;
            }
            else
            {
                return RecipientType.None;
            }
        }

        internal static uint ParseRecipientType(RecipientType type)
        {
            if (type == RecipientType.To)
            {
                return 1;
            }
            else if (type == RecipientType.Cc)
            {
                return 2;
            }
            else if (type == RecipientType.Bcc)
            {
                return 3;
            }
            else if (type == RecipientType.P1)
            {
                return 0x10000000;
            }
            else
            {
                return 0;
            }
        }


        internal static AttachmentMethod ParseAttachmentMethod(uint method)
        {
            if (method == 0)
            {
                return AttachmentMethod.NoAttachment;
            }
            else if (method == 1)
            {
                return AttachmentMethod.AttachByValue;
            }
            else if (method == 2)
            {
                return AttachmentMethod.AttachByReference;
            }
            else if (method == 3)
            {
                return AttachmentMethod.AttachByReferenceResolve;
            }
            else if (method == 4)
            {
                return AttachmentMethod.AttachByReferenceOnly;
            }
            else if (method == 5)
            {
                return AttachmentMethod.EmbeddedMessage;
            }
            else if (method == 6)
            {
                return AttachmentMethod.Ole;
            }
            else
            {
                return AttachmentMethod.None;
            }
        }

        internal static uint ParseAttachmentMethod(AttachmentMethod method)
        {
            if (method == AttachmentMethod.NoAttachment)
            {
                return 0;
            }
            else if (method == AttachmentMethod.AttachByValue)
            {
                return 1;
            }
            else if (method == AttachmentMethod.AttachByReference)
            {
                return 2;
            }
            else if (method == AttachmentMethod.AttachByReferenceResolve)
            {
                return 3;
            }
            else if (method == AttachmentMethod.AttachByReferenceOnly)
            {
                return 4;
            }
            else if (method == AttachmentMethod.EmbeddedMessage)
            {
                return 5;
            }
            else if (method == AttachmentMethod.Ole)
            {
                return 6;
            }
            else
            {
                return 0;
            }
        }

        internal static AttachmentFlags ParseAttachmentFlags(uint flags)
        {
            if (flags == 1)
            {
                return AttachmentFlags.InvisibleInHtml;
            }
            else if (flags == 2)
            {
                return AttachmentFlags.InvisibleInRtf;
            }
            else
            {
                return AttachmentFlags.None;
            }
        }

        internal static uint ParseAttachmentFlags(AttachmentFlags flags)
        {
            if (flags == AttachmentFlags.InvisibleInHtml)
            {
                return 1;
            }
            else if (flags == AttachmentFlags.InvisibleInRtf)
            {
                return 2;
            }
            else
            {
                return 0;
            }
        }

        internal static byte ParseEncryptionType(EncryptionType type)
        {
            if (type == EncryptionType.Compressible)
            {
                return 1;
            }
            else if (type == EncryptionType.High)
            {
                return 2;
            }
            else
            {
                return 0;
            }
        }

        internal static EncryptionType ParseEncryptionType(byte type)
        {
            if (type == 1)
            {
                return EncryptionType.Compressible;
            }
            else if (type == 2)
            {
                return EncryptionType.High;
            }
            else
            {
                return EncryptionType.None;
            }
        }

        internal static ushort ParsePropertyType(PropertyType type)
        {
            if (type == PropertyType.ApplicationTime)
            {
                return 0x0007;
            }
            else if (type == PropertyType.ApplicationTimeArray)
            {
                return 0x1007;
            }
            else if (type == PropertyType.Binary)
            {
                return 0x0102;
            }
            else if (type == PropertyType.BinaryArray)
            {
                return 0x1102;
            }
            else if (type == PropertyType.Boolean)
            {
                return 0x000B;
            }
            else if (type == PropertyType.Currency)
            {
                return 0x0006;
            }
            else if (type == PropertyType.CurrencyArray)
            {
                return 0x1006;
            }
            else if (type == PropertyType.Double)
            {
                return 0x0005;
            }
            else if (type == PropertyType.DoubleArray)
            {
                return 0x1005;
            }
            else if (type == PropertyType.Error)
            {
                return 0x000A;
            }
            else if (type == PropertyType.Float)
            {
                return 0x0004;
            }
            else if (type == PropertyType.FloatArray)
            {
                return 0x1004;
            }
            else if (type == PropertyType.Guid)
            {
                return 0x0048;
            }
            else if (type == PropertyType.GuidArray)
            {
                return 0x1048;
            }
            else if (type == PropertyType.Integer)
            {
                return 0x0003;
            }
            else if (type == PropertyType.IntegerArray)
            {
                return 0x1003;
            }
            else if (type == PropertyType.Long)
            {
                return 0x0014;
            }
            else if (type == PropertyType.LongArray)
            {
                return 0x1014;
            }
            else if (type == PropertyType.Null)
            {
                return 0x0001;
            }
            else if (type == PropertyType.Object)
            {
                return 0x000D;
            }
            else if (type == PropertyType.Restriction)
            {
                return 0x00FD;
            }
            else if (type == PropertyType.RuleAction)
            {
                return 0x00FE;
            }
            else if (type == PropertyType.ServerId)
            {
                return 0x00FB;
            }
            else if (type == PropertyType.Short)
            {
                return 0x0002;
            }
            else if (type == PropertyType.ShortArray)
            {
                return 0x1002;
            }
            else if (type == PropertyType.String)
            {
                return 0x001F;
            }
            else if (type == PropertyType.String8)
            {
                return 0x001E;
            }
            else if (type == PropertyType.String8Array)
            {
                return 0x101E;
            }
            else if (type == PropertyType.StringArray)
            {
                return 0x101F;
            }
            else if (type == PropertyType.SystemTime)
            {
                return 0x0040;
            }
            else if (type == PropertyType.SystemTimeArray)
            {
                return 0x1040;
            }
            else
            {
                return 0x0000;
            }
        }

        internal static PropertyType ParsePropertyType(ushort type)
        {
            if (type == 0x0007)
            {
                return PropertyType.ApplicationTime;
            }
            else if (type == 0x1007)
            {
                return PropertyType.ApplicationTimeArray;
            }
            else if (type == 0x0102)
            {
                return PropertyType.Binary;
            }
            else if (type == 0x1102)
            {
                return PropertyType.BinaryArray;
            }
            else if (type == 0x000B)
            {
                return PropertyType.Boolean;
            }
            else if (type == 0x0006)
            {
                return PropertyType.Currency;
            }
            else if (type == 0x1006)
            {
                return PropertyType.CurrencyArray;
            }
            else if (type == 0x0005)
            {
                return PropertyType.Double;
            }
            else if (type == 0x1005)
            {
                return PropertyType.DoubleArray;
            }
            else if (type == 0x000A)
            {
                return PropertyType.Error;
            }
            else if (type == 0x0004)
            {
                return PropertyType.Float;
            }
            else if (type == 0x1004)
            {
                return PropertyType.FloatArray;
            }
            else if (type == 0x0048)
            {
                return PropertyType.Guid;
            }
            else if (type == 0x1048)
            {
                return PropertyType.GuidArray;
            }
            else if (type == 0x0003)
            {
                return PropertyType.Integer;
            }
            else if (type == 0x1003)
            {
                return PropertyType.IntegerArray;
            }
            else if (type == 0x0014)
            {
                return PropertyType.Long;
            }
            else if (type == 0x1014)
            {
                return PropertyType.LongArray;
            }
            else if (type == 0x0001)
            {
                return PropertyType.Null;
            }
            else if (type == 0x000D)
            {
                return  PropertyType.Object;
            }
            else if (type == 0x00FD)
            {
                return PropertyType.Restriction;
            }
            else if (type == 0x00FE)
            {
                return PropertyType.RuleAction;
            }
            else if (type == 0x00FB)
            {
                return PropertyType.ServerId;
            }
            else if (type == 0x0002)
            {
                return PropertyType.Short;
            }
            else if (type == 0x1002)
            {
                return PropertyType.ShortArray;
            }
            else if (type == 0x001F)
            {
                return PropertyType.String;
            }
            else if (type == 0x001E)
            {
                return PropertyType.String8;
            }
            else if (type == 0x101E)
            {
                return PropertyType.String8Array;
            }
            else if (type == 0x101F)
            {
                return PropertyType.StringArray;
            }
            else if (type == 0x0040)
            {
                return PropertyType.SystemTime;
            }
            else if (type == 0x1040)
            {
                return PropertyType.SystemTimeArray;
            }
            else
            {
                return PropertyType.Unspecified;
            }
        }

        internal static LastVerbExecuted ParseLastVerbExecuted(uint verb)
        {
            if (verb == 102)
            {
                return LastVerbExecuted.ReplyToSender;
            }
            else if (verb == 103)
            {
                return LastVerbExecuted.ReplyToAll;
            }
            else if (verb == 104)
            {
                return LastVerbExecuted.Forward;
            }
            else
            {
                return LastVerbExecuted.None;
            }
        }

        internal static uint ParseLastVerbExecuted(LastVerbExecuted verb)
        {
            if (verb == LastVerbExecuted.ReplyToSender)
            {
                return 102;
            }
            else if (verb == LastVerbExecuted.ReplyToAll)
            {
                return 103;
            }
            else if (verb == LastVerbExecuted.Forward)
            {
                return 104;
            }
            else
            {
                return 0;
            }
        }

        internal static ObjectType ParseObjectType(uint type)
        {
            if (type == 1)
            {
                return ObjectType.MessageStore;
            }
            else if (type == 2)
            {
                return ObjectType.AddressBook;
            }
            else if (type == 3)
            {
                return ObjectType.Folder;
            }
            else if (type == 4)
            {
                return ObjectType.AddressBookContainer;
            }
            else if (type == 5)
            {
                return ObjectType.Message;
            }
            else if (type == 6)
            {
                return ObjectType.MailUser;
            }
            else if (type == 7)
            {
                return ObjectType.Attachment;
            }
            else if (type == 8)
            {
                return ObjectType.DistributionList;
            }
            else if (type == 9)
            {
                return ObjectType.ProfileSelection;
            }
            else if (type == 0x0000000A)
            {
                return ObjectType.Status;
            }
            else if (type == 0x0000000B)
            {
                return ObjectType.Session;
            }
            else if (type == 0x0000000C)
            {
                return ObjectType.Form;
            }
            else
            {
                return ObjectType.None;
            }
        }

        internal static uint ParseObjectType(ObjectType type)
        {
            if (type == ObjectType.MessageStore)
            {
                return 1;
            }
            else if (type == ObjectType.AddressBook)
            {
                return 2;
            }
            else if (type == ObjectType.Folder)
            {
                return 3;
            }
            else if (type == ObjectType.AddressBookContainer)
            {
                return 4;
            }
            else if (type == ObjectType.Message)
            {
                return 5;
            }
            else if (type == ObjectType.MailUser)
            {
                return 6;
            }
            else if (type == ObjectType.Attachment)
            {
                return 7;
            }
            else if (type == ObjectType.DistributionList)
            {
                return 8;
            }
            else if (type == ObjectType.ProfileSelection)
            {
                return 9;
            }
            else if (type == ObjectType.Status)
            {
                return 0x0000000A;
            }
            else if (type == ObjectType.Session)
            {
                return 0x0000000B;
            }
            else if (type == ObjectType.Form)
            {
                return 0x0000000C;
            }
            else
            {
                return 0;
            }
        }

        internal static IList<MessageFlag> ParseMessageFlag(uint flags)
        {
            IList<MessageFlag> list = new List<MessageFlag>();

            if ((flags & 0x00000040) == 0x00000040)
            {
                list.Add(MessageFlag.Associated);
            }

            if ((flags & 0x00000020) == 0x00000020)
            {
                list.Add(MessageFlag.FromMe);
            }
            if ((flags & 0x00000010) == 0x00000010)
            {
                list.Add(MessageFlag.HasAttachment);
            }

            if ((flags & 0x00000200) == 0x00000200)
            {
                list.Add(MessageFlag.NonReadReportPending);
            }

            if ((flags & 0x00002000) == 0x00002000)
            {
                list.Add(MessageFlag.OriginInternet);
            }

            if ((flags & 0x00008000) == 0x00008000)
            {
                list.Add(MessageFlag.OriginMiscExternal);
            }

            if ((flags & 0x00001000) == 0x00001000)
            {
                list.Add(MessageFlag.OriginX400);
            }

            if ((flags & 0x00000001) == 0x00000001)
            {
                list.Add(MessageFlag.Read);
            }

            if ((flags & 0x00000080) == 0x00000080)
            {
                list.Add(MessageFlag.Resend);
            }

            if ((flags & 0x00000100) == 0x00000100)
            {
                list.Add(MessageFlag.ReadReportPending);
            }

            if ((flags & 0x00000004) == 0x00000004)
            {
                list.Add(MessageFlag.Submit);
            }

            if ((flags & 0x00000002) == 0x00000002)
            {
                list.Add(MessageFlag.Unmodified);
            }

            if ((flags & 0x00000008) == 0x00000008)
            {
                list.Add(MessageFlag.Unsent);
            }

            return list;
        }

        internal static uint ParseMessageFlag(IList<MessageFlag> flags)
        {
            uint messageFlags = 0;

            for (int i = 0; i < flags.Count; i++)
            {
                MessageFlag currentMessageFlag = flags[i];

                if (currentMessageFlag == MessageFlag.Associated)
                {
                    messageFlags += 0x00000040;
                }
                else if (currentMessageFlag == MessageFlag.FromMe)
                {
                    messageFlags += 0x00000020;
                }
                else if (currentMessageFlag == MessageFlag.HasAttachment)
                {
                    messageFlags += 0x00000010;
                }
                else if (currentMessageFlag == MessageFlag.NonReadReportPending)
                {
                    messageFlags += 0x00000200;
                }
                else if (currentMessageFlag == MessageFlag.OriginInternet)
                {
                    messageFlags += 0x00002000;
                }
                else if (currentMessageFlag == MessageFlag.OriginMiscExternal)
                {
                    messageFlags += 0x00008000;
                }
                else if (currentMessageFlag == MessageFlag.OriginX400)
                {
                    messageFlags += 0x00001000;
                }
                else if (currentMessageFlag == MessageFlag.Read)
                {
                    messageFlags += 0x00000001;
                }
                else if (currentMessageFlag == MessageFlag.Resend)
                {
                    messageFlags += 0x00000080;
                }
                else if (currentMessageFlag == MessageFlag.ReadReportPending)
                {
                    messageFlags += 0x00000100;
                }
                else if (currentMessageFlag == MessageFlag.Submit)
                {
                    messageFlags += 0x00000004;
                }
                else if (currentMessageFlag == MessageFlag.Unmodified)
                {
                    messageFlags += 0x00000002;
                }
                else if (currentMessageFlag == MessageFlag.Unsent)
                {
                    messageFlags += 0x00000008;
                }
            }

            return messageFlags;
        }

        internal static IList<Independentsoft.Msg.MessageFlag> ConvertMessageFlags(IList<MessageFlag> flags)
        {
            IList<Independentsoft.Msg.MessageFlag> newMessageFlags = new List<Independentsoft.Msg.MessageFlag>();

            for (int i = 0; i < flags.Count; i++)
            {
                if (flags[i] == MessageFlag.Associated)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Associated);
                }
                else if (flags[i] == MessageFlag.FromMe)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.FromMe);
                }
                else if (flags[i] == MessageFlag.HasAttachment)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.HasAttachment);
                }
                else if (flags[i] == MessageFlag.NonReadReportPending)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.NonReadReportPending);
                }
                else if (flags[i] == MessageFlag.OriginInternet)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.OriginInternet);
                }
                else if (flags[i] == MessageFlag.OriginMiscExternal)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.OriginMiscExternal);
                }
                else if (flags[i] == MessageFlag.OriginX400)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.OriginX400);
                }
                else if (flags[i] == MessageFlag.Read)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Read);
                }
                else if (flags[i] == MessageFlag.Resend)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Resend);
                }
                else if (flags[i] == MessageFlag.ReadReportPending)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.ReadReportPending);
                }
                else if (flags[i] == MessageFlag.Submit)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Submit);
                }
                else if (flags[i] == MessageFlag.Unmodified)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Unmodified);
                }
                else if (flags[i] == MessageFlag.Unsent)
                {
                    newMessageFlags.Add(Independentsoft.Msg.MessageFlag.Unsent);
                }
            }

            return newMessageFlags;
        }

        internal static FlagIcon ParseFlagIcon(uint icon)
        {
            if (icon == 1)
            {
                return FlagIcon.Purple;
            }
            else if (icon == 2)
            {
                return FlagIcon.Orange;
            }
            else if (icon == 3)
            {
                return FlagIcon.Green;
            }
            else if (icon == 4)
            {
                return FlagIcon.Yellow;
            }
            else if (icon == 5)
            {
                return FlagIcon.Blue;
            }
            else if (icon == 6)
            {
                return FlagIcon.Red;
            }
            else
            {
                return FlagIcon.None;
            }
        }

        internal static uint ParseFlagIcon(FlagIcon icon)
        {
            if (icon == FlagIcon.Purple)
            {
                return 1;
            }
            else if (icon == FlagIcon.Orange)
            {
                return 2;
            }
            else if (icon == FlagIcon.Green)
            {
                return 3;
            }
            else if (icon == FlagIcon.Yellow)
            {
                return 4;
            }
            else if (icon == FlagIcon.Blue)
            {
                return 5;
            }
            else if (icon == FlagIcon.Red)
            {
                return 6;
            }
            else
            {
                return 0;
            }
        }

        internal static SelectedMailingAddress ParseSelectedMailingAddress(uint address)
        {
            if (address == 1)
            {
                return SelectedMailingAddress.Home;
            }
            else if (address == 2)
            {
                return SelectedMailingAddress.Business;
            }
            else if (address == 3)
            {
                return SelectedMailingAddress.Other;
            }
            else
            {
                return SelectedMailingAddress.None;
            }
        }

        internal static uint ParseSelectedMailingAddress(SelectedMailingAddress address)
        {
            if (address == SelectedMailingAddress.Home)
            {
                return 1;
            }
            else if (address == SelectedMailingAddress.Business)
            {
                return 2;
            }
            else if (address == SelectedMailingAddress.Other)
            {
                return 3;
            }
            else
            {
                return 0;
            }
        }

        internal static FlagStatus ParseFlagStatus(uint status)
        {
            if (status == 1)
            {
                return FlagStatus.Complete;
            }
            else if (status == 2)
            {
                return FlagStatus.Marked;
            }
            else
            {
                return FlagStatus.None;
            }
        }

        internal static uint ParseFlagStatus(FlagStatus status)
        {
            if (status == FlagStatus.Complete)
            {
                return 1;
            }
            else if (status == FlagStatus.Marked)
            {
                return 2;
            }
            else
            {
                return 0;
            }
        }

        internal static NoteColor ParseNoteColor(uint color)
        {
            if (color == 0)
            {
                return NoteColor.Blue;
            }
            else if (color == 1)
            {
                return NoteColor.Green;
            }
            else if (color == 2)
            {
                return NoteColor.Pink;
            }
            else if (color == 3)
            {
                return NoteColor.Yellow;
            }
            else if (color == 4)
            {
                return NoteColor.White;
            }
            else
            {
                return NoteColor.None;
            }
        }

        internal static uint ParseNoteColor(NoteColor color)
        {
            if (color == NoteColor.Blue)
            {
                return 0;
            }
            else if (color == NoteColor.Green)
            {
                return 1;
            }
            else if (color == NoteColor.Pink)
            {
                return 2;
            }
            else if (color == NoteColor.Yellow)
            {
                return 3;
            }
            else if (color == NoteColor.White)
            {
                return 4;
            }
            else
            {
                return 0;
            }
        }

        internal static Priority ParsePriority(uint priority)
        {
            if (priority == 0xFFFFFFFF)
            {
                return Priority.Low;
            }
            else if (priority == 1)
            {
                return Priority.High;
            }
            else if (priority == 0)
            {
                return Priority.Normal;
            }
            else
            {
                return Priority.None;
            }
        }

        internal static uint ParsePriority(Priority priority)
        {
            if (priority == Priority.Low)
            {
                return 0xFFFFFFFF;
            }
            else if (priority == Priority.High)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        internal static Sensitivity ParseSensitivity(uint sensitivity)
        {
            if (sensitivity == 1)
            {
                return Sensitivity.Personal;
            }
            else if (sensitivity == 2)
            {
                return Sensitivity.Private;
            }
            else if (sensitivity == 3)
            {
                return Sensitivity.Confidential;
            }
            else
            {
                return Sensitivity.None;
            }
        }

        internal static uint ParseSensitivity(Sensitivity sensitivity)
        {
            if (sensitivity == Sensitivity.Personal)
            {
                return 1;
            }
            else if (sensitivity == Sensitivity.Private)
            {
                return 2;
            }
            else if (sensitivity == Sensitivity.Confidential)
            {
                return 3;
            }
            else
            {
                return 0;
            }
        }

        internal static Importance ParseImportance(uint importance)
        {
            if (importance == 1)
            {
                return Importance.Normal;
            }
            else if (importance == 2)
            {
                return Importance.High;
            }
            else if (importance == 0)
            {
                return Importance.Low;
            }
            else
            {
                return Importance.None;
            }
        }

        internal static uint ParseImportance(Importance importance)
        {
            if (importance == Importance.Normal)
            {
                return 1;
            }
            else if (importance == Importance.High)
            {
                return 2;
            }
            else if (importance == Importance.Low)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        internal static TaskOwnership ParseTaskOwnership(uint ownership)
        {
            if (ownership == 0)
            {
                return TaskOwnership.New;
            }
            else if (ownership == 1)
            {
                return TaskOwnership.Delegated;
            }
            else if (ownership == 2)
            {
                return TaskOwnership.Own;
            }
            else
            {
                return TaskOwnership.None;
            }
        }

        internal static uint ParseTaskOwnership(TaskOwnership ownership)
        {
            if (ownership == TaskOwnership.New)
            {
                return 0;
            }
            else if (ownership == TaskOwnership.Delegated)
            {
                return 1;
            }
            else if (ownership == TaskOwnership.Own)
            {
                return 2;
            }
            else
            {
                return 0;
            }
        }

        internal static TaskDelegationState ParseTaskDelegationState(uint state)
        {
            if (state == 0)
            {
                return TaskDelegationState.Owned;
            }
            else if (state == 1)
            {
                return TaskDelegationState.OwnNew;
            }
            else if (state == 2)
            {
                return TaskDelegationState.Accepted;
            }
            else if (state == 3)
            {
                return TaskDelegationState.Declined;
            }
            else if (state == 4)
            {
                return TaskDelegationState.NoMatch;
            }
            else
            {
                return TaskDelegationState.None;
            }
        }

        internal static uint ParseTaskDelegationState(TaskDelegationState state)
        {
            if (state == TaskDelegationState.Owned)
            {
                return 0;
            }
            else if (state == TaskDelegationState.OwnNew)
            {
                return 1;
            }
            else if (state == TaskDelegationState.Accepted)
            {
                return 2;
            }
            else if (state == TaskDelegationState.Declined)
            {
                return 3;
            }
            else if (state == TaskDelegationState.NoMatch)
            {
                return 4;
            }
            else
            {
                return 0;
            }
        }

        internal static BusyStatus ParseBusyStatus(uint status)
        {
            if (status == 0)
            {
                return BusyStatus.Free;
            }
            else if (status == 1)
            {
                return BusyStatus.Tentative;
            }
            else if (status == 2)
            {
                return BusyStatus.Busy;
            }
            else if (status == 3)
            {
                return BusyStatus.OutOfOffice;
            }
            else
            {
                return BusyStatus.None;
            }
        }

        internal static uint ParseBusyStatus(BusyStatus status)
        {
            if (status == BusyStatus.Free)
            {
                return 0;
            }
            else if (status == BusyStatus.Tentative)
            {
                return 1;
            }
            else if (status == BusyStatus.Busy)
            {
                return 2;
            }
            else if (status == BusyStatus.OutOfOffice)
            {
                return 3;
            }
            else
            {
                return 0;
            }
        }


        internal static RecurrenceType ParseRecurrenceType(uint type)
        {
            if (type == 1)
            {
                return RecurrenceType.Daily;
            }
            else if (type == 2)
            {
                return RecurrenceType.Weekly;
            }
            else if (type == 3)
            {
                return RecurrenceType.Monthly;
            }
            else if (type == 4)
            {
                return RecurrenceType.Yearly;
            }
            else if (type == 5)
            {
                return RecurrenceType.MonthNth;
            }
            else if (type == 6)
            {
                return RecurrenceType.YearNth;
            }
            else
            {
                return RecurrenceType.None;
            }
        }

        internal static uint ParseRecurrenceType(RecurrenceType type)
        {
            if (type == RecurrenceType.Daily)
            {
                return 1;
            }
            else if (type == RecurrenceType.Weekly)
            {
                return 2;
            }
            else if (type == RecurrenceType.Monthly)
            {
                return 3;
            }
            else if (type == RecurrenceType.Yearly)
            {
                return 4;
            }
            else if (type == RecurrenceType.MonthNth)
            {
                return 5;
            }
            else if (type == RecurrenceType.YearNth)
            {
                return 6;
            }
            else
            {
                return 0;
            }
        }

        internal static ResponseStatus ParseResponseStatus(uint status)
        {
            if (status == 1)
            {
                return ResponseStatus.Organized;
            }
            else if (status == 2)
            {
                return ResponseStatus.Tentative;
            }
            else if (status == 3)
            {
                return ResponseStatus.Accepted;
            }
            else if (status == 4)
            {
                return ResponseStatus.Declined;
            }
            else if (status == 5)
            {
                return ResponseStatus.NotResponded;
            }
            else
            {
                return ResponseStatus.None;
            }
        }

        internal static uint ParseResponseStatus(ResponseStatus status)
        {
            if (status == ResponseStatus.Organized)
            {
                return 1;
            }
            else if (status == ResponseStatus.Tentative)
            {
                return 2;
            }
            else if (status == ResponseStatus.Accepted)
            {
                return 3;
            }
            else if (status == ResponseStatus.Declined)
            {
                return 4;
            }
            else if (status == ResponseStatus.NotResponded)
            {
                return 5;
            }
            else
            {
                return 0;
            }
        }

        internal static MeetingStatus ParseMeetingStatus(uint status)
        {
            if (status == 0)
            {
                return MeetingStatus.NonMeeting;
            }
            else if (status == 1)
            {
                return MeetingStatus.Meeting;
            }
            else if (status == 3)
            {
                return MeetingStatus.Received;
            }
            else if (status == 4)
            {
                return MeetingStatus.CanceledOrganizer;
            }
            else if (status == 5)
            {
                return MeetingStatus.Canceled;
            }
            else
            {
                return MeetingStatus.None;
            }
        }

        internal static uint ParseMeetingStatus(MeetingStatus status)
        {
            if (status == MeetingStatus.NonMeeting)
            {
                return 0;
            }
            else if (status == MeetingStatus.Meeting)
            {
                return 1;
            }
            else if (status == MeetingStatus.Received)
            {
                return 3;
            }
            else if (status == MeetingStatus.CanceledOrganizer)
            {
                return 4;
            }
            else if (status == MeetingStatus.Canceled)
            {
                return 5;
            }
            else
            {
                return 0;
            }
        }

        internal static TaskStatus ParseTaskStatus(uint status)
        {
            if (status == 0)
            {
                return TaskStatus.NotStarted;
            }
            else if (status == 1)
            {
                return TaskStatus.InProgress;
            }
            else if (status == 2)
            {
                return TaskStatus.Completed;
            }
            else if (status == 3)
            {
                return TaskStatus.WaitingOnOthers;
            }
            else if (status == 4)
            {
                return TaskStatus.Deferred;
            }
            else
            {
                return TaskStatus.None;
            }
        }

        internal static uint ParseTaskStatus(TaskStatus status)
        {
            if (status == TaskStatus.NotStarted)
            {
                return 0;
            }
            else if (status == TaskStatus.InProgress)
            {
                return 1;
            }
            else if (status == TaskStatus.Completed)
            {
                return 2;
            }
            else if (status == TaskStatus.WaitingOnOthers)
            {
                return 3;
            }
            else if (status == TaskStatus.Deferred)
            {
                return 4;
            }
            else
            {
                return 0;
            }
        }
    }
}
