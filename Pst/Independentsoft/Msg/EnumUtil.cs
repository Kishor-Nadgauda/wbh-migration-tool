using System;
using System.Collections.Generic;

namespace Independentsoft.Msg
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
                return  0x0000000A;
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

        internal static IList<StoreSupportMask> ParseStoreSupportMask(uint mask)
        {
            IList<StoreSupportMask> list = new List<StoreSupportMask>();

            if ((mask & 0x00020000) == 0x00020000)
            {
                list.Add(StoreSupportMask.Ansi);
            }
            
            if ((mask & 0x00000020) == 0x00000020)
            {
                list.Add(StoreSupportMask.Attachments);
            }
            
            if ((mask & 0x00000400) == 0x00000400)
            {
                list.Add(StoreSupportMask.Categorize);
            }
            
            if ((mask & 0x00000010) == 0x00000010)
            {
                list.Add(StoreSupportMask.Create);
            }

            if ((mask & 0x00010000) == 0x00010000)
            {
                list.Add(StoreSupportMask.Html);
            }
            
            if ((mask & 0x00200000) == 0x00200000)
            {
                list.Add(StoreSupportMask.ItemProc);
            }
            
            if ((mask & 0x00080000) == 0x00080000)
            {
                list.Add(StoreSupportMask.LocalStore);
            }
            
            if ((mask & 0x00000008) == 0x00000008)
            {
                list.Add(StoreSupportMask.Modify);
            }
            
            if ((mask & 0x00000200) == 0x00000200)
            {
                list.Add(StoreSupportMask.MultiValueProperties);
            }
            
            if ((mask & 0x00000100) == 0x00000100)
            {
                list.Add(StoreSupportMask.Notify);
            }
            
            if ((mask & 0x00000040) == 0x00000040)
            {
                list.Add(StoreSupportMask.Ole);
            }
            
            if ((mask & 0x00004000) == 0x00004000)
            {
                list.Add(StoreSupportMask.PublicFolders);
            }
            
            if ((mask & 0x00800000) == 0x00800000)
            {
                list.Add(StoreSupportMask.Pusher);
            }
            
            if ((mask & 0x00000002) == 0x00000002)
            {
                list.Add(StoreSupportMask.ReadOnly);
            }
            
            if ((mask & 0x00001000) == 0x00001000)
            {
                list.Add(StoreSupportMask.Restrictions);
            }
            
            if ((mask & 0x00000800) == 0x00000800)
            {
                list.Add(StoreSupportMask.Rtf);
            }
            
            if ((mask & 0x00000004) == 0x00000004)
            {
                list.Add(StoreSupportMask.Search);
            }
            
            if ((mask & 0x00002000) == 0x00002000)
            {
                list.Add(StoreSupportMask.Sort);
            }
            
            if ((mask & 0x00000080) == 0x00000080)
            {
                list.Add(StoreSupportMask.Submit);
            }
            
            if ((mask & 0x00008000) == 0x00008000)
            {
                list.Add(StoreSupportMask.UncompressedRtf);
            }
            
            if ((mask & 0x00040000) == 0x00040000)
            {
                list.Add(StoreSupportMask.Unicode);
            }

            return list;
        }

        internal static uint ParseStoreSupportMask(IList<StoreSupportMask> supportMask)
        {
            uint storeSupportMasks = 0;

            for (int i = 0; i < supportMask.Count; i++)
            {
                StoreSupportMask currentMask = supportMask[i];

                if (currentMask == StoreSupportMask.Ansi)
                {
                    storeSupportMasks += 0x00020000;
                }
                else if (currentMask == StoreSupportMask.Attachments)
                {
                    storeSupportMasks += 0x00000020;
                }
                else if (currentMask == StoreSupportMask.Categorize)
                {
                    storeSupportMasks += 0x00000400;
                }
                else if (currentMask == StoreSupportMask.Create)
                {
                    storeSupportMasks += 0x00000010;
                }
                else if (currentMask == StoreSupportMask.Html)
                {
                    storeSupportMasks += 0x00010000;
                }
                else if (currentMask == StoreSupportMask.Html)
                {
                    storeSupportMasks += 0x00010000;
                }
                else if (currentMask == StoreSupportMask.ItemProc)
                {
                    storeSupportMasks += 0x00200000;
                }
                else if (currentMask == StoreSupportMask.LocalStore)
                {
                    storeSupportMasks += 0x00080000;
                }
                else if (currentMask == StoreSupportMask.Modify)
                {
                    storeSupportMasks += 0x00000008;
                }
                else if (currentMask == StoreSupportMask.MultiValueProperties)
                {
                    storeSupportMasks += 0x00000200;
                }
                else if (currentMask == StoreSupportMask.Notify)
                {
                    storeSupportMasks += 0x00000100;
                }
                else if (currentMask == StoreSupportMask.Ole)
                {
                    storeSupportMasks += 0x00000040;
                }
                else if (currentMask == StoreSupportMask.PublicFolders)
                {
                    storeSupportMasks += 0x00004000;
                }
                else if (currentMask == StoreSupportMask.Pusher)
                {
                    storeSupportMasks += 0x00800000;
                }
                else if (currentMask == StoreSupportMask.ReadOnly)
                {
                    storeSupportMasks += 0x00000002;
                }
                else if (currentMask == StoreSupportMask.Restrictions)
                {
                    storeSupportMasks += 0x00001000;
                }
                else if (currentMask == StoreSupportMask.Rtf)
                {
                    storeSupportMasks += 0x00000800;
                }
                else if (currentMask == StoreSupportMask.Search)
                {
                    storeSupportMasks += 0x00000004;
                }
                else if (currentMask == StoreSupportMask.Sort)
                {
                    storeSupportMasks += 0x00002000;
                }
                else if (currentMask == StoreSupportMask.Submit)
                {
                    storeSupportMasks += 0x00000080;
                }
                else if (currentMask == StoreSupportMask.UncompressedRtf)
                {
                    storeSupportMasks += 0x00008000;
                }
                else if (currentMask == StoreSupportMask.Unicode)
                {
                    storeSupportMasks += 0x00040000;
                }
            }

            return storeSupportMasks;
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

        internal static RecurrencePatternFrequency ParseRecurrencePatternFrequency(short frequency)
        {
            if (frequency == 0x200A)
            {
                return RecurrencePatternFrequency.Daily;
            }
            else if (frequency == 0x200B)
            {
                return RecurrencePatternFrequency.Weekly;
            }
            else if (frequency == 0x200C)
            {
                return RecurrencePatternFrequency.Monthly;
            }
            else
            {
                return RecurrencePatternFrequency.Yearly;
            }
        }

        internal static DayOfWeek ParseDayOfWeek(int dayOfWeek)
        {
            if (dayOfWeek == 0x00000000)
            {
                return DayOfWeek.Sunday;
            }
            else if (dayOfWeek == 0x00000001)
            {
                return DayOfWeek.Monday;
            }
            else if (dayOfWeek == 0x00000002)
            {
                return DayOfWeek.Tuesday;
            }
            else if (dayOfWeek == 0x00000003)
            {
                return DayOfWeek.Wednesday;
            }
            else if (dayOfWeek == 0x00000004)
            {
                return DayOfWeek.Thursday;
            }
            else if (dayOfWeek == 0x00000005)
            {
                return DayOfWeek.Friday;
            }
            else
            {
                return DayOfWeek.Saturday;
            }
        }

        internal static DayOfWeekIndex ParseDayOfWeekIndex(int index)
        {
            if (index == 0x00000001)
            {
                return DayOfWeekIndex.First;
            }
            else if (index == 0x00000002)
            {
                return DayOfWeekIndex.Second;
            }
            else if (index == 0x00000003)
            {
                return DayOfWeekIndex.Third;
            }
            else if (index == 0x00000004)
            {
                return DayOfWeekIndex.Fourth;
            }
            else
            {
                return DayOfWeekIndex.Last;
            }
        }

        internal static RecurrenceEndType ParseRecurrenceEndType(int end)
        {
            if (end == 0x00002021)
            {
                return RecurrenceEndType.EndAfterDate;
            }
            else if (end == 0x00002022)
            {
                return RecurrenceEndType.EndAfterNOccurrences;
            }
            else
            {
                return RecurrenceEndType.NeverEnd;
            }
        }

        internal static RecurrencePatternType ParseRecurrencePatternType(short type)
        {
            if (type == 0x0000)
            {
                return RecurrencePatternType.Day;
            }
            else if (type == 0x0001)
            {
                return RecurrencePatternType.Week;
            }
            else if (type == 0x0002)
            {
                return RecurrencePatternType.Month;
            }
            else if (type == 0x0003)
            {
                return RecurrencePatternType.MonthEnd;
            }
            else if (type == 0x0004)
            {
                return RecurrencePatternType.MonthNth;
            }
            else if (type == 0x200A)
            {
                return RecurrencePatternType.HijriMonth;
            }
            else if (type == 0x200B)
            {
                return RecurrencePatternType.HijriMonthEnd;
            }
            else
            {
                return RecurrencePatternType.HijriMonthNth;
            }
        }

        internal static CalendarType ParseCalendarType(short type)
        {
            if (type == 0x0001)
            {
                return CalendarType.Gregorian;
            }
            else if (type == 0x0002)
            {
                return CalendarType.GregorianUS;
            }
            else if (type == 0x0003)
            {
                return CalendarType.JapaneseEmperorEra;
            }
            else if (type == 0x0004)
            {
                return CalendarType.Taiwan;
            }
            else if (type == 0x0005)
            {
                return CalendarType.KoreanTangunEra;
            }
            else if (type == 0x0006)
            {
                return CalendarType.Hijri;
            }
            else if (type == 0x0007)
            {
                return CalendarType.Thai;
            }
            else if (type == 0x0008)
            {
                return CalendarType.HebrewLunar;
            }
            else if (type == 0x0009)
            {
                return CalendarType.GregorianMiddleEastFrench;
            }
            else if (type == 0x000A)
            {
                return CalendarType.GregorianArabic;
            }
            else if (type == 0x000B)
            {
                return CalendarType.GregorianTransliteratedEnglish;
            }
            else if (type == 0x000C)
            {
                return CalendarType.GregorianTransliteratedFrench;
            }
            else if (type == 0x000E)
            {
                return CalendarType.JapaneseLunar;
            }
            else if (type == 0x000F)
            {
                return CalendarType.ChineseLunar;
            }
            else if (type == 0x0010)
            {
                return CalendarType.SakaEra;
            }
            else if (type == 0x0011)
            {
                return CalendarType.LunarETOChinese;
            }
            else if (type == 0x0012)
            {
                return CalendarType.LunarETOKorean;
            }
            else if (type == 0x0013)
            {
                return CalendarType.LunarRokuyou;
            }
            else if (type == 0x0014)
            {
                return CalendarType.KoreanLunar;
            }
            else if (type == 0x0017)
            {
                return CalendarType.UmAlQura;
            }
            else
            {
                return CalendarType.None;
            }
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

        internal static PropertyType ParsePropertyType(uint tag)
        {
            ushort typeValue = (ushort)(tag & 0xFFFF);
            ushort rawTypeValue = (ushort)(typeValue & 0x0FFF);

            PropertyType type;

            if ((typeValue & 0xF000) != 0)
            {
                if (rawTypeValue == 2)
                {
                    type = PropertyType.MultipleInteger16;
                }
                else if (rawTypeValue == 3)
                {
                    type = PropertyType.MultipleInteger32;
                }
                else if (rawTypeValue == 4)
                {
                    type = PropertyType.MultipleFloating32;
                }
                else if (rawTypeValue == 5)
                {
                    type = PropertyType.MultipleFloating64;
                }
                else if (rawTypeValue == 6)
                {
                    type = PropertyType.MultipleCurrency;
                }
                else if (rawTypeValue == 7)
                {
                    type = PropertyType.MultipleFloatingTime;
                }
                else if (rawTypeValue == 0x14)
                {
                    type = PropertyType.MultipleInteger64;
                }
                else if (rawTypeValue == 0x1E)
                {
                    type = PropertyType.MultipleString8;
                }
                else if (rawTypeValue == 0x1F)
                {
                    type = PropertyType.MultipleString;
                }
                else if (rawTypeValue == 0x40)
                {
                    type = PropertyType.MultipleTime;
                }
                else if (rawTypeValue == 0x48)
                {
                    type = PropertyType.MultipleGuid;
                }
                else if (rawTypeValue == 0x102)
                {
                    type = PropertyType.MultipleBinary;
                }
                else
                {
                    type = PropertyType.MultipleString;
                }
            }
            else
            {
                if (rawTypeValue == 2)
                {
                    type = PropertyType.Integer16;
                }
                else if (rawTypeValue == 3)
                {
                    type = PropertyType.Integer32;
                }
                else if (rawTypeValue == 4)
                {
                    type = PropertyType.Floating32;
                }
                else if (rawTypeValue == 5)
                {
                    type = PropertyType.Floating64;
                }
                else if (rawTypeValue == 6)
                {
                    type = PropertyType.Currency;
                }
                else if (rawTypeValue == 7)
                {
                    type = PropertyType.FloatingTime;
                }
                else if (rawTypeValue == 7)
                {
                    type = PropertyType.ErrorCode;
                }
                else if (rawTypeValue == 0xB)
                {
                    type = PropertyType.Boolean;
                }
                else if (rawTypeValue == 0xD)
                {
                    type = PropertyType.Object;
                }
                else if (rawTypeValue == 0x14)
                {
                    type = PropertyType.Integer64;
                }
                else if (rawTypeValue == 0x1E)
                {
                    type = PropertyType.String8;
                }
                else if (rawTypeValue == 0x1F)
                {
                    type = PropertyType.String;
                }
                else if (rawTypeValue == 0x40)
                {
                    type = PropertyType.Time;
                }
                else if (rawTypeValue == 0x48)
                {
                    type = PropertyType.Guid;
                }
                else if (rawTypeValue == 0x102)
                {
                    type = PropertyType.Binary;
                }
                else
                {
                    type = PropertyType.String;
                }
            }

            return type;
        }
    }
}
