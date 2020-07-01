using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Util.
    /// </summary>
    internal class Util
    {
        internal static byte[] encryptionCompressible = {
	                            0x47, 0xf1, 0xb4, 0xe6, 0x0b, 0x6a, 0x72, 0x48, 0x85, 0x4e, 0x9e, 0xeb, 0xe2, 0xf8, 0x94, 0x53,
	                            0xe0, 0xbb, 0xa0, 0x02, 0xe8, 0x5a, 0x09, 0xab, 0xdb, 0xe3, 0xba, 0xc6, 0x7c, 0xc3, 0x10, 0xdd,
	                            0x39, 0x05, 0x96, 0x30, 0xf5, 0x37, 0x60, 0x82, 0x8c, 0xc9, 0x13, 0x4a, 0x6b, 0x1d, 0xf3, 0xfb,
	                            0x8f, 0x26, 0x97, 0xca, 0x91, 0x17, 0x01, 0xc4, 0x32, 0x2d, 0x6e, 0x31, 0x95, 0xff, 0xd9, 0x23,
	                            0xd1, 0x00, 0x5e, 0x79, 0xdc, 0x44, 0x3b, 0x1a, 0x28, 0xc5, 0x61, 0x57, 0x20, 0x90, 0x3d, 0x83,
	                            0xb9, 0x43, 0xbe, 0x67, 0xd2, 0x46, 0x42, 0x76, 0xc0, 0x6d, 0x5b, 0x7e, 0xb2, 0x0f, 0x16, 0x29,
	                            0x3c, 0xa9, 0x03, 0x54, 0x0d, 0xda, 0x5d, 0xdf, 0xf6, 0xb7, 0xc7, 0x62, 0xcd, 0x8d, 0x06, 0xd3,
	                            0x69, 0x5c, 0x86, 0xd6, 0x14, 0xf7, 0xa5, 0x66, 0x75, 0xac, 0xb1, 0xe9, 0x45, 0x21, 0x70, 0x0c,
	                            0x87, 0x9f, 0x74, 0xa4, 0x22, 0x4c, 0x6f, 0xbf, 0x1f, 0x56, 0xaa, 0x2e, 0xb3, 0x78, 0x33, 0x50,
	                            0xb0, 0xa3, 0x92, 0xbc, 0xcf, 0x19, 0x1c, 0xa7, 0x63, 0xcb, 0x1e, 0x4d, 0x3e, 0x4b, 0x1b, 0x9b,
	                            0x4f, 0xe7, 0xf0, 0xee, 0xad, 0x3a, 0xb5, 0x59, 0x04, 0xea, 0x40, 0x55, 0x25, 0x51, 0xe5, 0x7a,
	                            0x89, 0x38, 0x68, 0x52, 0x7b, 0xfc, 0x27, 0xae, 0xd7, 0xbd, 0xfa, 0x07, 0xf4, 0xcc, 0x8e, 0x5f,
	                            0xef, 0x35, 0x9c, 0x84, 0x2b, 0x15, 0xd5, 0x77, 0x34, 0x49, 0xb6, 0x12, 0x0a, 0x7f, 0x71, 0x88,
	                            0xfd, 0x9d, 0x18, 0x41, 0x7d, 0x93, 0xd8, 0x58, 0x2c, 0xce, 0xfe, 0x24, 0xaf, 0xde, 0xb8, 0x36,
	                            0xc8, 0xa1, 0x80, 0xa6, 0x99, 0x98, 0xa8, 0x2f, 0x0e, 0x81, 0x65, 0x73, 0xe4, 0xc2, 0xa2, 0x8a,
	                            0xd4, 0xe1, 0x11, 0xd0, 0x08, 0x8b, 0x2a, 0xf2, 0xed, 0x9a, 0x64, 0x3f, 0xc1, 0x6c, 0xf9, 0xec
                            };

        internal static byte[] encryptionHigh1 = {
	                            0x41, 0x36, 0x13, 0x62, 0xa8, 0x21, 0x6e, 0xbb, 0xf4, 0x16, 0xcc, 0x04, 0x7f, 0x64, 0xe8, 0x5d,
	                            0x1e, 0xf2, 0xcb, 0x2a, 0x74, 0xc5, 0x5e, 0x35, 0xd2, 0x95, 0x47, 0x9e, 0x96, 0x2d, 0x9a, 0x88,
	                            0x4c, 0x7d, 0x84, 0x3f, 0xdb, 0xac, 0x31, 0xb6, 0x48, 0x5f, 0xf6, 0xc4, 0xd8, 0x39, 0x8b, 0xe7,
	                            0x23, 0x3b, 0x38, 0x8e, 0xc8, 0xc1, 0xdf, 0x25, 0xb1, 0x20, 0xa5, 0x46, 0x60, 0x4e, 0x9c, 0xfb,
	                            0xaa, 0xd3, 0x56, 0x51, 0x45, 0x7c, 0x55, 0x00, 0x07, 0xc9, 0x2b, 0x9d, 0x85, 0x9b, 0x09, 0xa0,
	                            0x8f, 0xad, 0xb3, 0x0f, 0x63, 0xab, 0x89, 0x4b, 0xd7, 0xa7, 0x15, 0x5a, 0x71, 0x66, 0x42, 0xbf,
	                            0x26, 0x4a, 0x6b, 0x98, 0xfa, 0xea, 0x77, 0x53, 0xb2, 0x70, 0x05, 0x2c, 0xfd, 0x59, 0x3a, 0x86,
	                            0x7e, 0xce, 0x06, 0xeb, 0x82, 0x78, 0x57, 0xc7, 0x8d, 0x43, 0xaf, 0xb4, 0x1c, 0xd4, 0x5b, 0xcd,
	                            0xe2, 0xe9, 0x27, 0x4f, 0xc3, 0x08, 0x72, 0x80, 0xcf, 0xb0, 0xef, 0xf5, 0x28, 0x6d, 0xbe, 0x30,
	                            0x4d, 0x34, 0x92, 0xd5, 0x0e, 0x3c, 0x22, 0x32, 0xe5, 0xe4, 0xf9, 0x9f, 0xc2, 0xd1, 0x0a, 0x81,
	                            0x12, 0xe1, 0xee, 0x91, 0x83, 0x76, 0xe3, 0x97, 0xe6, 0x61, 0x8a, 0x17, 0x79, 0xa4, 0xb7, 0xdc,
	                            0x90, 0x7a, 0x5c, 0x8c, 0x02, 0xa6, 0xca, 0x69, 0xde, 0x50, 0x1a, 0x11, 0x93, 0xb9, 0x52, 0x87,
	                            0x58, 0xfc, 0xed, 0x1d, 0x37, 0x49, 0x1b, 0x6a, 0xe0, 0x29, 0x33, 0x99, 0xbd, 0x6c, 0xd9, 0x94,
	                            0xf3, 0x40, 0x54, 0x6f, 0xf0, 0xc6, 0x73, 0xb8, 0xd6, 0x3e, 0x65, 0x18, 0x44, 0x1f, 0xdd, 0x67,
	                            0x10, 0xf1, 0x0c, 0x19, 0xec, 0xae, 0x03, 0xa1, 0x14, 0x7b, 0xa9, 0x0b, 0xff, 0xf8, 0xa3, 0xc0,
	                            0xa2, 0x01, 0xf7, 0x2e, 0xbc, 0x24, 0x68, 0x75, 0x0d, 0xfe, 0xba, 0x2f, 0xb5, 0xd0, 0xda, 0x3d
                            };

        internal static byte[] encryptionHigh2 = {
	                            0x14, 0x53, 0x0f, 0x56, 0xb3, 0xc8, 0x7a, 0x9c, 0xeb, 0x65, 0x48, 0x17, 0x16, 0x15, 0x9f, 0x02,
	                            0xcc, 0x54, 0x7c, 0x83, 0x00, 0x0d, 0x0c, 0x0b, 0xa2, 0x62, 0xa8, 0x76, 0xdb, 0xd9, 0xed, 0xc7,
	                            0xc5, 0xa4, 0xdc, 0xac, 0x85, 0x74, 0xd6, 0xd0, 0xa7, 0x9b, 0xae, 0x9a, 0x96, 0x71, 0x66, 0xc3,
	                            0x63, 0x99, 0xb8, 0xdd, 0x73, 0x92, 0x8e, 0x84, 0x7d, 0xa5, 0x5e, 0xd1, 0x5d, 0x93, 0xb1, 0x57,
	                            0x51, 0x50, 0x80, 0x89, 0x52, 0x94, 0x4f, 0x4e, 0x0a, 0x6b, 0xbc, 0x8d, 0x7f, 0x6e, 0x47, 0x46,
	                            0x41, 0x40, 0x44, 0x01, 0x11, 0xcb, 0x03, 0x3f, 0xf7, 0xf4, 0xe1, 0xa9, 0x8f, 0x3c, 0x3a, 0xf9,
	                            0xfb, 0xf0, 0x19, 0x30, 0x82, 0x09, 0x2e, 0xc9, 0x9d, 0xa0, 0x86, 0x49, 0xee, 0x6f, 0x4d, 0x6d,
	                            0xc4, 0x2d, 0x81, 0x34, 0x25, 0x87, 0x1b, 0x88, 0xaa, 0xfc, 0x06, 0xa1, 0x12, 0x38, 0xfd, 0x4c,
	                            0x42, 0x72, 0x64, 0x13, 0x37, 0x24, 0x6a, 0x75, 0x77, 0x43, 0xff, 0xe6, 0xb4, 0x4b, 0x36, 0x5c,
	                            0xe4, 0xd8, 0x35, 0x3d, 0x45, 0xb9, 0x2c, 0xec, 0xb7, 0x31, 0x2b, 0x29, 0x07, 0x68, 0xa3, 0x0e,
	                            0x69, 0x7b, 0x18, 0x9e, 0x21, 0x39, 0xbe, 0x28, 0x1a, 0x5b, 0x78, 0xf5, 0x23, 0xca, 0x2a, 0xb0,
	                            0xaf, 0x3e, 0xfe, 0x04, 0x8c, 0xe7, 0xe5, 0x98, 0x32, 0x95, 0xd3, 0xf6, 0x4a, 0xe8, 0xa6, 0xea,
	                            0xe9, 0xf3, 0xd5, 0x2f, 0x70, 0x20, 0xf2, 0x1f, 0x05, 0x67, 0xad, 0x55, 0x10, 0xce, 0xcd, 0xe3,
	                            0x27, 0x3b, 0xda, 0xba, 0xd7, 0xc2, 0x26, 0xd4, 0x91, 0x1d, 0xd2, 0x1c, 0x22, 0x33, 0xf8, 0xfa,
	                            0xf1, 0x5a, 0xef, 0xcf, 0x90, 0xb6, 0x8b, 0xb5, 0xbd, 0xc0, 0xbf, 0x08, 0x97, 0x1e, 0x6c, 0xe2,
	                            0x61, 0xe0, 0xc6, 0xc1, 0x59, 0xab, 0xbb, 0x58, 0xde, 0x5f, 0xdf, 0x60, 0x79, 0x7e, 0xb2, 0x8a
                            };

        internal static byte[] DecryptCompressibleEncryption(byte[] encrypted)
        {
            byte[] descrypted = new byte[encrypted.Length];

            for (int i = 0; i < encrypted.Length; i++)
            {
                byte index = encrypted[i];
                descrypted[i] = encryptionCompressible[index];
            }

            return descrypted;
        }

        internal static byte[] DecryptHighEncryption(byte[] encrypted, ulong id)
        {
            byte[] descrypted = new byte[encrypted.Length];

            ushort salt = (ushort)(((id & 0xffff0000) >> 16) ^ (id & 0x0000ffff));

            for (int i = 0; i < encrypted.Length; i++)
            {
                byte lowerSalt = (byte)(salt & 0x00ff);
                byte upperSalt = (byte)((salt & 0xff00) >> 8);

                byte index = encrypted[i];
                index += lowerSalt;
                index = encryptionHigh1[index];
                index += upperSalt;
                index = encryptionHigh2[index];
                index -= upperSalt;
                index = encryptionCompressible[index];
                index -= lowerSalt;

                descrypted[i] = index;

                salt++;
            }

            return descrypted;
        }

        private static DateTime year1601 = new DateTime(1601, 1, 1);
        internal static DateTime GetDateTime(int minutes)
        {
            if (minutes > 0)
            {
                //DateTime year1601 = new DateTime(1601, 1, 1);

                try
                {
                    DateTime dateTimeValue = year1601.AddMinutes(minutes);
                    return dateTimeValue;
                }
                catch (Exception) //ignore wrong dates
                {
                }

                return DateTime.MinValue;
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        internal static byte[] Decrypt(PstFile pstFile, byte[] buffer, ulong nodeId)
        {
            if (pstFile.EncryptionType == EncryptionType.Compressible)
            {
                buffer = Util.DecryptCompressibleEncryption(buffer);
            }
            else if (pstFile.EncryptionType == EncryptionType.High)
            {
                buffer = Util.DecryptHighEncryption(buffer, nodeId);
            }

            return buffer;
        }

        internal static Item GetItem(PstFileReader reader, LocalDescriptorList localDescriptorList, Table table, ulong id, ulong parentId)
        {
            return GetItem(reader, localDescriptorList, table, id, parentId, true);
        }

        internal static Item GetItem(PstFileReader reader, LocalDescriptorList localDescriptorList, Table table, ulong id, ulong parentId, bool includeAttachments)
        {
            Item item = null;
            string messageClass = null;

            if (table.Entries[MapiPropertyTag.PR_MESSAGE_CLASS] != null)
            {
                messageClass = table.Entries[MapiPropertyTag.PR_MESSAGE_CLASS].GetStringValue();
            }

            if (messageClass != null && (messageClass == "IPM.Note" || messageClass.StartsWith("IPM.Note.")))
            {
                item = new Message();
            }
            else if (messageClass != null && (messageClass == "IPM.Appointment" || messageClass.StartsWith("IPM.Appointment.")))
            {
                item = new Appointment();
            }
            else if (messageClass != null && (messageClass == "IPM.Contact" || messageClass.StartsWith("IPM.Contact.")))
            {
                item = new Contact();
            }
            else if (messageClass != null && (messageClass == "IPM.Task" || messageClass.StartsWith("IPM.Task.")))
            {
                item = new Task();
            }
            else if (messageClass != null && (messageClass.StartsWith("IPM.Schedule.Meeting.") || messageClass.StartsWith("IPM.TaskRequest.")))
            {
                item = new CalendarMessage();
            }
            else if (messageClass != null && (messageClass == "IPM.Report" || messageClass.StartsWith("IPM.Report.") || messageClass.StartsWith("Report.") || messageClass.StartsWith("REPORT.")))
            {
                item = new ReportMessage();
            }
            else if (messageClass != null && messageClass.StartsWith("IPM.Document."))
            {
                item = new Document();
            }
            else if (messageClass != null && messageClass == "IPM.Activity")
            {
                item = new Journal();
            }
            else if (messageClass != null && messageClass == "IPM.Post")
            {
                item = new Post();
            }
            else if (messageClass != null && messageClass == "IPM.StickyNote")
            {
                item = new Note();
            }
            else if (messageClass != null && messageClass == "IPM.DistList")
            {
                item = new DistributionList();
            }
            else if (messageClass != null && messageClass == "IPM.TaskRequest")
            {
                item = new TaskRequest();
            }
            else
            {
                item = new Item();
            }

            item.reader = reader;
            item.id = id;
            item.parentId = parentId;
            item.table = table;
            item.messageClass = messageClass;

            //Compute EntryID
            if (reader.PstFile.MessageStore != null && reader.PstFile.MessageStore.RecordKey != null)
            {
                item.entryId = new byte[24];
                System.Array.Copy(reader.PstFile.MessageStore.RecordKey, 0, item.entryId, 4, 16);

                byte[] idBytes = BitConverter.GetBytes((uint)id);
                System.Array.Copy(idBytes, 0, item.entryId, 20, 4);
            }

            //set encoding based on Subject, Body or DisplayTo. We can include more properties.
            if (table.Entries[MapiPropertyTag.PR_SUBJECT] != null)
            {
                TableEntry subjectTableEntry = table.Entries[MapiPropertyTag.PR_SUBJECT];

                if (subjectTableEntry.PropertyTag != null && subjectTableEntry.PropertyTag.Type == PropertyType.String)
                {
                    item.encoding = System.Text.Encoding.Unicode;
                }
            }
            else if (table.Entries[MapiPropertyTag.PR_BODY] != null)
            {
                TableEntry bodyTableEntry = table.Entries[MapiPropertyTag.PR_BODY];

                if (bodyTableEntry.PropertyTag != null && bodyTableEntry.PropertyTag.Type == PropertyType.String)
                {
                    item.encoding = System.Text.Encoding.Unicode;
                }
            }
            else if (table.Entries[MapiPropertyTag.PR_DISPLAY_TO] != null)
            {
                TableEntry displayToTableEntry = table.Entries[MapiPropertyTag.PR_DISPLAY_TO];

                if (displayToTableEntry.PropertyTag != null && displayToTableEntry.PropertyTag.Type == PropertyType.String)
                {
                    item.encoding = System.Text.Encoding.Unicode;
                }
            }


            if (table.Entries[MapiPropertyTag.PR_SUBJECT] != null)
            {
                item.subject = table.Entries[MapiPropertyTag.PR_SUBJECT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SUBJECT_PREFIX] != null)
            {
                item.subjectPrefix = table.Entries[MapiPropertyTag.PR_SUBJECT_PREFIX].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CONVERSATION_TOPIC] != null)
            {
                item.conversationTopic = table.Entries[MapiPropertyTag.PR_CONVERSATION_TOPIC].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_BCC] != null)
            {
                item.displayBcc = table.Entries[MapiPropertyTag.PR_DISPLAY_BCC].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_CC] != null)
            {
                item.displayCc = table.Entries[MapiPropertyTag.PR_DISPLAY_CC].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_TO] != null)
            {
                item.displayTo = table.Entries[MapiPropertyTag.PR_DISPLAY_TO].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ORIGINAL_DISPLAY_TO] != null)
            {
                item.originalDisplayTo = table.Entries[MapiPropertyTag.PR_ORIGINAL_DISPLAY_TO].GetStringValue();
            }

            if (table.Entries[new PropertyTag(0x0050001f)] != null)
            {
                item.replyTo = table.Entries[new PropertyTag(0x0050001f)].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_NORMALIZED_SUBJECT] != null)
            {
                item.normalizedSubject = table.Entries[MapiPropertyTag.PR_NORMALIZED_SUBJECT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BODY.Tag, PropertyType.String] != null)
            {
                item.body = table.Entries[MapiPropertyTag.PR_BODY.Tag, PropertyType.String].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BODY.Tag, PropertyType.String8] != null)
            {
                item.body = table.Entries[MapiPropertyTag.PR_BODY.Tag, PropertyType.String8].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HTML] != null)
            {
                item.bodyHtml = table.Entries[MapiPropertyTag.PR_HTML].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RTF_COMPRESSED] != null)
            {
                item.rtfCompressed = table.Entries[MapiPropertyTag.PR_RTF_COMPRESSED].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SEARCH_KEY] != null)
            {
                item.searchKey = table.Entries[MapiPropertyTag.PR_SEARCH_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CREATION_TIME] != null)
            {
                item.creationTime = table.Entries[MapiPropertyTag.PR_CREATION_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_LAST_MODIFICATION_TIME] != null)
            {
                item.lastModificationTime = table.Entries[MapiPropertyTag.PR_LAST_MODIFICATION_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_MESSAGE_DELIVERY_TIME] != null)
            {
                item.messageDeliveryTime = table.Entries[MapiPropertyTag.PR_MESSAGE_DELIVERY_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CLIENT_SUBMIT_TIME] != null)
            {
                item.clientSubmitTime = table.Entries[MapiPropertyTag.PR_CLIENT_SUBMIT_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PROVIDER_SUBMIT_TIME] != null)
            {
                item.providerSubmitTime = table.Entries[MapiPropertyTag.PR_PROVIDER_SUBMIT_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_REPORT_TIME] != null)
            {
                item.reportTime = table.Entries[MapiPropertyTag.PR_REPORT_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_REPORT_TEXT] != null)
            {
                item.reportText = table.Entries[MapiPropertyTag.PR_REPORT_TEXT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CREATOR_NAME] != null)
            {
                item.creatorName = table.Entries[MapiPropertyTag.PR_CREATOR_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_LAST_MODIFIER_NAME] != null)
            {
                item.lastModifierName = table.Entries[MapiPropertyTag.PR_LAST_MODIFIER_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_INTERNET_MESSAGE_ID] != null)
            {
                item.internetMessageId = table.Entries[MapiPropertyTag.PR_INTERNET_MESSAGE_ID].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_IN_REPLY_TO_ID] != null)
            {
                item.inReplyTo = table.Entries[MapiPropertyTag.PR_IN_REPLY_TO_ID].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_INTERNET_REFERENCES] != null)
            {
                item.internetReferences = table.Entries[MapiPropertyTag.PR_INTERNET_REFERENCES].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_MESSAGE_CODEPAGE] != null)
            {
                item.messageCodePage = table.Entries[MapiPropertyTag.PR_MESSAGE_CODEPAGE].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ACTION] != null)
            {
                item.iconIndex = table.Entries[MapiPropertyTag.PR_ACTION].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_MESSAGE_SIZE] != null)
            {
                item.messageSize = table.Entries[MapiPropertyTag.PR_MESSAGE_SIZE].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_INTERNET_CPID] != null)
            {
                item.internetCodePage = table.Entries[MapiPropertyTag.PR_INTERNET_CPID].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CONVERSATION_INDEX] != null)
            {
                item.conversationIndex = table.Entries[MapiPropertyTag.PR_CONVERSATION_INDEX].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ATTR_HIDDEN] != null)
            {
                item.isHidden = table.Entries[MapiPropertyTag.PR_ATTR_HIDDEN].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ATTR_READONLY] != null)
            {
                item.isReadOnly = table.Entries[MapiPropertyTag.PR_ATTR_READONLY].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ATTR_SYSTEM] != null)
            {
                item.isSystem = table.Entries[MapiPropertyTag.PR_ATTR_SYSTEM].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISABLE_FULL_FIDELITY] != null)
            {
                item.disableFullFidelity = table.Entries[MapiPropertyTag.PR_DISABLE_FULL_FIDELITY].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HASATTACH] != null)
            {
                item.hasAttachment = table.Entries[MapiPropertyTag.PR_HASATTACH].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RTF_IN_SYNC] != null)
            {
                item.rtfInSync = table.Entries[MapiPropertyTag.PR_RTF_IN_SYNC].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENSITIVITY] != null)
            {
                int sensitivityValue = table.Entries[MapiPropertyTag.PR_SENSITIVITY].GetIntegerValue();
                item.sensitivity = EnumUtil.ParseSensitivity((uint)sensitivityValue);
            }

            if (table.Entries[MapiPropertyTag.PR_IMPORTANCE] != null)
            {
                int importanceValue = table.Entries[MapiPropertyTag.PR_IMPORTANCE].GetIntegerValue();
                item.importance = EnumUtil.ParseImportance((uint)importanceValue);
            }

            if (table.Entries[MapiPropertyTag.PR_PRIORITY] != null)
            {
                int priorityValue = table.Entries[MapiPropertyTag.PR_PRIORITY].GetIntegerValue();
                item.priority = EnumUtil.ParsePriority((uint)priorityValue);
            }

            if (table.Entries[MapiPropertyTag.PR_FLAG_ICON] != null)
            {
                int flagIconValue = table.Entries[MapiPropertyTag.PR_FLAG_ICON].GetIntegerValue();
                item.flagIcon = EnumUtil.ParseFlagIcon((uint)flagIconValue);
            }

            if (table.Entries[MapiPropertyTag.PR_FLAG_STATUS] != null)
            {
                int flagStatusValue = table.Entries[MapiPropertyTag.PR_FLAG_STATUS].GetIntegerValue();
                item.flagStatus = EnumUtil.ParseFlagStatus((uint)flagStatusValue);
            }

            if (table.Entries[MapiPropertyTag.PR_OBJECT_TYPE] != null)
            {
                int objectTypeValue = table.Entries[MapiPropertyTag.PR_OBJECT_TYPE].GetIntegerValue();
                item.objectType = EnumUtil.ParseObjectType((uint)objectTypeValue);
            }

            if (table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_ADDRTYPE] != null)
            {
                item.receivedRepresentingAddressType = table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_ADDRTYPE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_EMAIL_ADDRESS] != null)
            {
                item.receivedRepresentingEmailAddress = table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_EMAIL_ADDRESS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_ENTRYID] != null)
            {
                item.receivedRepresentingEntryId = table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_ENTRYID].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_NAME] != null)
            {
                item.receivedRepresentingName = table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_SEARCH_KEY] != null)
            {
                item.receivedRepresentingSearchKey = table.Entries[MapiPropertyTag.PR_RCVD_REPRESENTING_SEARCH_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECEIVED_BY_ADDRTYPE] != null)
            {
                item.receivedByAddressType = table.Entries[MapiPropertyTag.PR_RECEIVED_BY_ADDRTYPE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS] != null)
            {
                item.receivedByEmailAddress = table.Entries[MapiPropertyTag.PR_RECEIVED_BY_EMAIL_ADDRESS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECEIVED_BY_ENTRYID] != null)
            {
                item.receivedByEntryId = table.Entries[MapiPropertyTag.PR_RECEIVED_BY_ENTRYID].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECEIVED_BY_NAME] != null)
            {
                item.receivedByName = table.Entries[MapiPropertyTag.PR_RECEIVED_BY_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECEIVED_BY_SEARCH_KEY] != null)
            {
                item.receivedBySearchKey = table.Entries[MapiPropertyTag.PR_RECEIVED_BY_SEARCH_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENDER_ADDRTYPE] != null)
            {
                item.senderAddressType = table.Entries[MapiPropertyTag.PR_SENDER_ADDRTYPE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.EX_SENDER_EMAIL_ADDRESS] != null)
            {
                item.senderEmailAddress = table.Entries[MapiPropertyTag.EX_SENDER_EMAIL_ADDRESS].GetStringValue();
            }

            if (string.IsNullOrWhiteSpace(item.senderEmailAddress) && table.Entries[MapiPropertyTag.PR_SENDER_EMAIL_ADDRESS] != null)
            {
                item.senderEmailAddress = table.Entries[MapiPropertyTag.PR_SENDER_EMAIL_ADDRESS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENDER_ENTRYID] != null)
            {
                item.senderEntryId = table.Entries[MapiPropertyTag.PR_SENDER_ENTRYID].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENDER_NAME] != null)
            {
                item.senderName = table.Entries[MapiPropertyTag.PR_SENDER_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENDER_SEARCH_KEY] != null)
            {
                item.senderSearchKey = table.Entries[MapiPropertyTag.PR_SENDER_SEARCH_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_ADDRTYPE] != null)
            {
                item.sentRepresentingAddressType = table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_ADDRTYPE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS] != null)
            {
                item.sentRepresentingEmailAddress = table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_EMAIL_ADDRESS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_ENTRYID] != null)
            {
                item.sentRepresentingEntryId = table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_ENTRYID].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_NAME] != null)
            {
                item.sentRepresentingName = table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_SEARCH_KEY] != null)
            {
                item.sentRepresentingSearchKey = table.Entries[MapiPropertyTag.PR_SENT_REPRESENTING_SEARCH_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_TRANSPORT_MESSAGE_HEADERS] != null)
            {
                item.transportMessageHeaders = table.Entries[MapiPropertyTag.PR_TRANSPORT_MESSAGE_HEADERS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_LAST_VERB_EXECUTION_TIME] != null)
            {
                item.lastVerbExecutionTime = table.Entries[MapiPropertyTag.PR_LAST_VERB_EXECUTION_TIME].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_LAST_VERB_EXECUTED] != null)
            {
                int lastVerbExecutedValue = table.Entries[MapiPropertyTag.PR_LAST_VERB_EXECUTED].GetIntegerValue();
                item.lastVerbExecuted = EnumUtil.ParseLastVerbExecuted((uint)lastVerbExecutedValue);
            }

            if (table.Entries[MapiPropertyTag.PR_MESSAGE_FLAGS] != null)
            {
                int messageFlagsValue = table.Entries[MapiPropertyTag.PR_MESSAGE_FLAGS].GetIntegerValue();
                item.messageFlags = EnumUtil.ParseMessageFlag((uint)messageFlagsValue);
            }

            if (table.Entries[MapiPropertyTag.PR_BIRTHDAY] != null)
            {
                item.birthday = table.Entries[MapiPropertyTag.PR_BIRTHDAY].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CHILDRENS_NAMES] != null)
            {
                item.childrenNames = table.Entries[MapiPropertyTag.PR_CHILDRENS_NAMES].GetStringArrayValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ASSISTANT] != null)
            {
                item.assistentName = table.Entries[MapiPropertyTag.PR_ASSISTANT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ASSISTANT_TELEPHONE_NUMBER] != null)
            {
                item.assistentPhone = table.Entries[MapiPropertyTag.PR_ASSISTANT_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_TELEPHONE_NUMBER] != null)
            {
                item.businessPhone = table.Entries[MapiPropertyTag.PR_BUSINESS_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS2_TELEPHONE_NUMBER] != null)
            {
                item.businessPhone2 = table.Entries[MapiPropertyTag.PR_BUSINESS2_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_FAX_NUMBER] != null)
            {
                item.businessFax = table.Entries[MapiPropertyTag.PR_BUSINESS_FAX_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_HOME_PAGE] != null)
            {
                item.businessHomePage = table.Entries[MapiPropertyTag.PR_BUSINESS_HOME_PAGE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CALLBACK_TELEPHONE_NUMBER] != null)
            {
                item.callbackPhone = table.Entries[MapiPropertyTag.PR_CALLBACK_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CAR_TELEPHONE_NUMBER] != null)
            {
                item.carPhone = table.Entries[MapiPropertyTag.PR_CAR_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CELLULAR_TELEPHONE_NUMBER] != null)
            {
                item.cellularPhone = table.Entries[MapiPropertyTag.PR_CELLULAR_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_COMPANY_MAIN_PHONE_NUMBER] != null)
            {
                item.companyMainPhone = table.Entries[MapiPropertyTag.PR_COMPANY_MAIN_PHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_COMPANY_NAME] != null)
            {
                item.companyName = table.Entries[MapiPropertyTag.PR_COMPANY_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_COMPUTER_NETWORK_NAME] != null)
            {
                item.computerNetworkName = table.Entries[MapiPropertyTag.PR_COMPUTER_NETWORK_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CUSTOMER_ID] != null)
            {
                item.customerId = table.Entries[MapiPropertyTag.PR_CUSTOMER_ID].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DEPARTMENT_NAME] != null)
            {
                item.departmentName = table.Entries[MapiPropertyTag.PR_DEPARTMENT_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_NAME] != null)
            {
                item.displayName = table.Entries[MapiPropertyTag.PR_DISPLAY_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_NAME_PREFIX] != null)
            {
                item.displayNamePrefix = table.Entries[MapiPropertyTag.PR_DISPLAY_NAME_PREFIX].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_FTP_SITE] != null)
            {
                item.ftpSite = table.Entries[MapiPropertyTag.PR_FTP_SITE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_GENERATION] != null)
            {
                item.generation = table.Entries[MapiPropertyTag.PR_GENERATION].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_GIVEN_NAME] != null)
            {
                item.givenName = table.Entries[MapiPropertyTag.PR_GIVEN_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_GOVERNMENT_ID_NUMBER] != null)
            {
                item.governmentId = table.Entries[MapiPropertyTag.PR_GOVERNMENT_ID_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOBBIES] != null)
            {
                item.hobbies = table.Entries[MapiPropertyTag.PR_HOBBIES].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME2_TELEPHONE_NUMBER] != null)
            {
                item.homePhone2 = table.Entries[MapiPropertyTag.PR_HOME2_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_CITY] != null)
            {
                item.homeAddressCity = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_CITY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_COUNTRY] != null)
            {
                item.homeAddressCountry = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_COUNTRY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_POSTAL_CODE] != null)
            {
                item.homeAddressPostalCode = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_POSTAL_CODE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_POST_OFFICE_BOX] != null)
            {
                item.homeAddressPostOfficeBox = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_POST_OFFICE_BOX].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_STATE_OR_PROVINCE] != null)
            {
                item.homeAddressState = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_STATE_OR_PROVINCE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_STREET] != null)
            {
                item.homeAddressStreet = table.Entries[MapiPropertyTag.PR_HOME_ADDRESS_STREET].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_FAX_NUMBER] != null)
            {
                item.homeFax = table.Entries[MapiPropertyTag.PR_HOME_FAX_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_HOME_TELEPHONE_NUMBER] != null)
            {
                item.homePhone = table.Entries[MapiPropertyTag.PR_HOME_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_INITIALS] != null)
            {
                item.initials = table.Entries[MapiPropertyTag.PR_INITIALS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_ISDN_NUMBER] != null)
            {
                item.isdn = table.Entries[MapiPropertyTag.PR_ISDN_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_MANAGER_NAME] != null)
            {
                item.managerName = table.Entries[MapiPropertyTag.PR_MANAGER_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_MIDDLE_NAME] != null)
            {
                item.middleName = table.Entries[MapiPropertyTag.PR_MIDDLE_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_NICKNAME] != null)
            {
                item.nickname = table.Entries[MapiPropertyTag.PR_NICKNAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OFFICE_LOCATION] != null)
            {
                item.officeLocation = table.Entries[MapiPropertyTag.PR_OFFICE_LOCATION].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_CITY] != null)
            {
                item.otherAddressCity = table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_CITY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_COUNTRY] != null)
            {
                item.otherAddressCountry = table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_COUNTRY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_POSTAL_CODE] != null)
            {
                item.otherAddressPostalCode = table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_POSTAL_CODE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_STATE_OR_PROVINCE] != null)
            {
                item.otherAddressState = table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_STATE_OR_PROVINCE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_STREET] != null)
            {
                item.otherAddressStreet = table.Entries[MapiPropertyTag.PR_OTHER_ADDRESS_STREET].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_OTHER_TELEPHONE_NUMBER] != null)
            {
                item.otherPhone = table.Entries[MapiPropertyTag.PR_OTHER_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PAGER_TELEPHONE_NUMBER] != null)
            {
                item.pager = table.Entries[MapiPropertyTag.PR_PAGER_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PERSONAL_HOME_PAGE] != null)
            {
                item.personalHomePage = table.Entries[MapiPropertyTag.PR_PERSONAL_HOME_PAGE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_POSTAL_ADDRESS] != null)
            {
                item.postalAddress = table.Entries[MapiPropertyTag.PR_POSTAL_ADDRESS].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_COUNTRY] != null)
            {
                item.businessAddressCountry = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_COUNTRY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_CITY] != null)
            {
                item.businessAddressCity = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_CITY].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_POSTAL_CODE] != null)
            {
                item.businessAddressPostalCode = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_POSTAL_CODE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_POST_OFFICE_BOX] != null)
            {
                item.businessAddressPostOfficeBox = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_POST_OFFICE_BOX].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE] != null)
            {
                item.businessAddressState = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_STREET] != null)
            {
                item.businessAddressStreet = table.Entries[MapiPropertyTag.PR_BUSINESS_ADDRESS_STREET].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PRIMARY_FAX_NUMBER] != null)
            {
                item.primaryFax = table.Entries[MapiPropertyTag.PR_PRIMARY_FAX_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PRIMARY_TELEPHONE_NUMBER] != null)
            {
                item.primaryPhone = table.Entries[MapiPropertyTag.PR_PRIMARY_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PROFESSION] != null)
            {
                item.profession = table.Entries[MapiPropertyTag.PR_PROFESSION].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RADIO_TELEPHONE_NUMBER] != null)
            {
                item.radioPhone = table.Entries[MapiPropertyTag.PR_RADIO_TELEPHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SPOUSE_NAME] != null)
            {
                item.spouseName = table.Entries[MapiPropertyTag.PR_SPOUSE_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SURNAME] != null)
            {
                item.surname = table.Entries[MapiPropertyTag.PR_SURNAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_TELEX_NUMBER] != null)
            {
                item.telex = table.Entries[MapiPropertyTag.PR_TELEX_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_TITLE] != null)
            {
                item.title = table.Entries[MapiPropertyTag.PR_TITLE].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_TTYTDD_PHONE_NUMBER] != null)
            {
                item.ttyTddPhone = table.Entries[MapiPropertyTag.PR_TTYTDD_PHONE_NUMBER].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_WEDDING_ANNIVERSARY] != null)
            {
                item.weddingAnniversary = table.Entries[MapiPropertyTag.PR_WEDDING_ANNIVERSARY].GetDateTimeValue();
            }

            if (table.Entries[MapiPropertyTag.PR_GENDER] != null)
            {
                item.gender = EnumUtil.ParseGender(table.Entries[MapiPropertyTag.PR_GENDER].GetShortValue());
            }

            uint outlookVersionNamedPropertyId = 0x8554;

            if (reader.PstFile.NamedIdToTag.ContainsKey(outlookVersionNamedPropertyId))
            {
                long outlookVersionTag = reader.PstFile.NamedIdToTag[outlookVersionNamedPropertyId];

                if (outlookVersionTag > -1)
                {
                    outlookVersionTag = (outlookVersionTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(outlookVersionTag)] != null)
                    {
                        item.outlookVersion = table.Entries[new PropertyTag(outlookVersionTag)].GetStringValue();
                    }
                }
            }

            uint outlookInternalVersionNamedPropertyId = 0x8552;

            if (reader.PstFile.NamedIdToTag.ContainsKey(outlookInternalVersionNamedPropertyId))
            {
                long outlookInternalVersionTag = reader.PstFile.NamedIdToTag[outlookInternalVersionNamedPropertyId];

                if (outlookInternalVersionTag > -1)
                {
                    outlookInternalVersionTag = (outlookInternalVersionTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(outlookInternalVersionTag)] != null)
                    {
                        item.outlookInternalVersion = table.Entries[new PropertyTag(outlookInternalVersionTag)].GetIntegerValue();
                    }
                }
            }

            uint commonStartTimeNamedPropertyId = 0x8516;

            if (reader.PstFile.NamedIdToTag.ContainsKey(commonStartTimeNamedPropertyId))
            {
                long commonStartTimeTag = reader.PstFile.NamedIdToTag[commonStartTimeNamedPropertyId];

                if (commonStartTimeTag > -1)
                {
                    commonStartTimeTag = (commonStartTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(commonStartTimeTag)] != null)
                    {
                        item.commonStartTime = table.Entries[new PropertyTag(commonStartTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint commonEndTimeNamedPropertyId = 0x8517;

            if (reader.PstFile.NamedIdToTag.ContainsKey(commonEndTimeNamedPropertyId))
            {
                long commonEndTimeTag = reader.PstFile.NamedIdToTag[commonEndTimeNamedPropertyId];

                if (commonEndTimeTag > -1)
                {
                    commonEndTimeTag = (commonEndTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(commonEndTimeTag)] != null)
                    {
                        item.commonEndTime = table.Entries[new PropertyTag(commonEndTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint flagDueByNamedPropertyId = 0x8560;

            if (reader.PstFile.NamedIdToTag.ContainsKey(flagDueByNamedPropertyId))
            {
                long flagDueByTag = reader.PstFile.NamedIdToTag[flagDueByNamedPropertyId];

                if (flagDueByTag > -1)
                {
                    flagDueByTag = (flagDueByTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(flagDueByTag)] != null)
                    {
                        item.flagDueBy = table.Entries[new PropertyTag(flagDueByTag)].GetDateTimeValue();
                    }
                }
            }

            uint isRecurringNamedPropertyId = 0x8223;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isRecurringNamedPropertyId))
            {
                long isRecurringTag = reader.PstFile.NamedIdToTag[isRecurringNamedPropertyId];

                if (isRecurringTag > -1)
                {
                    isRecurringTag = (isRecurringTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isRecurringTag)] != null)
                    {
                        item.isRecurring = table.Entries[new PropertyTag(isRecurringTag)].GetBooleanValue();
                    }
                }
            }

            uint reminderTimeNamedPropertyId = 0x8502;

            if (reader.PstFile.NamedIdToTag.ContainsKey(reminderTimeNamedPropertyId))
            {
                long reminderTimeTag = reader.PstFile.NamedIdToTag[reminderTimeNamedPropertyId];

                if (reminderTimeTag > -1)
                {
                    reminderTimeTag = (reminderTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(reminderTimeTag)] != null)
                    {
                        item.reminderTime = table.Entries[new PropertyTag(reminderTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint reminderMinutesBeforeStartNamedPropertyId = 0x8501;

            if (reader.PstFile.NamedIdToTag.ContainsKey(reminderMinutesBeforeStartNamedPropertyId))
            {
                long reminderMinutesBeforeStartTag = reader.PstFile.NamedIdToTag[reminderMinutesBeforeStartNamedPropertyId];

                if (reminderMinutesBeforeStartTag > -1)
                {
                    reminderMinutesBeforeStartTag = (reminderMinutesBeforeStartTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(reminderMinutesBeforeStartTag)] != null)
                    {
                        item.reminderMinutesBeforeStart = table.Entries[new PropertyTag(reminderMinutesBeforeStartTag)].GetIntegerValue();
                    }
                }
            }

            uint companiesNamedPropertyId = 0x8539;

            if (reader.PstFile.NamedIdToTag.ContainsKey(companiesNamedPropertyId))
            {
                long companiesTag = reader.PstFile.NamedIdToTag[companiesNamedPropertyId];

                if (companiesTag > -1)
                {
                    companiesTag = (companiesTag << 16) + 0x101E;

                    if (table.Entries[new PropertyTag(companiesTag)] != null)
                    {
                        item.companies = table.Entries[new PropertyTag(companiesTag)].GetStringArrayValue();
                    }
                }
            }

            uint contactNamesNamedPropertyId = 0x853A;

            if (reader.PstFile.NamedIdToTag.ContainsKey(contactNamesNamedPropertyId))
            {
                long contactNamesTag = reader.PstFile.NamedIdToTag[contactNamesNamedPropertyId];

                if (contactNamesTag > -1)
                {
                    contactNamesTag = (contactNamesTag << 16) + 0x101E;

                    if (table.Entries[new PropertyTag(contactNamesTag)] != null)
                    {
                        item.contactNames = table.Entries[new PropertyTag(contactNamesTag)].GetStringArrayValue();
                    }
                }
            }

            uint billingInformationNamedPropertyId = 0x8535;

            if (reader.PstFile.NamedIdToTag.ContainsKey(billingInformationNamedPropertyId))
            {
                long billingInformationTag = reader.PstFile.NamedIdToTag[billingInformationNamedPropertyId];

                if (billingInformationTag > -1)
                {
                    billingInformationTag = (billingInformationTag << 16) + 0x101E;

                    if (table.Entries[new PropertyTag(billingInformationTag)] != null)
                    {
                        item.billingInformation = table.Entries[new PropertyTag(billingInformationTag)].GetStringValue();
                    }
                }
            }

            uint mileageNamedPropertyId = 0x8534;

            if (reader.PstFile.NamedIdToTag.ContainsKey(mileageNamedPropertyId))
            {
                long mileageTag = reader.PstFile.NamedIdToTag[mileageNamedPropertyId];

                if (mileageTag > -1)
                {
                    mileageTag = (mileageTag << 16) + 0x101E;

                    if (table.Entries[new PropertyTag(mileageTag)] != null)
                    {
                        item.mileage = table.Entries[new PropertyTag(mileageTag)].GetStringValue();
                    }
                }
            }

            uint reminderSoundFileNamedPropertyId = 0x851F;

            if (reader.PstFile.NamedIdToTag.ContainsKey(reminderSoundFileNamedPropertyId))
            {
                long reminderSoundFileTag = reader.PstFile.NamedIdToTag[reminderSoundFileNamedPropertyId];

                if (reminderSoundFileTag > -1)
                {
                    reminderSoundFileTag = (reminderSoundFileTag << 16) + 0x101E;

                    if (table.Entries[new PropertyTag(reminderSoundFileTag)] != null)
                    {
                        item.reminderSoundFile = table.Entries[new PropertyTag(reminderSoundFileTag)].GetStringValue();
                    }
                }
            }

            uint isPrivateNamedPropertyId = 0x8506;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isPrivateNamedPropertyId))
            {
                long isPrivateTag = reader.PstFile.NamedIdToTag[isPrivateNamedPropertyId];

                if (isPrivateTag > -1)
                {
                    isPrivateTag = (isPrivateTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isPrivateTag)] != null)
                    {
                        item.isPrivate = table.Entries[new PropertyTag(isPrivateTag)].GetBooleanValue();
                    }
                }
            }

            uint isReminderSetNamedPropertyId = 0x8503;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isReminderSetNamedPropertyId))
            {
                long isReminderSetTag = reader.PstFile.NamedIdToTag[isReminderSetNamedPropertyId];

                if (isReminderSetTag > -1)
                {
                    isReminderSetTag = (isReminderSetTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isReminderSetTag)] != null)
                    {
                        item.isReminderSet = table.Entries[new PropertyTag(isReminderSetTag)].GetBooleanValue();
                    }
                }
            }

            uint reminderOverrideDefaultNamedPropertyId = 0x851C;

            if (reader.PstFile.NamedIdToTag.ContainsKey(reminderOverrideDefaultNamedPropertyId))
            {
                long reminderOverrideDefaultTag = reader.PstFile.NamedIdToTag[reminderOverrideDefaultNamedPropertyId];

                if (reminderOverrideDefaultTag > -1)
                {
                    reminderOverrideDefaultTag = (reminderOverrideDefaultTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(reminderOverrideDefaultTag)] != null)
                    {
                        item.reminderOverrideDefault = table.Entries[new PropertyTag(reminderOverrideDefaultTag)].GetBooleanValue();
                    }
                }
            }

            uint reminderPlaySoundNamedPropertyId = 0x851E;

            if (reader.PstFile.NamedIdToTag.ContainsKey(reminderPlaySoundNamedPropertyId))
            {
                long reminderPlaySoundTag = reader.PstFile.NamedIdToTag[reminderPlaySoundNamedPropertyId];

                if (reminderPlaySoundTag > -1)
                {
                    reminderPlaySoundTag = (reminderPlaySoundTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(reminderPlaySoundTag)] != null)
                    {
                        item.reminderPlaySound = table.Entries[new PropertyTag(reminderPlaySoundTag)].GetBooleanValue();
                    }
                }
            }

            uint appointmentStartTimeNamedPropertyId = 0x820D;

            if (reader.PstFile.NamedIdToTag.ContainsKey(appointmentStartTimeNamedPropertyId))
            {
                long appointmentStartTimeTag = reader.PstFile.NamedIdToTag[appointmentStartTimeNamedPropertyId];

                if (appointmentStartTimeTag > -1)
                {
                    appointmentStartTimeTag = (appointmentStartTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(appointmentStartTimeTag)] != null)
                    {
                        item.appointmentStartTime = table.Entries[new PropertyTag(appointmentStartTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint appointmentEndTimeNamedPropertyId = 0x820E;

            if (reader.PstFile.NamedIdToTag.ContainsKey(appointmentEndTimeNamedPropertyId))
            {
                long appointmentEndTimeTag = reader.PstFile.NamedIdToTag[appointmentEndTimeNamedPropertyId];

                if (appointmentEndTimeTag > -1)
                {
                    appointmentEndTimeTag = (appointmentEndTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(appointmentEndTimeTag)] != null)
                    {
                        item.appointmentEndTime = table.Entries[new PropertyTag(appointmentEndTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint isAllDayEventNamedPropertyId = 0x8215;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isAllDayEventNamedPropertyId))
            {
                long isAllDayEventTag = reader.PstFile.NamedIdToTag[isAllDayEventNamedPropertyId];

                if (isAllDayEventTag > -1)
                {
                    isAllDayEventTag = (isAllDayEventTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isAllDayEventTag)] != null)
                    {
                        item.isAllDayEvent = table.Entries[new PropertyTag(isAllDayEventTag)].GetBooleanValue();
                    }
                }
            }

            uint locationNamedPropertyId = 0x8208;

            if (reader.PstFile.NamedIdToTag.ContainsKey(locationNamedPropertyId))
            {
                long locationTag = reader.PstFile.NamedIdToTag[locationNamedPropertyId];

                if (locationTag > -1)
                {
                    locationTag = (locationTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(locationTag)] != null)
                    {
                        item.location = table.Entries[new PropertyTag(locationTag)].GetStringValue();
                    }
                }
            }

            uint busyStatusNamedPropertyId = 0x8205;

            if (reader.PstFile.NamedIdToTag.ContainsKey(busyStatusNamedPropertyId))
            {
                long busyStatusTag = reader.PstFile.NamedIdToTag[busyStatusNamedPropertyId];

                if (busyStatusTag > -1)
                {
                    busyStatusTag = (busyStatusTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(busyStatusTag)] != null)
                    {
                        int busyStatusValue = table.Entries[new PropertyTag(busyStatusTag)].GetIntegerValue();

                        item.busyStatus = EnumUtil.ParseBusyStatus((uint)busyStatusValue);
                    }
                }
            }

            uint meetingStatusNamedPropertyId = 0x8217;

            if (reader.PstFile.NamedIdToTag.ContainsKey(meetingStatusNamedPropertyId))
            {
                long meetingStatusTag = reader.PstFile.NamedIdToTag[meetingStatusNamedPropertyId];

                if (meetingStatusTag > -1)
                {
                    meetingStatusTag = (meetingStatusTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(meetingStatusTag)] != null)
                    {
                        int meetingStatusValue = table.Entries[new PropertyTag(meetingStatusTag)].GetIntegerValue();

                        item.meetingStatus = EnumUtil.ParseMeetingStatus((uint)meetingStatusValue);
                    }
                }
            }

            uint responseStatusNamedPropertyId = 0x8218;

            if (reader.PstFile.NamedIdToTag.ContainsKey(responseStatusNamedPropertyId))
            {
                long responseStatusTag = reader.PstFile.NamedIdToTag[responseStatusNamedPropertyId];

                if (responseStatusTag > -1)
                {
                    responseStatusTag = (responseStatusTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(responseStatusTag)] != null)
                    {
                        int responseStatusValue = table.Entries[new PropertyTag(responseStatusTag)].GetIntegerValue();

                        item.responseStatus = EnumUtil.ParseResponseStatus((uint)responseStatusValue);
                    }
                }
            }

            uint recurrenceTypeNamedPropertyId = 0x8231;

            if (reader.PstFile.NamedIdToTag.ContainsKey(recurrenceTypeNamedPropertyId))
            {
                long recurrenceTypeTag = reader.PstFile.NamedIdToTag[recurrenceTypeNamedPropertyId];

                if (recurrenceTypeTag > -1)
                {
                    recurrenceTypeTag = (recurrenceTypeTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(recurrenceTypeTag)] != null)
                    {
                        int recurrenceTypeValue = table.Entries[new PropertyTag(recurrenceTypeTag)].GetIntegerValue();

                        item.recurrenceType = EnumUtil.ParseRecurrenceType((uint)recurrenceTypeValue);
                    }
                }
            }

            uint appointmentMessageClassNamedPropertyId = 0x24;

            if (reader.PstFile.NamedIdToTag.ContainsKey(appointmentMessageClassNamedPropertyId))
            {
                long appointmentMessageClassTag = reader.PstFile.NamedIdToTag[appointmentMessageClassNamedPropertyId];

                if (appointmentMessageClassTag > -1)
                {
                    appointmentMessageClassTag = (appointmentMessageClassTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(appointmentMessageClassTag)] != null)
                    {
                        item.appointmentMessageClass = table.Entries[new PropertyTag(appointmentMessageClassTag)].GetStringValue();
                    }
                }
            }

            uint timeZoneNamedPropertyId = 0x8234;

            if (reader.PstFile.NamedIdToTag.ContainsKey(timeZoneNamedPropertyId))
            {
                long timeZoneTag = reader.PstFile.NamedIdToTag[timeZoneNamedPropertyId];

                if (timeZoneTag > -1)
                {
                    timeZoneTag = (timeZoneTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(timeZoneTag)] != null)
                    {
                        item.timeZone = table.Entries[new PropertyTag(timeZoneTag)].GetStringValue();
                    }
                }
            }

            uint recurrencePatternDescriptionNamedPropertyId = 0x8232;

            if (reader.PstFile.NamedIdToTag.ContainsKey(recurrencePatternDescriptionNamedPropertyId))
            {
                long recurrencePatternDescriptionTag = reader.PstFile.NamedIdToTag[recurrencePatternDescriptionNamedPropertyId];

                if (recurrencePatternDescriptionTag > -1)
                {
                    recurrencePatternDescriptionTag = (recurrencePatternDescriptionTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(recurrencePatternDescriptionTag)] != null)
                    {
                        item.recurrencePatternDescription = table.Entries[new PropertyTag(recurrencePatternDescriptionTag)].GetStringValue();
                    }
                }
            }

            uint recurrencePatternNamedPropertyId = 0x8216;

            if (reader.PstFile.NamedIdToTag.ContainsKey(recurrencePatternNamedPropertyId))
            {
                long recurrencePatternTag = reader.PstFile.NamedIdToTag[recurrencePatternNamedPropertyId];

                if (recurrencePatternTag > -1)
                {
                    recurrencePatternTag = (recurrencePatternTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(recurrencePatternTag)] != null)
                    {
                        byte[] buffer = table.Entries[new PropertyTag(recurrencePatternTag)].GetBinaryValue();
                        
                        if(buffer != null)
                        {
                          try
                          {
                            item.recurrencePattern = new Independentsoft.Msg.RecurrencePattern(buffer);
                          }
                          catch (Exception e)
                          {
                            System.Diagnostics.Trace.TraceWarning("Util.GetItem({0}): Parsing recurrencePattern {1}", item.id, e.Message);
                          }
                        }
                    }
                }
            }

            uint taskRecurrencePatternNamedPropertyId = 0x8116;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskRecurrencePatternNamedPropertyId))
            {
                long recurrencePatternTag = reader.PstFile.NamedIdToTag[taskRecurrencePatternNamedPropertyId];

                if (recurrencePatternTag > -1)
                {
                    recurrencePatternTag = (recurrencePatternTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(recurrencePatternTag)] != null)
                    {
                        byte[] buffer = table.Entries[new PropertyTag(recurrencePatternTag)].GetBinaryValue();

                        if (buffer != null)
                        {
                            item.recurrencePattern = new Independentsoft.Msg.RecurrencePattern(buffer);
                        }
                    }
                }
            }

            uint guidNamedPropertyId = 0x3;

            if (reader.PstFile.NamedIdToTag.ContainsKey(guidNamedPropertyId))
            {
                long guidTag = reader.PstFile.NamedIdToTag[guidNamedPropertyId];

                if (guidTag > -1)
                {
                    guidTag = (guidTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(guidTag)] != null)
                    {
                        item.guid = table.Entries[new PropertyTag(guidTag)].GetBinaryValue();
                    }
                }
            }

            uint labelNamedPropertyId = 0x8214;

            if (reader.PstFile.NamedIdToTag.ContainsKey(labelNamedPropertyId))
            {
                long labelTag = reader.PstFile.NamedIdToTag[labelNamedPropertyId];

                if (labelTag > -1)
                {
                    labelTag = (labelTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(labelTag)] != null)
                    {
                        item.label = table.Entries[new PropertyTag(labelTag)].GetIntegerValue();
                    }
                }
            }

            uint durationNamedPropertyId = 0x8213;

            if (reader.PstFile.NamedIdToTag.ContainsKey(durationNamedPropertyId))
            {
                long durationTag = reader.PstFile.NamedIdToTag[durationNamedPropertyId];

                if (durationTag > -1)
                {
                    durationTag = (durationTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(durationTag)] != null)
                    {
                        item.duration = table.Entries[new PropertyTag(durationTag)].GetIntegerValue();
                    }
                }
            }

            uint taskStartDateNamedPropertyId = 0x8104;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskStartDateNamedPropertyId))
            {
                long taskStartDateTag = reader.PstFile.NamedIdToTag[taskStartDateNamedPropertyId];

                if (taskStartDateTag > -1)
                {
                    taskStartDateTag = (taskStartDateTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(taskStartDateTag)] != null)
                    {
                        item.taskStartDate = table.Entries[new PropertyTag(taskStartDateTag)].GetDateTimeValue();
                    }
                }
            }

            uint taskDueDateNamedPropertyId = 0x8105;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskDueDateNamedPropertyId))
            {
                long taskDueDateTag = reader.PstFile.NamedIdToTag[taskDueDateNamedPropertyId];

                if (taskDueDateTag > -1)
                {
                    taskDueDateTag = (taskDueDateTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(taskDueDateTag)] != null)
                    {
                        item.taskDueDate = table.Entries[new PropertyTag(taskDueDateTag)].GetDateTimeValue();
                    }
                }
            }

            uint ownerNamedPropertyId = 0x811F;

            if (reader.PstFile.NamedIdToTag.ContainsKey(ownerNamedPropertyId))
            {
                long ownerTag = reader.PstFile.NamedIdToTag[ownerNamedPropertyId];

                if (ownerTag > -1)
                {
                    ownerTag = (ownerTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(ownerTag)] != null)
                    {
                        item.owner = table.Entries[new PropertyTag(ownerTag)].GetStringValue();
                    }
                }
            }

            uint delegatorNamedPropertyId = 0x8121;

            if (reader.PstFile.NamedIdToTag.ContainsKey(delegatorNamedPropertyId))
            {
                long delegatorTag = reader.PstFile.NamedIdToTag[delegatorNamedPropertyId];

                if (delegatorTag > -1)
                {
                    delegatorTag = (delegatorTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(delegatorTag)] != null)
                    {
                        item.delegator = table.Entries[new PropertyTag(delegatorTag)].GetStringValue();
                    }
                }
            }

            uint percentCompleteNamedPropertyId = 0x8102;

            if (reader.PstFile.NamedIdToTag.ContainsKey(percentCompleteNamedPropertyId))
            {
                long percentCompleteTag = reader.PstFile.NamedIdToTag[percentCompleteNamedPropertyId];

                if (percentCompleteTag > -1)
                {
                    percentCompleteTag = (percentCompleteTag << 16) + 0x0005;

                    if (table.Entries[new PropertyTag(percentCompleteTag)] != null)
                    {
                        item.percentComplete = table.Entries[new PropertyTag(percentCompleteTag)].GetDoubleValue();
                        item.percentComplete = item.percentComplete * 100;
                    }
                }
            }

            uint actualWorkNamedPropertyId = 0x8110;

            if (reader.PstFile.NamedIdToTag.ContainsKey(actualWorkNamedPropertyId))
            {
                long actualWorkTag = reader.PstFile.NamedIdToTag[actualWorkNamedPropertyId];

                if (actualWorkTag > -1)
                {
                    actualWorkTag = (actualWorkTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(actualWorkTag)] != null)
                    {
                        item.actualWork = table.Entries[new PropertyTag(actualWorkTag)].GetIntegerValue();
                    }
                }
            }

            uint totalWorkNamedPropertyId = 0x8111;

            if (reader.PstFile.NamedIdToTag.ContainsKey(totalWorkNamedPropertyId))
            {
                long totalWorkTag = reader.PstFile.NamedIdToTag[totalWorkNamedPropertyId];

                if (totalWorkTag > -1)
                {
                    totalWorkTag = (totalWorkTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(totalWorkTag)] != null)
                    {
                        item.totalWork = table.Entries[new PropertyTag(totalWorkTag)].GetIntegerValue();
                    }
                }
            }

            uint isTeamTaskNamedPropertyId = 0x8103;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isTeamTaskNamedPropertyId))
            {
                long isTeamTaskTag = reader.PstFile.NamedIdToTag[isTeamTaskNamedPropertyId];

                if (isTeamTaskTag > -1)
                {
                    isTeamTaskTag = (isTeamTaskTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isTeamTaskTag)] != null)
                    {
                        item.isTeamTask = table.Entries[new PropertyTag(isTeamTaskTag)].GetBooleanValue();
                    }
                }
            }

            uint isCompleteNamedPropertyId = 0x811C;

            if (reader.PstFile.NamedIdToTag.ContainsKey(isCompleteNamedPropertyId))
            {
                long isCompleteTag = reader.PstFile.NamedIdToTag[isCompleteNamedPropertyId];

                if (isCompleteTag > -1)
                {
                    isCompleteTag = (isCompleteTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(isCompleteTag)] != null)
                    {
                        item.isComplete = table.Entries[new PropertyTag(isCompleteTag)].GetBooleanValue();
                    }
                }
            }

            uint dateCompletedNamedPropertyId = 0x810F;

            if (reader.PstFile.NamedIdToTag.ContainsKey(dateCompletedNamedPropertyId))
            {
                long dateCompletedTag = reader.PstFile.NamedIdToTag[dateCompletedNamedPropertyId];

                if (dateCompletedTag > -1)
                {
                    dateCompletedTag = (dateCompletedTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(dateCompletedTag)] != null)
                    {
                        item.dateCompleted = table.Entries[new PropertyTag(dateCompletedTag)].GetDateTimeValue();
                    }
                }
            }

            uint taskStatusNamedPropertyId = 0x8101;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskStatusNamedPropertyId))
            {
                long taskStatusTag = reader.PstFile.NamedIdToTag[taskStatusNamedPropertyId];

                if (taskStatusTag > -1)
                {
                    taskStatusTag = (taskStatusTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(taskStatusTag)] != null)
                    {
                        int taskStatusValue = table.Entries[new PropertyTag(taskStatusTag)].GetIntegerValue();

                        item.taskStatus = EnumUtil.ParseTaskStatus((uint)taskStatusValue);
                    }
                }
            }

            uint taskOwnershipNamedPropertyId = 0x8129;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskOwnershipNamedPropertyId))
            {
                long taskOwnershipTag = reader.PstFile.NamedIdToTag[taskOwnershipNamedPropertyId];

                if (taskOwnershipTag > -1)
                {
                    taskOwnershipTag = (taskOwnershipTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(taskOwnershipTag)] != null)
                    {
                        int taskOwnershipValue = table.Entries[new PropertyTag(taskOwnershipTag)].GetIntegerValue();

                        item.taskOwnership = EnumUtil.ParseTaskOwnership((uint)taskOwnershipValue);
                    }
                }
            }

            uint taskDelegationStateNamedPropertyId = 0x812A;

            if (reader.PstFile.NamedIdToTag.ContainsKey(taskDelegationStateNamedPropertyId))
            {
                long taskDelegationStateTag = reader.PstFile.NamedIdToTag[taskDelegationStateNamedPropertyId];

                if (taskDelegationStateTag > -1)
                {
                    taskDelegationStateTag = (taskDelegationStateTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(taskDelegationStateTag)] != null)
                    {
                        int taskDelegationStateValue = table.Entries[new PropertyTag(taskDelegationStateTag)].GetIntegerValue();

                        item.taskDelegationState = EnumUtil.ParseTaskDelegationState((uint)taskDelegationStateValue);
                    }
                }
            }

            uint noteWidthNamedPropertyId = 0x8B02;

            if (reader.PstFile.NamedIdToTag.ContainsKey(noteWidthNamedPropertyId))
            {
                long noteWidthTag = reader.PstFile.NamedIdToTag[noteWidthNamedPropertyId];

                if (noteWidthTag > -1)
                {
                    noteWidthTag = (noteWidthTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(noteWidthTag)] != null)
                    {
                        item.noteWidth = table.Entries[new PropertyTag(noteWidthTag)].GetIntegerValue();
                    }
                }
            }

            uint noteHeightNamedPropertyId = 0x8B03;

            if (reader.PstFile.NamedIdToTag.ContainsKey(noteHeightNamedPropertyId))
            {
                long noteHeightTag = reader.PstFile.NamedIdToTag[noteHeightNamedPropertyId];

                if (noteHeightTag > -1)
                {
                    noteHeightTag = (noteHeightTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(noteHeightTag)] != null)
                    {
                        item.noteHeight = table.Entries[new PropertyTag(noteHeightTag)].GetIntegerValue();
                    }
                }
            }

            uint noteLeftNamedPropertyId = 0x8B04;

            if (reader.PstFile.NamedIdToTag.ContainsKey(noteLeftNamedPropertyId))
            {
                long noteLeftTag = reader.PstFile.NamedIdToTag[noteLeftNamedPropertyId];

                if (noteLeftTag > -1)
                {
                    noteLeftTag = (noteLeftTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(noteLeftTag)] != null)
                    {
                        item.noteLeft = table.Entries[new PropertyTag(noteLeftTag)].GetIntegerValue();
                    }
                }
            }

            uint noteTopNamedPropertyId = 0x8B05;

            if (reader.PstFile.NamedIdToTag.ContainsKey(noteTopNamedPropertyId))
            {
                long noteTopTag = reader.PstFile.NamedIdToTag[noteTopNamedPropertyId];

                if (noteTopTag > -1)
                {
                    noteTopTag = (noteTopTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(noteTopTag)] != null)
                    {
                        item.noteTop = table.Entries[new PropertyTag(noteTopTag)].GetIntegerValue();
                    }
                }
            }

            uint noteColorNamedPropertyId = 0x8B00;

            if (reader.PstFile.NamedIdToTag.ContainsKey(noteColorNamedPropertyId))
            {
                long noteColorTag = reader.PstFile.NamedIdToTag[noteColorNamedPropertyId];

                if (noteColorTag > -1)
                {
                    noteColorTag = (noteColorTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(noteColorTag)] != null)
                    {
                        int noteColorValue = table.Entries[new PropertyTag(noteColorTag)].GetIntegerValue();

                        item.noteColor = EnumUtil.ParseNoteColor((uint)noteColorValue);
                    }
                }
            }

            uint journalStartTimeNamedPropertyId = 0x8706;

            if (reader.PstFile.NamedIdToTag.ContainsKey(journalStartTimeNamedPropertyId))
            {
                long journalStartTimeTag = reader.PstFile.NamedIdToTag[journalStartTimeNamedPropertyId];

                if (journalStartTimeTag > -1)
                {
                    journalStartTimeTag = (journalStartTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(journalStartTimeTag)] != null)
                    {
                        item.journalStartTime = table.Entries[new PropertyTag(journalStartTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint journalEndTimeNamedPropertyId = 0x8708;

            if (reader.PstFile.NamedIdToTag.ContainsKey(journalEndTimeNamedPropertyId))
            {
                long journalEndTimeTag = reader.PstFile.NamedIdToTag[journalEndTimeNamedPropertyId];

                if (journalEndTimeTag > -1)
                {
                    journalEndTimeTag = (journalEndTimeTag << 16) + 0x0040;

                    if (table.Entries[new PropertyTag(journalEndTimeTag)] != null)
                    {
                        item.journalEndTime = table.Entries[new PropertyTag(journalEndTimeTag)].GetDateTimeValue();
                    }
                }
            }

            uint journalTypeNamedPropertyId = 0x8700;

            if (reader.PstFile.NamedIdToTag.ContainsKey(journalTypeNamedPropertyId))
            {
                long journalTypeTag = reader.PstFile.NamedIdToTag[journalTypeNamedPropertyId];

                if (journalTypeTag > -1)
                {
                    journalTypeTag = (journalTypeTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(journalTypeTag)] != null)
                    {
                        item.journalType = table.Entries[new PropertyTag(journalTypeTag)].GetStringValue();
                    }
                }
            }

            uint journalTypeDescriptionNamedPropertyId = 0x8712;

            if (reader.PstFile.NamedIdToTag.ContainsKey(journalTypeDescriptionNamedPropertyId))
            {
                long journalTypeDescriptionTag = reader.PstFile.NamedIdToTag[journalTypeDescriptionNamedPropertyId];

                if (journalTypeDescriptionTag > -1)
                {
                    journalTypeDescriptionTag = (journalTypeDescriptionTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(journalTypeDescriptionTag)] != null)
                    {
                        item.journalTypeDescription = table.Entries[new PropertyTag(journalTypeDescriptionTag)].GetStringValue();
                    }
                }
            }

            uint journalDurationNamedPropertyId = 0x8707;

            if (reader.PstFile.NamedIdToTag.ContainsKey(journalDurationNamedPropertyId))
            {
                long journalDurationTag = reader.PstFile.NamedIdToTag[journalDurationNamedPropertyId];

                if (journalDurationTag > -1)
                {
                    journalDurationTag = (journalDurationTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(journalDurationTag)] != null)
                    {
                        item.journalDuration = table.Entries[new PropertyTag(journalDurationTag)].GetIntegerValue();
                    }
                }
            }

            uint selectedMailingAddressNamedPropertyId = 0x8022;

            if (reader.PstFile.NamedIdToTag.ContainsKey(selectedMailingAddressNamedPropertyId))
            {
                long selectedMailingAddressTag = reader.PstFile.NamedIdToTag[selectedMailingAddressNamedPropertyId];

                if (selectedMailingAddressTag > -1)
                {
                    selectedMailingAddressTag = (selectedMailingAddressTag << 16) + 0x0003;

                    if (table.Entries[new PropertyTag(selectedMailingAddressTag)] != null)
                    {
                        int selectedMailingAddressValue = table.Entries[new PropertyTag(selectedMailingAddressTag)].GetIntegerValue();

                        item.selectedMailingAddress = EnumUtil.ParseSelectedMailingAddress((uint)selectedMailingAddressValue);
                    }
                }
            }

            uint contactHasPictureNamedPropertyId = 0x8015;

            if (reader.PstFile.NamedIdToTag.ContainsKey(contactHasPictureNamedPropertyId))
            {
                long contactHasPictureTag = reader.PstFile.NamedIdToTag[contactHasPictureNamedPropertyId];

                if (contactHasPictureTag > -1)
                {
                    contactHasPictureTag = (contactHasPictureTag << 16) + 0x000B;

                    if (table.Entries[new PropertyTag(contactHasPictureTag)] != null)
                    {
                        item.contactHasPicture = table.Entries[new PropertyTag(contactHasPictureTag)].GetBooleanValue();
                    }
                }
            }

            uint fileAsNamedPropertyId = 0x8005;

            if (reader.PstFile.NamedIdToTag.ContainsKey(fileAsNamedPropertyId))
            {
                long fileAsTag = reader.PstFile.NamedIdToTag[fileAsNamedPropertyId];

                if (fileAsTag > -1)
                {
                    fileAsTag = (fileAsTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(fileAsTag)] != null)
                    {
                        item.fileAs = table.Entries[new PropertyTag(fileAsTag)].GetStringValue();
                    }
                }
            }

            uint instantMessengerAddressNamedPropertyId = 0x8062;

            if (reader.PstFile.NamedIdToTag.ContainsKey(instantMessengerAddressNamedPropertyId))
            {
                long instantMessengerAddressTag = reader.PstFile.NamedIdToTag[instantMessengerAddressNamedPropertyId];

                if (instantMessengerAddressTag > -1)
                {
                    instantMessengerAddressTag = (instantMessengerAddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(instantMessengerAddressTag)] != null)
                    {
                        item.instantMessengerAddress = table.Entries[new PropertyTag(instantMessengerAddressTag)].GetStringValue();
                    }
                }
            }

            uint internetFreeBusyAddressNamedPropertyId = 0x80D8;

            if (reader.PstFile.NamedIdToTag.ContainsKey(internetFreeBusyAddressNamedPropertyId))
            {
                long internetFreeBusyAddressTag = reader.PstFile.NamedIdToTag[internetFreeBusyAddressNamedPropertyId];

                if (internetFreeBusyAddressTag > -1)
                {
                    internetFreeBusyAddressTag = (internetFreeBusyAddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(internetFreeBusyAddressTag)] != null)
                    {
                        item.internetFreeBusyAddress = table.Entries[new PropertyTag(internetFreeBusyAddressTag)].GetStringValue();
                    }
                }
            }

            uint businessAddressNamedPropertyId = 0x801B;

            if (reader.PstFile.NamedIdToTag.ContainsKey(businessAddressNamedPropertyId))
            {
                long businessAddressTag = reader.PstFile.NamedIdToTag[businessAddressNamedPropertyId];

                if (businessAddressTag > -1)
                {
                    businessAddressTag = (businessAddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(businessAddressTag)] != null)
                    {
                        item.businessAddress = table.Entries[new PropertyTag(businessAddressTag)].GetStringValue();
                    }
                }
            }

            uint homeAddressNamedPropertyId = 0x801A;

            if (reader.PstFile.NamedIdToTag.ContainsKey(homeAddressNamedPropertyId))
            {
                long homeAddressTag = reader.PstFile.NamedIdToTag[homeAddressNamedPropertyId];

                if (homeAddressTag > -1)
                {
                    homeAddressTag = (homeAddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(homeAddressTag)] != null)
                    {
                        item.homeAddress = table.Entries[new PropertyTag(homeAddressTag)].GetStringValue();
                    }
                }
            }

            uint otherAddressNamedPropertyId = 0x801C;

            if (reader.PstFile.NamedIdToTag.ContainsKey(otherAddressNamedPropertyId))
            {
                long otherAddressTag = reader.PstFile.NamedIdToTag[otherAddressNamedPropertyId];

                if (otherAddressTag > -1)
                {
                    otherAddressTag = (otherAddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(otherAddressTag)] != null)
                    {
                        item.otherAddress = table.Entries[new PropertyTag(otherAddressTag)].GetStringValue();
                    }
                }
            }

            uint email1AddressNamedPropertyId = 0x8083;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email1AddressNamedPropertyId))
            {
                long email1AddressTag = reader.PstFile.NamedIdToTag[email1AddressNamedPropertyId];

                if (email1AddressTag > -1)
                {
                    email1AddressTag = (email1AddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email1AddressTag)] != null)
                    {
                        item.email1Address = table.Entries[new PropertyTag(email1AddressTag)].GetStringValue();
                    }
                }
            }

            uint email2AddressNamedPropertyId = 0x8093;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email2AddressNamedPropertyId))
            {
                long email2AddressTag = reader.PstFile.NamedIdToTag[email2AddressNamedPropertyId];

                if (email2AddressTag > -1)
                {
                    email2AddressTag = (email2AddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email2AddressTag)] != null)
                    {
                        item.email2Address = table.Entries[new PropertyTag(email2AddressTag)].GetStringValue();
                    }
                }
            }

            uint email3AddressNamedPropertyId = 0x80A3;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email3AddressNamedPropertyId))
            {
                long email3AddressTag = reader.PstFile.NamedIdToTag[email3AddressNamedPropertyId];

                if (email3AddressTag > -1)
                {
                    email3AddressTag = (email3AddressTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email3AddressTag)] != null)
                    {
                        item.email3Address = table.Entries[new PropertyTag(email3AddressTag)].GetStringValue();
                    }
                }
            }

            uint email1DisplayNameNamedPropertyId = 0x8084;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email1DisplayNameNamedPropertyId))
            {
                long email1DisplayNameTag = reader.PstFile.NamedIdToTag[email1DisplayNameNamedPropertyId];

                if (email1DisplayNameTag > -1)
                {
                    email1DisplayNameTag = (email1DisplayNameTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email1DisplayNameTag)] != null)
                    {
                        item.email1DisplayName = table.Entries[new PropertyTag(email1DisplayNameTag)].GetStringValue();
                    }
                }
            }

            uint email2DisplayNameNamedPropertyId = 0x8094;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email2DisplayNameNamedPropertyId))
            {
                long email2DisplayNameTag = reader.PstFile.NamedIdToTag[email2DisplayNameNamedPropertyId];

                if (email2DisplayNameTag > -1)
                {
                    email2DisplayNameTag = (email2DisplayNameTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email2DisplayNameTag)] != null)
                    {
                        item.email2DisplayName = table.Entries[new PropertyTag(email2DisplayNameTag)].GetStringValue();
                    }
                }
            }

            uint email3DisplayNameNamedPropertyId = 0x80A4;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email3DisplayNameNamedPropertyId))
            {
                long email3DisplayNameTag = reader.PstFile.NamedIdToTag[email3DisplayNameNamedPropertyId];

                if (email3DisplayNameTag > -1)
                {
                    email3DisplayNameTag = (email3DisplayNameTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email3DisplayNameTag)] != null)
                    {
                        item.email3DisplayName = table.Entries[new PropertyTag(email3DisplayNameTag)].GetStringValue();
                    }
                }
            }

            uint email1DisplayAsNamedPropertyId = 0x8080;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email1DisplayAsNamedPropertyId))
            {
                long email1DisplayAsTag = reader.PstFile.NamedIdToTag[email1DisplayAsNamedPropertyId];

                if (email1DisplayAsTag > -1)
                {
                    email1DisplayAsTag = (email1DisplayAsTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email1DisplayAsTag)] != null)
                    {
                        item.email1DisplayAs = table.Entries[new PropertyTag(email1DisplayAsTag)].GetStringValue();
                    }
                }
            }

            uint email2DisplayAsNamedPropertyId = 0x8090;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email2DisplayAsNamedPropertyId))
            {
                long email2DisplayAsTag = reader.PstFile.NamedIdToTag[email2DisplayAsNamedPropertyId];

                if (email2DisplayAsTag > -1)
                {
                    email2DisplayAsTag = (email2DisplayAsTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email2DisplayAsTag)] != null)
                    {
                        item.email2DisplayAs = table.Entries[new PropertyTag(email2DisplayAsTag)].GetStringValue();
                    }
                }
            }

            uint email3DisplayAsNamedPropertyId = 0x80A0;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email3DisplayAsNamedPropertyId))
            {
                long email3DisplayAsTag = reader.PstFile.NamedIdToTag[email3DisplayAsNamedPropertyId];

                if (email3DisplayAsTag > -1)
                {
                    email3DisplayAsTag = (email3DisplayAsTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email3DisplayAsTag)] != null)
                    {
                        item.email3DisplayAs = table.Entries[new PropertyTag(email3DisplayAsTag)].GetStringValue();
                    }
                }
            }

            uint email1TypeNamedPropertyId = 0x8082;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email1TypeNamedPropertyId))
            {
                long email1TypeTag = reader.PstFile.NamedIdToTag[email1TypeNamedPropertyId];

                if (email1TypeTag > -1)
                {
                    email1TypeTag = (email1TypeTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email1TypeTag)] != null)
                    {
                        item.email1Type = table.Entries[new PropertyTag(email1TypeTag)].GetStringValue();
                    }
                }
            }

            uint email2TypeNamedPropertyId = 0x8092;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email2TypeNamedPropertyId))
            {
                long email2TypeTag = reader.PstFile.NamedIdToTag[email2TypeNamedPropertyId];

                if (email2TypeTag > -1)
                {
                    email2TypeTag = (email2TypeTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email2TypeTag)] != null)
                    {
                        item.email2Type = table.Entries[new PropertyTag(email2TypeTag)].GetStringValue();
                    }
                }
            }

            uint email3TypeNamedPropertyId = 0x80A2;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email3TypeNamedPropertyId))
            {
                long email3TypeTag = reader.PstFile.NamedIdToTag[email3TypeNamedPropertyId];

                if (email3TypeTag > -1)
                {
                    email3TypeTag = (email3TypeTag << 16) + 0x001E;

                    if (table.Entries[new PropertyTag(email3TypeTag)] != null)
                    {
                        item.email3Type = table.Entries[new PropertyTag(email3TypeTag)].GetStringValue();
                    }
                }
            }

            uint email1EntryIdNamedPropertyId = 0x8085;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email1EntryIdNamedPropertyId))
            {
                long email1EntryIdTag = reader.PstFile.NamedIdToTag[email1EntryIdNamedPropertyId];

                if (email1EntryIdTag > -1)
                {
                    email1EntryIdTag = (email1EntryIdTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(email1EntryIdTag)] != null)
                    {
                        item.email1EntryId = table.Entries[new PropertyTag(email1EntryIdTag)].GetBinaryValue();
                    }
                }
            }

            uint email2EntryIdNamedPropertyId = 0x8095;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email2EntryIdNamedPropertyId))
            {
                long email2EntryIdTag = reader.PstFile.NamedIdToTag[email2EntryIdNamedPropertyId];

                if (email2EntryIdTag > -1)
                {
                    email2EntryIdTag = (email2EntryIdTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(email2EntryIdTag)] != null)
                    {
                        item.email2EntryId = table.Entries[new PropertyTag(email2EntryIdTag)].GetBinaryValue();
                    }
                }
            }

            uint email3EntryIdNamedPropertyId = 0x80A5;

            if (reader.PstFile.NamedIdToTag.ContainsKey(email3EntryIdNamedPropertyId))
            {
                long email3EntryIdTag = reader.PstFile.NamedIdToTag[email3EntryIdNamedPropertyId];

                if (email3EntryIdTag > -1)
                {
                    email3EntryIdTag = (email3EntryIdTag << 16) + 0x0102;

                    if (table.Entries[new PropertyTag(email3EntryIdTag)] != null)
                    {
                        item.email3EntryId = table.Entries[new PropertyTag(email3EntryIdTag)].GetBinaryValue();
                    }
                }
            }

            string keywordsNamedPropertyName = "Keywords";
            byte[] keywordsGuid = new byte[16] { 41, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };

            string keywordsNamedPropertyHex = Util.ConvertNamedPropertyToHex(keywordsNamedPropertyName, keywordsGuid);

            if (reader.PstFile.NamedIdToTag.ContainsKey(keywordsNamedPropertyHex))
            {
                long keywordsTag = reader.PstFile.NamedIdToTag[keywordsNamedPropertyHex];

                if (table.GetEntry((int)keywordsTag) != null)
                {
                    item.keywords = table.GetEntry((int)keywordsTag).GetStringArrayValue();
                }
            }

            if (reader.PstFile.NameToIdMap != null)
            {
              //Extended properties
              foreach (uint key in reader.PstFile.NameToIdMap.Map.Keys)
              {
                  NamedProperty currentNamedProperty = reader.PstFile.NameToIdMap.Map[key];

                  if (currentNamedProperty != null)
                  {
                      object currentNamedPropertyHex = null;
                      ExtendedProperty extendedProperty = new ExtendedProperty(reader.PstFile.Encoding);

                      if (currentNamedProperty.Name == null && currentNamedProperty.Id > 0)
                      {
                          currentNamedPropertyHex = currentNamedProperty.Id;

                          ExtendedPropertyId extendedPropertyId = new ExtendedPropertyId((int)currentNamedProperty.Id, currentNamedProperty.Guid);
                          extendedProperty.Tag = extendedPropertyId;
                      }
                      else if (currentNamedProperty.Name != null)
                      {
                          currentNamedPropertyHex = Util.ConvertNamedPropertyToHex(currentNamedProperty.Name, currentNamedProperty.Guid);

                          ExtendedPropertyName extendedPropertyName = new ExtendedPropertyName(currentNamedProperty.Name, currentNamedProperty.Guid);
                          extendedProperty.Tag = extendedPropertyName;
                      }

                      if (currentNamedPropertyHex != null)
                      {
                    	  if (reader.PstFile.NamedIdToTag.ContainsKey(currentNamedPropertyHex))
                          {
                              long currentTag = reader.PstFile.NamedIdToTag[currentNamedPropertyHex];

                              if (table.GetEntry((int)currentTag) != null)
                              {
                                  TableEntry tableEntry = table.GetEntry((int)currentTag);

                                  extendedProperty.Tag.Type = tableEntry.PropertyTag.Type;
                                  extendedProperty.Value = tableEntry.ValueBuffer;

                                  item.ExtendedProperties.Add(extendedProperty);

                              }
                          }
                      }
                  }
              }
            }

            Table recipientsTable = Util.GetRecipientsTable(reader, localDescriptorList);

            if (recipientsTable != null)
            {
                for (int i = 0; i < recipientsTable.EntriesArray.Length; i++)
                {
                    TableEntryList recipientEntries = recipientsTable.EntriesArray[i];

                    Recipient recipient = new Recipient();

                    if (recipientEntries[MapiPropertyTag.PR_DISPLAY_TYPE] != null)
                    {
                        int displayTypeValue = recipientEntries[MapiPropertyTag.PR_DISPLAY_TYPE].GetIntegerValue();
                        recipient.displayType = EnumUtil.ParseDisplayType((uint)displayTypeValue);
                    }

                    if (recipientEntries[MapiPropertyTag.PR_OBJECT_TYPE] != null)
                    {
                        int objectTypeValue = recipientEntries[MapiPropertyTag.PR_OBJECT_TYPE].GetIntegerValue();
                        recipient.objectType = EnumUtil.ParseObjectType((uint)objectTypeValue);
                    }

                    if (recipientEntries[MapiPropertyTag.PR_RECIPIENT_TYPE] != null)
                    {
                        int recipientTypeValue = recipientEntries[MapiPropertyTag.PR_RECIPIENT_TYPE].GetIntegerValue();
                        recipient.recipientType = EnumUtil.ParseRecipientType((uint)recipientTypeValue);
                    }

                    if (recipientEntries[MapiPropertyTag.PR_DISPLAY_NAME] != null)
                    {
                        recipient.displayName = recipientEntries[MapiPropertyTag.PR_DISPLAY_NAME].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_ADDRTYPE] != null)
                    {
                        recipient.addressType = recipientEntries[MapiPropertyTag.PR_ADDRTYPE].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_EMAIL_ADDRESS] != null)
                    {
                        recipient.emailAddress = recipientEntries[MapiPropertyTag.PR_EMAIL_ADDRESS].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_ENTRYID] != null)
                    {
                        recipient.entryId = recipientEntries[MapiPropertyTag.PR_ENTRYID].GetBinaryValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_SEARCH_KEY] != null)
                    {
                        recipient.searchKey = recipientEntries[MapiPropertyTag.PR_SEARCH_KEY].GetBinaryValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_INSTANCE_KEY] != null)
                    {
                        recipient.instanceKey = recipientEntries[MapiPropertyTag.PR_INSTANCE_KEY].GetBinaryValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_SMTP_ADDRESS] != null)
                    {
                        recipient.smtpAddress = recipientEntries[MapiPropertyTag.PR_SMTP_ADDRESS].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_7BIT_DISPLAY_NAME] != null)
                    {
                        recipient.displayName7Bit = recipientEntries[MapiPropertyTag.PR_7BIT_DISPLAY_NAME].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_TRANSMITABLE_DISPLAY_NAME] != null)
                    {
                        recipient.transmitableDisplayName = recipientEntries[MapiPropertyTag.PR_TRANSMITABLE_DISPLAY_NAME].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_ORG_ADDR_TYPE] != null)
                    {
                        recipient.originatingAddressType = recipientEntries[MapiPropertyTag.PR_ORG_ADDR_TYPE].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_ORG_EMAIL_ADDR] != null)
                    {
                        recipient.originatingEmailAddress = recipientEntries[MapiPropertyTag.PR_ORG_EMAIL_ADDR].GetStringValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_RESPONSIBILITY] != null)
                    {
                        recipient.responsibility = recipientEntries[MapiPropertyTag.PR_RESPONSIBILITY].GetBooleanValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_SEND_RICH_INFO] != null)
                    {
                        recipient.sendRichInfo = recipientEntries[MapiPropertyTag.PR_SEND_RICH_INFO].GetBooleanValue();
                    }

                    if (recipientEntries[MapiPropertyTag.PR_SEND_INTERNET_ENCODING] != null)
                    {
                        recipient.sendInternetEncoding = recipientEntries[MapiPropertyTag.PR_SEND_INTERNET_ENCODING].GetIntegerValue();
                    }

                    item.Recipients.Add(recipient);
                }
            }

            if (includeAttachments)
            {
                IList<Attachment> attachments = GetAttachments(reader, localDescriptorList);

                foreach(Attachment attachment in attachments)
                {
                    item.Attachments.Add(attachment);
                }
            }

            return item;
        }

        internal static IList<Attachment> GetAttachments(PstFileReader reader, LocalDescriptorList localDescriptorList)
        {
            Table attachmentsTable = Util.GetAttachmentsTable(reader, localDescriptorList);
            IList<Attachment> attachments = new List<Attachment>();
            long attachmentIdTag = 0x67F20003;

            if (attachmentsTable != null)
            {
                for (int i = 0; i < attachmentsTable.EntriesArray.Length; i++)
                {
                    TableEntryList attachmentEntries = attachmentsTable.EntriesArray[i];
                    Attachment attachment = GetAttachment(attachmentEntries);

                    DataStructure attachmentDataNode = null;
                    DataStructure subListLocalDescriptorListNode = null;
                    ulong dataStructureId = 0;

                    if (attachmentEntries[new PropertyTag(attachmentIdTag)] != null)
                    {
                        uint attachmentId = (uint)attachmentEntries[new PropertyTag(attachmentIdTag)].GetIntegerValue();

                        if (attachmentsTable.SLEntriesTable.Count > 1)
                        {
                            SLEntry slEntry = attachmentsTable.SLEntriesTable.ContainsKey(attachmentId) ? attachmentsTable.SLEntriesTable[attachmentId] : null;

                            if (slEntry != null)
                            {
                                dataStructureId = slEntry.BlockId;

                                attachmentDataNode = reader.PstFile.DataIndexTree.GetDataStructure(dataStructureId);
                                subListLocalDescriptorListNode = GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, slEntry.SubNodeBlockId);
                            }
                        }
                        else if (localDescriptorList != null)
                        {
                            LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(attachmentId) ? localDescriptorList.Elements[attachmentId] : null;

                            if (element != null)
                            {
                                dataStructureId = element.DataStructureId;

                                attachmentDataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);
                                subListLocalDescriptorListNode = GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, element.SubListId);
                            }
                        }

                        if (subListLocalDescriptorListNode == null)
                        {
                            if (attachmentDataNode != null)
                            {
                                Table attachmentTable = null;

                                byte[] attachmentBuffer = Util.GetBuffer(reader, attachmentDataNode);

                                if (attachmentBuffer != null && attachmentBuffer.Length > 0)
                                {
                                    ushort tableType = BitConverter.ToUInt16(attachmentBuffer, 2);

                                    if (tableType == 0xBCEC)
                                    {
                                        attachmentTable = new TableBC(reader, attachmentBuffer, localDescriptorList, attachmentDataNode);
                                    }
                                    else if (tableType == 0x7CEC)
                                    {
                                        attachmentTable = new Table7C(reader, attachmentBuffer, localDescriptorList, attachmentDataNode);
                                    }
                                    else if (tableType == 0xACEC)
                                    {
                                        attachmentTable = new TableAC(reader, attachmentBuffer, localDescriptorList, attachmentDataNode);
                                    }
                                    else if (tableType == 0x9CEC)
                                    {
                                        attachmentTable = new Table9C(reader, attachmentBuffer, localDescriptorList, attachmentDataNode);
                                    }

                                    if (attachmentTable != null)
                                    {
                                        attachment = GetAttachment(attachmentTable.Entries);
                                    }
                                    else
                                    {
                                        attachment.data = attachmentBuffer;
                                    }
                                }
                            }
                        }
                        else
                        {
                            LocalDescriptorList subListLocalDescriptorList = GetLocalDescriptorList(reader, subListLocalDescriptorListNode);
                            DataStructure subListDataNode = reader.PstFile.DataIndexTree.GetDataStructure(dataStructureId);

                            if (subListDataNode != null)
                            {
                                Table attachmentTable = null;

                                byte[] attachmentBuffer = Util.GetBuffer(reader, subListDataNode);

                                if (attachmentBuffer != null && attachmentBuffer.Length > 0)
                                {
                                    ushort tableType = BitConverter.ToUInt16(attachmentBuffer, 2);

                                    if (tableType == 0xBCEC)
                                    {
                                        attachmentTable = new TableBC(reader, attachmentBuffer, subListLocalDescriptorList, subListDataNode);
                                    }
                                    else if (tableType == 0x7CEC)
                                    {
                                        attachmentTable = new Table7C(reader, attachmentBuffer, subListLocalDescriptorList, subListDataNode);
                                    }
                                    else if (tableType == 0xACEC)
                                    {
                                        attachmentTable = new TableAC(reader, attachmentBuffer, subListLocalDescriptorList, subListDataNode);
                                    }
                                    else if (tableType == 0x9CEC)
                                    {
                                        attachmentTable = new Table9C(reader, attachmentBuffer, subListLocalDescriptorList, subListDataNode);
                                    }

                                    if (attachmentTable != null && attachment.Method == AttachmentMethod.EmbeddedMessage)
                                    {
                                        attachment = GetAttachment(attachmentTable.Entries);

                                        LocalDescriptorListElement subListElement = subListLocalDescriptorList.GetFirstElement();

                                        if (subListElement != null)
                                        {
                                            subListLocalDescriptorList = subListElement.SubList;
                                        }

                                        if (attachment.DataObject != null)
                                        {
                                            Table embeddedItemTable = null;
                                            ushort embeddedItemTableType = BitConverter.ToUInt16(attachment.DataObject, 2);

                                            if (embeddedItemTableType == 0xBCEC)
                                            {
                                                embeddedItemTable = new TableBC(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                            }
                                            else if (embeddedItemTableType == 0x7CEC)
                                            {
                                                embeddedItemTable = new Table7C(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                            }
                                            else if (embeddedItemTableType == 0xACEC)
                                            {
                                                embeddedItemTable = new TableAC(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                            }
                                            else if (embeddedItemTableType == 0x9CEC)
                                            {
                                                embeddedItemTable = new Table9C(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                            }

                                            if (embeddedItemTable != null)
                                            {
                                                Item embeddedItem = GetItem(reader, subListLocalDescriptorList, embeddedItemTable, subListDataNode.Id, 0);

                                                if (embeddedItem.rtfCompressed != null && !embeddedItem.RtfInSync)
                                                {
                                                    embeddedItem.rtfInSync = true;
                                                }

                                                attachment.embeddedItem = embeddedItem;
                                            }
                                        }

                                    }
                                    else if (attachmentTable != null)
                                    {
                                        attachment = GetAttachment(attachmentTable.Entries);

                                        if (attachment.Method == AttachmentMethod.EmbeddedMessage)
                                        {
                                            LocalDescriptorListElement subListElement = subListLocalDescriptorList.GetFirstElement();

                                            if (subListElement != null)
                                            {
                                                subListLocalDescriptorList = subListElement.SubList;
                                            }

                                            if (attachment.DataObject != null)
                                            {
                                                Table embeddedItemTable = null;
                                                ushort embeddedItemTableType = BitConverter.ToUInt16(attachment.DataObject, 2);

                                                if (embeddedItemTableType == 0xBCEC)
                                                {
                                                    embeddedItemTable = new TableBC(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                                }
                                                else if (embeddedItemTableType == 0x7CEC)
                                                {
                                                    embeddedItemTable = new Table7C(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                                }
                                                else if (embeddedItemTableType == 0xACEC)
                                                {
                                                    embeddedItemTable = new TableAC(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                                }
                                                else if (embeddedItemTableType == 0x9CEC)
                                                {
                                                    embeddedItemTable = new Table9C(reader, attachment.DataObject, subListLocalDescriptorList, subListDataNode);
                                                }

                                                if (embeddedItemTable != null)
                                                {
                                                    Item embeddedItem = GetItem(reader, subListLocalDescriptorList, embeddedItemTable, subListDataNode.Id, 0);

                                                    if (embeddedItem.rtfCompressed != null && !embeddedItem.RtfInSync)
                                                    {
                                                        embeddedItem.rtfInSync = true;
                                                    }

                                                    attachment.embeddedItem = embeddedItem;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        attachment.data = attachmentBuffer;
                                    }
                                }
                            }
                        }
                    }

                    attachments.Add(attachment);
                }
            }

            return attachments;
        }

        internal static Attachment GetAttachment(TableEntryList attachmentEntries)
        {
            Attachment attachment = new Attachment();

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_SIZE] != null)
            {
                attachment.size = attachmentEntries[MapiPropertyTag.PR_ATTACH_SIZE].GetIntegerValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_FLAGS] != null)
            {
                int attachmentFlagsValue = attachmentEntries[MapiPropertyTag.PR_ATTACH_FLAGS].GetIntegerValue();
                attachment.flags = EnumUtil.ParseAttachmentFlags((uint)attachmentFlagsValue);
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_METHOD] != null)
            {
                int attachmentMethodValue = attachmentEntries[MapiPropertyTag.PR_ATTACH_METHOD].GetIntegerValue();
                attachment.method = EnumUtil.ParseAttachmentMethod((uint)attachmentMethodValue);
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_MIME_SEQUENCE] != null)
            {
                attachment.mimeSequence = attachmentEntries[MapiPropertyTag.PR_ATTACH_MIME_SEQUENCE].GetIntegerValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_RENDERING_POSITION] != null)
            {
                attachment.renderingPosition = attachmentEntries[MapiPropertyTag.PR_RENDERING_POSITION].GetIntegerValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_OBJECT_TYPE] != null)
            {
                int objectTypeValue = attachmentEntries[MapiPropertyTag.PR_OBJECT_TYPE].GetIntegerValue();
                attachment.objectType = EnumUtil.ParseObjectType((uint)objectTypeValue);
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACHMENT_HIDDEN] != null)
            {
                int intValue = attachmentEntries[MapiPropertyTag.PR_ATTACHMENT_HIDDEN].GetIntegerValue();

                if (intValue == 0)
                {
                    attachment.isHidden = false;
                }
                else
                {
                    attachment.isHidden = attachmentEntries[MapiPropertyTag.PR_ATTACHMENT_HIDDEN].GetBooleanValue();
                }
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACHMENT_CONTACTPHOTO] != null)
            {
                attachment.isContactPhoto = attachmentEntries[MapiPropertyTag.PR_ATTACHMENT_CONTACTPHOTO].GetBooleanValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_ADDITIONAL_INFO] != null)
            {
                attachment.additionalInfo = attachmentEntries[MapiPropertyTag.PR_ATTACH_ADDITIONAL_INFO].GetBinaryValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_BASE] != null)
            {
                attachment.contentBase = attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_BASE].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_ID] != null)
            {
                attachment.contentId = attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_ID].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_LOCATION] != null)
            {
                attachment.contentLocation = attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_LOCATION].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_DISPOSITION] != null)
            {
                attachment.contentDisposition = attachmentEntries[MapiPropertyTag.PR_ATTACH_CONTENT_DISPOSITION].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_ENCODING] != null)
            {
                attachment.encoding = attachmentEntries[MapiPropertyTag.PR_ATTACH_ENCODING].GetBinaryValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_RECORD_KEY] != null)
            {
                attachment.recordKey = attachmentEntries[MapiPropertyTag.PR_RECORD_KEY].GetBinaryValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_EXTENSION] != null)
            {
                attachment.extension = attachmentEntries[MapiPropertyTag.PR_ATTACH_EXTENSION].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_FILENAME] != null)
            {
                attachment.fileName = attachmentEntries[MapiPropertyTag.PR_ATTACH_FILENAME].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_LONG_FILENAME] != null)
            {
                attachment.longFileName = attachmentEntries[MapiPropertyTag.PR_ATTACH_LONG_FILENAME].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_LONG_PATHNAME] != null)
            {
                attachment.longPathName = attachmentEntries[MapiPropertyTag.PR_ATTACH_LONG_PATHNAME].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_MIME_TAG] != null)
            {
                attachment.mimeTag = attachmentEntries[MapiPropertyTag.PR_ATTACH_MIME_TAG].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_PATHNAME] != null)
            {
                attachment.pathName = attachmentEntries[MapiPropertyTag.PR_ATTACH_PATHNAME].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_RENDERING] != null)
            {
                attachment.rendering = attachmentEntries[MapiPropertyTag.PR_ATTACH_RENDERING].GetBinaryValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_TAG] != null)
            {
                attachment.tag = attachmentEntries[MapiPropertyTag.PR_ATTACH_TAG].GetBinaryValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_ATTACH_TRANSPORT_NAME] != null)
            {
                attachment.transportName = attachmentEntries[MapiPropertyTag.PR_ATTACH_TRANSPORT_NAME].GetStringValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_DISPLAY_NAME] != null)
            {
                attachment.displayName = attachmentEntries[MapiPropertyTag.PR_DISPLAY_NAME].GetStringValue();
            }

            if (attachment.size <= PstFile.AttachmentSizeLimit)
            {
                if (attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_BIN.Tag, PropertyType.Binary] != null)
                {
                    attachment.Data = attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_BIN.Tag, PropertyType.Binary].GetBinaryValue();
                }

                if (attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_OBJ.Tag, PropertyType.Object] != null && attachment.method == AttachmentMethod.Ole)
                {
                    attachment.Data = attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_OBJ.Tag, PropertyType.Object].GetBinaryValue();
                }
                else if (attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_OBJ.Tag, PropertyType.Object] != null)
                {
                    attachment.DataObject = attachmentEntries[MapiPropertyTag.PR_ATTACH_DATA_OBJ.Tag, PropertyType.Object].GetBinaryValue();
                }
            }

            if (attachmentEntries[MapiPropertyTag.PR_CREATION_TIME] != null)
            {
                attachment.creationTime = attachmentEntries[MapiPropertyTag.PR_CREATION_TIME].GetDateTimeValue();
            }

            if (attachmentEntries[MapiPropertyTag.PR_LAST_MODIFICATION_TIME] != null)
            {
                attachment.lastModificationTime = attachmentEntries[MapiPropertyTag.PR_LAST_MODIFICATION_TIME].GetDateTimeValue();
            }

            return attachment;
        }

        internal static Table GetTable(PstFileReader reader, DataStructure dataNode, LocalDescriptorList localDescriptorList)
        {
            Table table = null;

            byte[] tableBuffer = null;

            try
            {
              tableBuffer = Util.GetBuffer(reader, dataNode);
            }
            catch (Exception ex)
            {
              Console.WriteLine("Continuing after Table Error {0}", ex);
            }

            if (tableBuffer != null && tableBuffer.Length > 0)
            {
                //FileStream file = new FileStream("c:\\testfolder\\lastTable.bin", FileMode.Create);
                //file.Write(tableBuffer, 0, tableBuffer.Length);
                //file.Close();

                //ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);

                //if (tableType == 0xBCEC)
                //{
                //    table = new TableBC(reader, tableBuffer, localDescriptorList, dataNode);
                //}
                //else if (tableType == 0x7CEC)
                //{
                //    table = new Table7C(reader, tableBuffer, localDescriptorList, dataNode);
                //}
                //else if (tableType == 0xACEC)
                //{
                //    table = new TableAC(reader, tableBuffer, localDescriptorList, dataNode);
                //}
                //else if (tableType == 0x9CEC)
                //{
                //    table = new Table9C(reader, tableBuffer, localDescriptorList, dataNode);
                //}

              if (tableBuffer[2] == 0xEC)
              {
                switch (tableBuffer[3])
                {
                  case 0xbc:  // Property Context (PC/BTH)
                    table = new TableBC(reader, tableBuffer, localDescriptorList, dataNode);
                    break;
                  case 0x7c:  //Table Context (TC/HN)
                    table = new Table7C(reader, tableBuffer, localDescriptorList, dataNode);
                    break;
                  case 0xac:  // Reserved
                    table = new TableAC(reader, tableBuffer, localDescriptorList, dataNode);
                    break;
                  case 0x9c:  // Reserved
                    table = new Table9C(reader, tableBuffer, localDescriptorList, dataNode);
                    break;
                  case 0xcc:  // Reserved
                    table = new TableCC(reader, tableBuffer);
                    break;
                  case 0xb5:  // BTree-on-Heap (BTH)
                    break;
                  default:
                    break;
                }
              }
            }

            return table;
        }

        internal static Table GetAttachmentsTable(PstFileReader reader, LocalDescriptorList localDescriptorList)
        {
            Table table = null;

            if (localDescriptorList != null)
            {
                 LocalDescriptorListElement attachmentElement = localDescriptorList.Elements.ContainsKey((uint)1649) ? localDescriptorList.Elements[(uint)1649] : null;

                if (attachmentElement != null)
                {
                    DataStructure attachmentDataNode = reader.PstFile.DataIndexTree.GetDataStructure(attachmentElement.DataStructureId);

                    if (attachmentDataNode != null)
                    {
                        byte[] tableBuffer = Util.GetBuffer(reader, attachmentDataNode);

                        if (tableBuffer != null && tableBuffer.Length > 0)
                        {
                            //FileStream file = new FileStream("c:\\testfolder\\lastTable.bin", FileMode.Create);
                            //file.Write(tableBuffer, 0, tableBuffer.Length);
                            //file.Close();

                            ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);

                            if (tableType == 0xBCEC)
                            {
                                table = new TableBC(reader, tableBuffer, localDescriptorList, attachmentDataNode);
                            }
                            else if (tableType == 0x7CEC)
                            {
                                table = new Table7C(reader, tableBuffer, localDescriptorList, attachmentDataNode);
                            }
                            else if (tableType == 0xACEC)
                            {
                                table = new TableAC(reader, tableBuffer, localDescriptorList, attachmentDataNode);
                            }
                            else if (tableType == 0x9CEC)
                            {
                                table = new Table9C(reader, tableBuffer, localDescriptorList, attachmentDataNode);
                            }
                            else if (tableBuffer[0] == 0x02 && tableBuffer[1] == 0x00)
                            {
                                SLEntry[] rgEntries = GetSLEntries(reader, tableBuffer);

                                if (rgEntries.Length > 0)
                                {
                                    DataStructure nodeItem = reader.PstFile.DataIndexTree.GetDataStructure(rgEntries[0].BlockId);
                                    DataStructure subListLocalDescriptorListNode = GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, rgEntries[0].SubNodeBlockId);

                                    LocalDescriptorList subLocalDescriptorList = GetLocalDescriptorList(reader, subListLocalDescriptorListNode);

                                    if (nodeItem != null)
                                    {
                                        reader.BaseStream.Position = (long)nodeItem.Offset;
                                        byte[] nodeBuffer = reader.ReadBytes(nodeItem.Size);

                                        byte[] slTableBuffer = Decrypt(reader.PstFile, nodeBuffer, nodeItem.Id);

                                        if (slTableBuffer != null && slTableBuffer.Length > 0)
                                        {
                                            tableType = BitConverter.ToUInt16(slTableBuffer, 2);

                                            if (tableType == 0xBCEC)
                                            {
                                                table = new TableBC(reader, slTableBuffer, subLocalDescriptorList, nodeItem);
                                            }
                                            else if (tableType == 0x7CEC)
                                            {
                                                table = new Table7C(reader, slTableBuffer, subLocalDescriptorList, nodeItem);
                                            }
                                            else if (tableType == 0xACEC)
                                            {
                                                table = new TableAC(reader, slTableBuffer, subLocalDescriptorList, nodeItem);
                                            }
                                            else if (tableType == 0x9CEC)
                                            {
                                                table = new Table9C(reader, slTableBuffer, subLocalDescriptorList, nodeItem);
                                            }
                                        }

                                        if (table != null)
                                        {
                                            for (int t = 0; t < rgEntries.Length; t++)
                                            {
                                                table.SLEntriesTable.Add((uint)rgEntries[t].NodeId, rgEntries[t]);
                                            }
                                        }
                                    }

                                    //Check if there are more SLEntry elements
                                    foreach (uint key in localDescriptorList.Elements.Keys)
                                    {
                                        if (key != 1649)
                                        {
                                            attachmentElement = localDescriptorList.Elements.ContainsKey(key) ? localDescriptorList.Elements[key] : null;

                                            if (attachmentElement != null)
                                            {
                                                attachmentDataNode = reader.PstFile.DataIndexTree.GetDataStructure(attachmentElement.DataStructureId);

                                                if (attachmentDataNode != null)
                                                {
                                                    tableBuffer = Util.GetBuffer(reader, attachmentDataNode);

                                                    if (tableBuffer != null && tableBuffer.Length > 0)
                                                    {
                                                        if (tableBuffer[0] == 0x02 && tableBuffer[1] == 0x00)
                                                        {
                                                            rgEntries = GetSLEntries(reader, tableBuffer);

                                                            if (table != null)
                                                            {
                                                                for (int t = 0; t < rgEntries.Length; t++)
                                                                {
                                                                    table.SLEntriesTable.Add((uint)rgEntries[t].NodeId, rgEntries[t]);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return table;
        }

        private static SLEntry[] GetSLEntries(PstFileReader reader, byte[] buffer)
        {
            ushort cEnt = BitConverter.ToUInt16(buffer, 2);

            SLEntry[] rgEntries = new SLEntry[cEnt];

            if (reader.PstFile.Is64Bit)
            {
                uint dwPadding = BitConverter.ToUInt32(buffer, 4);

                for (int i = 0; i < cEnt; i++)
                {
                    byte[] entryBuffer = new byte[24];
                    System.Array.Copy(buffer, 8 + i * 24, entryBuffer, 0, 24);

                    rgEntries[i] = new SLEntry(entryBuffer);
                }
            }
            else
            {
                for (int i = 0; i < cEnt; i++)
                {
                    byte[] entryBuffer = new byte[12];
                    System.Array.Copy(buffer, 4 + i * 12, entryBuffer, 0, 12);

                    rgEntries[i] = new SLEntry(entryBuffer);
                }
            }

            return rgEntries;
        }

        internal static Table GetRecipientsTable(PstFileReader reader, LocalDescriptorList localDescriptorList)
        {
            Table table = null;

            if (localDescriptorList != null)
            {
                LocalDescriptorListElement recipientElement = localDescriptorList.Elements.ContainsKey((uint)1682) ? localDescriptorList.Elements[(uint)1682] : null;

                DataStructure recipientDataNode = null;

                if (recipientElement != null)
                {
                    recipientDataNode = reader.PstFile.DataIndexTree.GetDataStructure(recipientElement.DataStructureId);
                }
                else if (recipientElement == null)
                {
                    Table attachmentTable = GetAttachmentsTable(reader, localDescriptorList);

                    if (attachmentTable != null && attachmentTable.SLEntriesTable.Count > 0)
                    {
                        SLEntry slEntry = attachmentTable.SLEntriesTable.ContainsKey((uint)1682) ? attachmentTable.SLEntriesTable[(uint)1682] : null;

                        if (slEntry != null)
                        {
                            recipientDataNode = reader.PstFile.DataIndexTree.GetDataStructure(slEntry.BlockId);
                        }
                    }
                }

                if (recipientDataNode != null)
                {
                    byte[] tableBuffer = Util.GetBuffer(reader, recipientDataNode);

                    if (tableBuffer != null && tableBuffer.Length > 0)
                    {
                        //FileStream file = new FileStream("c:\\testfolder\\lastTable.bin", FileMode.Create);
                        //file.Write(tableBuffer, 0, tableBuffer.Length);
                        //file.Close();

                        ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);

                        if (tableType == 0xBCEC)
                        {
                            table = new TableBC(reader, tableBuffer, localDescriptorList, recipientDataNode);
                        }
                        else if (tableType == 0x7CEC)
                        {
                            table = new Table7C(reader, tableBuffer, localDescriptorList, recipientDataNode, true);
                        }
                        else if (tableType == 0xACEC)
                        {
                            table = new TableAC(reader, tableBuffer, localDescriptorList, recipientDataNode);
                        }
                        else if (tableType == 0x9CEC)
                        {
                            table = new Table9C(reader, tableBuffer, localDescriptorList, recipientDataNode);
                        }
                    }
                }
            }

            return table;
        }

        internal static LocalDescriptorList GetLocalDescriptorList(PstFileReader reader, DataStructure localDescriptorListNode)
        {
            LocalDescriptorList localDescriptorList = null;

            if (localDescriptorListNode != null)
            {
                reader.BaseStream.Position = (long)localDescriptorListNode.Offset;
                byte[] localDescriptorListBuffer = reader.ReadBytes(localDescriptorListNode.Size);

                //FileStream file = new FileStream("c:\\testfolder\\lastList.bin", FileMode.Create);
                //file.Write(localDescriptorListBuffer, 0, localDescriptorListBuffer.Length);
                //file.Close();

                localDescriptorList = new LocalDescriptorList();

                localDescriptorList.Type = localDescriptorListBuffer[0];
                localDescriptorList.Level = localDescriptorListBuffer[1];
                localDescriptorList.ElementCount = BitConverter.ToUInt16(localDescriptorListBuffer, 2);

                if (reader.PstFile.Is64Bit)
                {
                    reader.ReadUInt32();
                }

                int nodeSize = 12;

                if (localDescriptorList.Level == 0 && !reader.PstFile.Is64Bit)
                {
                    nodeSize = 12;
                }
                else if (localDescriptorList.Level == 0 && reader.PstFile.Is64Bit)
                {
                    nodeSize = 24;
                }
                else if (localDescriptorList.Level == 1 && !reader.PstFile.Is64Bit)
                {
                    nodeSize = 8;
                }
                else if (localDescriptorList.Level == 1 && reader.PstFile.Is64Bit)
                {
                    nodeSize = 16;
                }

                for (int i = 0; i < localDescriptorList.ElementCount; i++)
                {
                    LocalDescriptorListElement element = new LocalDescriptorListElement();

                    if (!reader.PstFile.Is64Bit)
                    {
                        element.Id = BitConverter.ToUInt32(localDescriptorListBuffer, i * nodeSize + 4);
                        element.DataStructureId = BitConverter.ToUInt32(localDescriptorListBuffer, i * nodeSize + 8);

                        if (localDescriptorList.Level == 0)
                        {
                            element.SubListId = BitConverter.ToUInt32(localDescriptorListBuffer, i * nodeSize + 12);

                            if (element.SubListId > 0)
                            {
                                DataStructure subNode = GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, element.SubListId);
                                element.SubList = GetLocalDescriptorList(reader, subNode);
                            }
                        }
                    }
                    else
                    {
                        element.Id = BitConverter.ToUInt32(localDescriptorListBuffer, i * nodeSize + 8);

                        //skip unknown
                        reader.ReadUInt32();

                        element.DataStructureId = BitConverter.ToUInt64(localDescriptorListBuffer, i * nodeSize + 16);

                        if (localDescriptorList.Level == 0)
                        {
                            element.SubListId = BitConverter.ToUInt64(localDescriptorListBuffer, i * nodeSize + 24);

                            if (element.SubListId > 0)
                            {
                                DataStructure subNode = GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, element.SubListId);
                                element.SubList = GetLocalDescriptorList(reader, subNode);
                            }
                        }
                    }

                    localDescriptorList.Elements.Add(element.Id, element);
                }
            }

            return localDescriptorList;
        }

        internal static DataStructure GetLocalDescriptionListNode(IndexTree dataIndexTree, ulong listId)
        {
            listId = (listId >> 1) << 1;

            DataStructure node = dataIndexTree.Nodes.ContainsKey(listId) ? (DataStructure)dataIndexTree.Nodes[listId] : null;

            return node;
        }

        internal static byte[] GetBuffer(PstFileReader reader, DataStructure dataNode)
        {
            reader.BaseStream.Position = (long)dataNode.Offset;
            byte[] buffer = reader.ReadBytes(dataNode.Size);

            if (buffer[0] == 0x01 && (buffer[1] == 0x01 || buffer[1] == 0x02))
            {
                bool hasInnerArray = false;

                if (buffer[1] == 0x02)
                {
                    hasInnerArray = true;
                }

                Array array = new Array(buffer, reader.PstFile.Is64Bit);

                if (array.IsValid)
                {
                    MemoryStream memoryStream = new MemoryStream();

                    for (int i = 0; i < array.Items.Length; i++)
                    {
                        DataStructure arrayLeafNodeItem = reader.PstFile.DataIndexTree.GetDataStructure(array.Items[i]);

                        if (arrayLeafNodeItem != null)
                        {
                            reader.BaseStream.Position = (long)arrayLeafNodeItem.Offset;
                            byte[] arrayLeafNodeBuffer = reader.ReadBytes(arrayLeafNodeItem.Size);

                            if (hasInnerArray && arrayLeafNodeBuffer[0] == 0x01 && arrayLeafNodeBuffer[1] == 0x01)
                            {
                                Array innerArray = new Array(arrayLeafNodeBuffer, reader.PstFile.Is64Bit);

                                MemoryStream innerMemoryStream = new MemoryStream();

                                for (int j = 0; j < innerArray.Items.Length; j++)
                                {
                                    DataStructure innerArrayLeafNodeItem = reader.PstFile.DataIndexTree.GetDataStructure(innerArray.Items[j]);

                                    if (innerArrayLeafNodeItem != null)
                                    {
                                        reader.BaseStream.Position = (long)innerArrayLeafNodeItem.Offset;
                                        byte[] innerArrayLeafNodeBuffer = reader.ReadBytes(innerArrayLeafNodeItem.Size);

                                        innerArrayLeafNodeBuffer = Decrypt(reader.PstFile, innerArrayLeafNodeBuffer, innerArrayLeafNodeItem.Id);

                                        innerMemoryStream.Write(innerArrayLeafNodeBuffer, 0, innerArrayLeafNodeBuffer.Length);
                                    }
                                }

                                arrayLeafNodeBuffer = innerMemoryStream.ToArray();
                            }

                            if (!hasInnerArray)
                            {
                                arrayLeafNodeBuffer = Decrypt(reader.PstFile, arrayLeafNodeBuffer, arrayLeafNodeItem.Id);
                            }

                            memoryStream.Write(arrayLeafNodeBuffer, 0, arrayLeafNodeBuffer.Length);
                        }
                    }

                    buffer = memoryStream.ToArray();

                    return buffer;
                }
            }

            buffer = Decrypt(reader.PstFile, buffer, dataNode.Id);

            return buffer;

        }

        internal static byte[] GetArrayBuffer(PstFileReader reader, DataStructure dataNode, int index)
        {
            reader.BaseStream.Position = (long)dataNode.Offset;
            byte[] buffer = reader.ReadBytes(dataNode.Size);

            if (buffer[0] == 0x01 && buffer[1] == 0x01)
            {
                Array array = new Array(buffer, reader.PstFile.Is64Bit);

                MemoryStream memoryStream = new MemoryStream();

                DataStructure arrayLeafNodeItem = reader.PstFile.DataIndexTree.GetDataStructure(array.Items[index]);

                if (arrayLeafNodeItem != null)
                {
                    reader.BaseStream.Position = (long)arrayLeafNodeItem.Offset;
                    byte[] arrayLeafNodeBuffer = reader.ReadBytes(arrayLeafNodeItem.Size);

                    memoryStream.Write(arrayLeafNodeBuffer, 0, arrayLeafNodeBuffer.Length);
                }

                buffer = memoryStream.ToArray();
            }

            buffer = Decrypt(reader.PstFile, buffer, dataNode.Id);

            return buffer;
        }

        internal static bool IsFolder(Table table)
        {
            if (table.Entries[MapiPropertyTag.PR_FOLDER_TYPE] != null || table.Entries[MapiPropertyTag.PR_SUBFOLDERS] != null || table.Entries[MapiPropertyTag.PR_CONTAINER_CLASS] != null || table.Entries[MapiPropertyTag.PR_CONTENT_COUNT] != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        internal static string[] GetStringArrayValue(byte[] valueBuffer, System.Text.Encoding encoding)
        {
            if (valueBuffer == null || valueBuffer.Length < 4)
            {
                return new string[0];
            }

            uint count = BitConverter.ToUInt32(valueBuffer, 0);

            if (count > UInt16.MaxValue || count < 0)
            {
                return new string[0];
            }

            string[] values = new string[count];

            for (int k = 0; k < count; k++)
            {
                int startIndex = BitConverter.ToInt32(valueBuffer, k * 4 + 4);
                int endIndex = valueBuffer.Length;

                if (k < count - 1)
                {
                    endIndex = BitConverter.ToInt32(valueBuffer, k * 4 + 8);
                }

                if (endIndex >= startIndex)
                {
                    values[k] = encoding.GetString(valueBuffer, startIndex, endIndex - startIndex);
                }
            }

            return values;
        }

        internal static IList<byte[]> GetBinaryArrayValues(byte[] valueBuffer)
        {
            if (valueBuffer == null || valueBuffer.Length < 4)
            {
                return new List<byte[]>();
            }

            uint count = BitConverter.ToUInt32(valueBuffer, 0);

            if (count > UInt16.MaxValue)
            {
                return new List<byte[]>();
            }

            IList<byte[]> values = new List<byte[]>();

            for (int k = 0; k < count; k++)
            {
                int startIndex = BitConverter.ToInt32(valueBuffer, k * 4 + 4);
                int endIndex = valueBuffer.Length;

                if (k < count - 1)
                {
                    endIndex = BitConverter.ToInt32(valueBuffer, k * 4 + 8);
                }

                if (endIndex >= startIndex)
                {
                    byte[] value = new byte[endIndex - startIndex];
                    System.Array.Copy(valueBuffer, startIndex, value, 0, value.Length);

                    values.Add(value);
                }
            }

            return values;
        }

        internal static string GetCorrectedStringValue(PropertyTag propertyTag, byte[] valueBuffer, System.Text.Encoding encoding)
        {
            int startIndex = 0;

            if (propertyTag.Tag == MapiPropertyTag.PR_SUBJECT.Tag && propertyTag.Type == PropertyType.String)
            {
                if (valueBuffer.Length > 3 && valueBuffer[0] == 1 && valueBuffer[1] == 0)
                {
                    startIndex = 4;
                }
            }
            else if (propertyTag.Tag == MapiPropertyTag.PR_SUBJECT.Tag && propertyTag.Type == PropertyType.String8)
            {
                if (valueBuffer.Length > 1 && valueBuffer[0] == 1)
                {
                    startIndex = 2;
                }
            }

            string stringValue = encoding.GetString(valueBuffer, startIndex, valueBuffer.Length - startIndex);

            return stringValue;
        }

        internal static bool IsInternalReference(uint valueReference)
        {
            if (valueReference < 0xFFFF * 8)
            {
                if (valueReference == ((valueReference >> 4) << 4))
                {
                    return true;
                }
            }

            return false;
        }

        internal static string ConvertNamedPropertyToHex(string name, byte[] guid)
        {
            if (guid != null && guid.Length == 16)
            {
                ulong long1 = BitConverter.ToUInt64(guid, 0);
                ulong long2 = BitConverter.ToUInt64(guid, 8);

                string stringValue = name + "-" + long1.ToString() + "-" + long2.ToString();

                return stringValue;
            }
            else
            {
                return name;
            }
        }

        internal static byte[] DecompressRtfBody(byte[] rtfCompressed)
        {

            String prebufferString = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" +
                                        "{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " +
                                        "\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" +
                                        "{\\colortbl\\red0\\green0\\blue0\n\r\\par " +
                                        "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";

            byte[] prebuffer = System.Text.ASCIIEncoding.ASCII.GetBytes(prebufferString);

            byte[] uncompressed;
            int input = 0;
            int output = 0;

            if (rtfCompressed == null || rtfCompressed.Length < 16)
            {
                throw new ArgumentException("Invalid PR_RTF_COMPRESSION header");
            }

            int compressedSize = (int)ReadUInt32(rtfCompressed, input);
            input += 4;

            int uncompressedSize = (int)ReadUInt32(rtfCompressed, input);
            input += 4;

            int magic = (int)ReadUInt32(rtfCompressed, input);
            input += 4;

            int crc32 = (int)ReadUInt32(rtfCompressed, input);
            input += 4;

            if (compressedSize != rtfCompressed.Length - 4)
            {
                throw new ArgumentException("Invalid PR_RTF_COMPRESSION size.");
            }

            if (magic == 0x414c454d)
            {
                // magic number that identifies the stream as a uncompressed stream
                uncompressed = new byte[uncompressedSize];
                System.Array.Copy(rtfCompressed, input, uncompressed, output, uncompressedSize);
            }
            else if (magic == 0x75465a4c)
            {
                // magic number that identifies the stream as a compressed stream
                uncompressed = new byte[prebuffer.Length + uncompressedSize];
                System.Array.Copy(prebuffer, 0, uncompressed, 0, prebuffer.Length);
                output = prebuffer.Length;

                int flagCount = 0;
                int flags = 0;

                while (output < uncompressed.Length)
                {
                    flags = (flagCount++ % 8 == 0) ? ReadUInt8(rtfCompressed, input++) : flags >> 1;

                    if ((flags & 1) == 1)
                    {
                        int offset = ReadUInt8(rtfCompressed, input++);
                        int length = ReadUInt8(rtfCompressed, input++);

                        offset = (offset << 4) | (length >> 4);
                        length = (length & 0xF) + 2;
                        offset = (output / 4096) * 4096 + offset;

                        if (offset >= output)
                        {
                            offset -= 4096;
                        }

                        int end = offset + length;

                        while (offset < end)
                        {
                            try
                            {
                                uncompressed[output++] = uncompressed[offset++];
                            }
                            catch (IndexOutOfRangeException)
                            {
                                return new byte[0];
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            uncompressed[output++] = rtfCompressed[input++];
                        }
                        catch (IndexOutOfRangeException)
                        {
                            return new byte[0];
                        }
                    }
                }

                rtfCompressed = uncompressed;
                uncompressed = new byte[uncompressedSize];
                System.Array.Copy(rtfCompressed, prebuffer.Length, uncompressed, 0, uncompressedSize);
            }
            else
            {
                throw new ArgumentException("Wrong magic number.");
            }

            return uncompressed;
        }

        internal static ulong ToUInt64(byte[] buffer, int startIndex)
        {
            if (buffer[startIndex] == 0 && buffer[startIndex + 1] == 0 && buffer[startIndex + 2] == 0 && buffer[startIndex + 3] == 0)
            {
                return BitConverter.ToUInt32(buffer, startIndex + 4);
            }
            else
            {
                return BitConverter.ToUInt64(buffer, startIndex);
            }
        }

        private static int ReadUInt8(byte[] buf, int offset)
        {
            return buf[offset] & 0xFF;
        }

        private static int ReadUInt16(int b1, int b2)
        {
            return ((b1 & 0xFF) | ((b2 & 0xFF) << 8)) & 0xFFFF;
        }

        private static long ReadUInt32(byte[] buf, int offset)
        {
            return ((buf[offset] & 0xFF) | ((buf[offset + 1] & 0xFF) << 8) | ((buf[offset + 2] & 0xFF) << 16) | ((buf[offset + 3] & 0xFF) << 24)) & 0x00000000FFFFFFFFL;
        }
    }
}
