using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Windows;
using System.Threading;
//using System.Windows.Controls;
//using System.Windows.Documents;
using System.Xml;
using PRGX.MT.RTF;
using PRGX.MT.RTF.converter.html;
using PRGX.MT.RTF.Sources;
using System.Security;

namespace Independentsoft.Msg
{
    internal class Util
    {
        internal static DateTime GetDateTime(int minutes)
        {
            if (minutes > 0)
            {
                DateTime year1601 = new DateTime(1601, 1, 1);

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

        internal static string ConvertNamedPropertyToHex(uint id, byte[] guid)
        {
            if (guid != null && guid.Length == 16)
            {
                ulong long1 = BitConverter.ToUInt64(guid, 0);
                ulong long2 = BitConverter.ToUInt64(guid, 8);

                string stringValue = id.ToString() + "-" + long1.ToString() + "-" + long2.ToString();

                return stringValue;
            }
            else
            {
                return id.ToString();
            }
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

        internal static int IndexOfNamedProperty(IList<NamedProperty> namedProperties, NamedProperty namedProperty)
        {
            if (namedProperties.Count == 0)
            {
                return -1;
            }

            bool found = false;

            for (int i = 0; i < namedProperties.Count; i++)
            {
                NamedProperty current = (NamedProperty)namedProperties[i];

                if (namedProperty.Name != null && current.Name == namedProperty.Name)
                {
                    found = true;
                }
                else if (current.Id == namedProperty.Id && namedProperty.Type == NamedPropertyType.Numerical)
                {
                    found = true;
                }

                if (found)
                {
                    bool isGuidEqual = true;

                    if (current.Guid != null && namedProperty.Guid != null && current.Guid.Length == namedProperty.Guid.Length)
                    {
                        for (int j = 0; j < current.Guid.Length; j++)
                        {
                            if (current.Guid[j] != namedProperty.Guid[j])
                            {
                                isGuidEqual = false;
                            }
                        }
                    }
                    else
                    {
                        isGuidEqual = false;
                    }

                    if (isGuidEqual)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        internal static string[] GetStringArrayValues(byte[] valueBuffer, System.Text.Encoding encoding)
        {
            if (valueBuffer == null || valueBuffer.Length < 4)
            {
                return new string[0];
            }

            uint count = BitConverter.ToUInt32(valueBuffer, 0);

            if (count > UInt16.MaxValue)
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

        internal static uint GetPropertySize(byte[] value, PropertyType type, System.Text.Encoding encoding)
        {
            if (value == null && (type == PropertyType.String || type == PropertyType.String8))
            {
                return 1;
            }
            else if (value != null && (type == PropertyType.String || type == PropertyType.String8))
            {
                return (uint)(value.Length + encoding.GetBytes("\0").Length);
            }
            else if (value == null)
            {
                return 0;
            }
            else
            {
                return (uint)value.Length;
            }
        }

        internal static uint GetTypeHexMask(PropertyType type)
        {
            if (type == PropertyType.Integer16)
            {
                return 0x0002;
            }
            else if (type == PropertyType.Integer32)
            {
                return 0x0003;
            }
            else if (type == PropertyType.Floating32)
            {
                return 0x0004;
            }
            else if (type == PropertyType.Floating64)
            {
                return 0x0005;
            }
            else if (type == PropertyType.Currency)
            {
                return 0x0006;
            }
            else if (type == PropertyType.FloatingTime)
            {
                return 0x0007;
            }
            else if (type == PropertyType.Boolean)
            {
                return 0x000B;
            }
            else if (type == PropertyType.Object)
            {
                return 0x000D;
            }
            else if (type == PropertyType.Integer64)
            {
                return 0x0014;
            }
            else if (type == PropertyType.String8)
            {
                return 0x001E;
            }
            else if (type == PropertyType.String)
            {
                return 0x001F;
            }
            else if (type == PropertyType.Time)
            {
                return 0x0040;
            }
            else if (type == PropertyType.Guid)
            {
                return 0x0048;
            }
            else if (type == PropertyType.Binary)
            {
                return 0x0102;
            }
            else if (type == PropertyType.MultipleInteger16)
            {
                return 0x1002;
            }
            else if (type == PropertyType.MultipleInteger32)
            {
                return 0x1003;
            }
            else if (type == PropertyType.MultipleFloating32)
            {
                return 0x1004;
            }
            else if (type == PropertyType.MultipleFloating64)
            {
                return 0x1005;
            }
            else if (type == PropertyType.MultipleCurrency)
            {
                return 0x1006;
            }
            else if (type == PropertyType.MultipleFloatingTime)
            {
                return 0x1007;
            }
            else if (type == PropertyType.MultipleInteger64)
            {
                return 0x1014;
            }
            else if (type == PropertyType.MultipleString8)
            {
                return 0x101E;
            }
            else if (type == PropertyType.MultipleString)
            {
                return 0x101F;
            }
            else if (type == PropertyType.MultipleTime)
            {
                return 0x1040;
            }
            else if (type == PropertyType.MultipleGuid)
            {
                return 0x1048;
            }
            else if (type == PropertyType.MultipleBinary)
            {
                return 0x1102;
            }
            else //string
            {
                return 0x001E;
            }
        }

        internal static byte[] CreateReplyToEntries(string replyTo)
        {
            String[] emailAddresses = replyTo.Split(new char[] { ';' });
            MemoryStream memoryStream = new MemoryStream();

            for (int i = 0; i < emailAddresses.Length; i++)
            {
                byte[] prefix = new byte[] { 0x00, 0x00, 0x00, 0x00, 0x81, 0x2B, 0x1F, 0xA4, 0xBE, 0xA3, 0x10, 0x19, 0x9D, 0x6E, 0x00, 0xDD, 0x01, 0x0F, 0x54, 0x02, 0x00, 0x00, 0x01, 0x80 };

                byte[] suffix = System.Text.Encoding.Unicode.GetBytes(emailAddresses[i] + "\0SMTP\0" + emailAddresses[i] + "\0");
                byte[] suffix2 = new byte[] { 0x00, 0x00 };

                int entryLength = prefix.Length + suffix.Length;
             
                memoryStream.Write(BitConverter.GetBytes(entryLength), 0, 4);
                memoryStream.Write(prefix, 0, prefix.Length);
                memoryStream.Write(suffix, 0, suffix.Length);
                memoryStream.Write(suffix2, 0, suffix2.Length);
            }

            byte[] entries = new byte[8 + memoryStream.Position];

            int count = emailAddresses.Length;
            int length = entries.Length - 8;

            System.Array.Copy(BitConverter.GetBytes(count), 0, entries, 0, 4);
            System.Array.Copy(BitConverter.GetBytes(length), 0, entries, 4, 4);
            System.Array.Copy(memoryStream.GetBuffer(), 0, entries, 8, memoryStream.Position);

            return entries;
        }

        private static System.Text.Encoding GetRtfEncoding(string rtf)
        {
            System.Text.Encoding encoding = System.Text.Encoding.GetEncoding(1252);

            if (rtf.Length > 7) // PRGX: Make sure the string is long enough to prevent errors
            {
              int ansicpgStartIndex = rtf.IndexOf("ansicpg");
              int ansicpgEndIndex = rtf.IndexOf("\\", ansicpgStartIndex + 7);

              if (ansicpgStartIndex > -1 && ansicpgEndIndex > -1)
              {
                string codePageString = rtf.Substring(ansicpgStartIndex + 7, ansicpgEndIndex - ansicpgStartIndex - 7);

                try
                {
                  int codePage = Int32.Parse(codePageString);
                  encoding = System.Text.Encoding.GetEncoding(codePage);
                }
                catch
                {
                }
              }
            }
            return encoding;
        }

        private static IDictionary<String, Encoding> GetFontTable(string rtf)
        {
            IDictionary<String, Encoding> fontTable = new Dictionary<String, Encoding>();

            int start = rtf.IndexOf("{\\fonttbl");

            if (start > 0)
            {
                int end = rtf.IndexOf("}}", start);

                if (end > 0)
                {
                    string fontTbl = rtf.Substring(start + 8, end - start);
                    StringReader reader = new StringReader(fontTbl);
                    string line = "";

                    while ((line = reader.ReadLine()) != null)
                    {
                        int lineStart = line.IndexOf("{");

                        if (lineStart > -1)
                        {
                            int lineEnd = line.IndexOf("}", lineStart);

                            if (lineEnd > -1)
                            {
                                line = line.Substring(lineStart + 1, lineEnd - lineStart - 1);

                                string[] words = line.Split(new char[] { ' ', '\\' });

                                if (words.Length > 1 && words[1].StartsWith("f"))
                                {
                                    string fontName = words[1];
                                    System.Text.Encoding fontEncoding = null;

                                    for (int i = 2; i < words.Length; i++)
                                    {
                                        if (words[i] == "fcharset128")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding(932);
                                        }
                                        else if (words[i] == "fcharset129")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("ks_c_5601-1987");
                                        }
                                        else if (words[i] == "fcharset134")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("gb2312");
                                        }
                                        else if (words[i] == "fcharset136")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("big5");
                                        }
                                        else if (words[i] == "fcharset161")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1253");
                                        }
                                        else if (words[i] == "fcharset162")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1254");
                                        }
                                        else if (words[i] == "fcharset163")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1258");
                                        }
                                        else if (words[i] == "fcharset177")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1255");
                                        }
                                        else if (words[i] == "fcharset178")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1256");
                                        }
                                        else if (words[i] == "fcharset186")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1257");
                                        }
                                        else if (words[i] == "fcharset204")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1251");
                                        }
                                        else if (words[i] == "fcharset222")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-874");
                                        }
                                        else if (words[i] == "fcharset238")
                                        {
                                            fontEncoding = System.Text.Encoding.GetEncoding("windows-1250");
                                        }
                                    }

                                    if (!fontTable.ContainsKey(fontName))
                                    {
                                        fontTable.Add(fontName, fontEncoding);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            return fontTable;
        }

        private static System.Text.Encoding GetHexCharEncoding(string html, IDictionary<String, Encoding> fontTable, System.Text.Encoding defaultEncoding, int end)
        {
            string subHtml = html.Substring(0, end);
            System.Text.Encoding hexCharEncoding = defaultEncoding;
            int lastPosition = -1;

            foreach (string fontName in fontTable.Keys)
            {
                int position = subHtml.LastIndexOf("\\" + fontName);

                if (position > -1 && position > lastPosition)
                {
                    hexCharEncoding = fontTable[fontName] != null ? fontTable[fontName] : defaultEncoding;
                    lastPosition = position;
                }
            }

            return hexCharEncoding;
        }
    [SecuritySafeCritical]
    internal static HtmlText ConvertRtfToHtml(byte[] rtfBody)
    {
      HtmlText retVal = null;
      IRtfConverter rtfConverter = new ConvertRtfToHtml();
      IRtfSource source = new RtfSourceByteArray(rtfBody);
      string textValue = string.Empty;

      try
      {
        using (Stream oFile = new MemoryStream())
        {
          rtfConverter.convert(source, oFile);

          oFile.Position = 0;
          using (var reader = new StreamReader(oFile))
          {
            textValue = reader.ReadToEnd();
          }
        }

        retVal = new HtmlText(textValue, rtfConverter.getEncoding());
      }
      catch (Exception ex)
      {
        // Ignore unhandled RTF conversions.        throw;
      }

      return retVal;
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
                if (uncompressedSize > rtfCompressed.Length - input)
                {
                    uncompressedSize = rtfCompressed.Length - input;
                }

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

        internal static Encoding GetEncoding(long codePage)
        {
            Encoding encoding = Encoding.Default;

            if (codePage > 0)
            {
                try
                {
                    encoding = Encoding.GetEncoding((int)codePage);
                }
                catch
                {
                    encoding = Encoding.UTF8;
                }
            }

            return encoding;
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

        internal static string Html2Rtf(string sHtml, int iCodePage)
        {
            int i;
            int p;
            int j;
            string strTag;
            string strExtra;
            char ch;
            string sz = new string(new char[5]);
            string pCmp;

            string sRtf = sHtml;
            if (iCodePage == 65001)
                iCodePage = 0;
            for (i = sRtf.Length - 1; i >= 0; --i)
            {
                switch (ch = sRtf[i])
                {
                    case '>':
                        {
                            break;
                        }
                    case '<':
                        {
                            break;
                        }
                    case '\t':
                        {
                            sRtf = sRtf.Remove(i, 1);
                            sRtf = sRtf.Insert(i, "\\tab");
                            break;
                        }
                    case '{':
                    case '}':
                    case '\\':
                        {
                            sRtf = sRtf.Insert(i, "\\");
                            break;
                        }
                    case '\r':
                        {
                            if (i < sRtf.Length - 1 && sRtf[i + 1] == '\n')
                                sRtf = sRtf.Insert(i, " ");
                            break;
                        }
                }

                if (ch > 128)
                {
                    sz = string.Format("\\\"{0:X2}", ch);
                    sRtf = sRtf.Remove(i, 1);
                    sRtf = sRtf.Insert(i, sz);
                }
            }

            for (i = sRtf.Length - 2; i >= 0; --i)
            {
                if (sRtf[i] == '<')
                {
                    p = 0;
                    for (j = i + 1; j < sRtf.Length; ++j)
                    {
                        if (sRtf[j] == '>')
                        {
                            p = j;
                            break;
                        }
                    }

                    if (p == 0)
                    {
                        sRtf = sHtml;
                        return sRtf;
                    }

                    strTag = sRtf.Substring(i, p + 1 - i);
                    pCmp = strTag;

                    if (pCmp.ToUpper().Substring(0, 2) != "P>" || pCmp.ToUpper().Substring(0, 2) != "P " || pCmp.ToUpper().Substring(0, 3) != "BR>" || pCmp.ToUpper().Substring(0, 3) != "BR ")
                        strExtra = "\\par ";
                    else
                        strExtra = "";

                    if (strTag[1] == '/')
                        strTag = "{\\*\\htmltag8 " + strExtra + strTag + "}";
                    else
                        strTag = "{\\*\\htmltag0 " + strExtra + strTag + "}";
                    sRtf = sRtf.Remove(i, p - i + 1);

                    sRtf = sRtf.Insert(i, strTag);
                }
            }

            string returnValue = "{\\rtf1\\ansi\\ansicpg" + iCodePage + "\\fromhtml1\\deff0{\\fonttbl\r\n" + "{\\f0\\fswiss\\fcharset" + iCodePage + " Arial;}\r\n" + "{\\f1\\fmodern\\fcharset" + iCodePage + " Courier New;}\r\n" + "{\\f2\\fnil\\fcharset" + iCodePage + " Symbol;}\r\n" + "{\\f3\\fmodern\\fcharset" + iCodePage + " Courier New;}}\r\n" + "\\uc1\\pard\\plain\\deftab360 \\f0\\fs24\r\n" + sRtf + "}";

            return returnValue;
        }
    }
}
