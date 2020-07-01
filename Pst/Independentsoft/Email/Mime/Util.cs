using System;
using System.IO;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Summary description for Util.
    /// </summary>
    internal class Util
    {
        private static Random random = new Random();
        internal static byte[] CrlfBuffer = new byte[] { 13, 10, 13, 10 };
        internal static readonly char[] CrlfSeparator = new char[] { '\r', '\n' };
        internal static readonly char[] TabAndSpaceSeparator = new char[] { ' ', '\t' };
        internal static readonly char[] SpaceSeparator = new char[] { ' ' };

        internal static string[] StandardHeaderOrder = new string[] {	
														"Resent-Date", "Resent-From", "Resent-Sender", "Resent-To", "Resent-Cc", 
														"Resent-Bcc", "Resent-Message-ID", "From", "Sender", "Reply-To", 
														"To", "Cc", "Bcc", "Subject", "Date", "Message-ID", "In-Reply-To", 
														"References", "Comments", "Keywords", "Return-Path", "Received", 
														"MIME-Version", "Content-Type", "Content-Transfer-Encoding",
														"Content-ID", "Content-Description" };

        internal static string[] StandardHeaderOrderToLower = new string[] {	
														"resent-date", "resent-from", "resent-sender", "resent-to", "resent-cc", 
														"resent-bcc", "resent-message-id", "from", "sender", "reply-to", 
														"to", "cc", "bcc", "subject", "date", "message-id", "in-reply-to", 
														"references", "comments", "keywords", "return-path", "received", 
														"mime-version", "content-type", "content-transfer-encoding",
														"content-id", "content-description" };

        internal static bool IsStandardKey(string headerKey)
        {
            string headerKeyToLower = headerKey.ToLower();

            for (int i = 0; i < StandardHeaderOrderToLower.Length; i++)
            {
                if (headerKeyToLower == StandardHeaderOrderToLower[i])
                {
                    return true;
                }
            }

            return false;
        }

        internal static string GetCorrectedStandardKey(string headerKey)
        {
            string headerKeyToLower = headerKey.ToLower();

            for (int i = 0; i < StandardHeaderOrderToLower.Length; i++)
            {
                if (headerKeyToLower == StandardHeaderOrderToLower[i])
                {
                    return StandardHeaderOrder[i];
                }
            }

            return headerKey;
        }

        internal static IList<byte[]> GetBodyPartsBuffers(byte[] buffer, string boundary)
        {
            IList<byte[]> bodyPartsBufferList = new List<byte[]>();

            string delimiter = "--" + boundary;
            string lastDelimiter = delimiter + "--";

            byte[] delimiterBuffer = System.Text.Encoding.UTF8.GetBytes(delimiter);
            byte[] lastDelimiterBuffer = System.Text.Encoding.UTF8.GetBytes(lastDelimiter);

            int startIndex = IndexOfBuffer(buffer, delimiterBuffer);

            if (startIndex > -1)
            {
                int lastIndex = IndexOfBuffer(buffer, lastDelimiterBuffer, startIndex + delimiterBuffer.Length);

                if (lastIndex > -1)
                {
                    int nextStartIndex = IndexOfBuffer(buffer, delimiterBuffer, startIndex + delimiterBuffer.Length);

                    while (true)
                    {
                        if (nextStartIndex > (startIndex + delimiterBuffer.Length))
                        {
                            byte[] bodyPartBuffer = new byte[nextStartIndex - startIndex - delimiterBuffer.Length]; //2 == "\r\n"
                            System.Array.Copy(buffer, startIndex + delimiterBuffer.Length, bodyPartBuffer, 0, bodyPartBuffer.Length);
                            bodyPartsBufferList.Add(bodyPartBuffer);
                        }
                        else
                        {
                            break;
                        }

                        if (nextStartIndex == lastIndex)
                        {
                            break;
                        }

                        startIndex = nextStartIndex;

                        nextStartIndex = IndexOfBuffer(buffer, delimiterBuffer, nextStartIndex + delimiterBuffer.Length);
                    }
                }
            }

            return bodyPartsBufferList;
        }

        internal static Encoding GetEncoding(string charset)
        {
            Encoding encoding = Encoding.UTF8;

            if (!string.IsNullOrEmpty(charset))
            {
                try
                {
                    encoding = Encoding.GetEncoding(charset);
                }
                catch
                {
                    if (charset.ToLower().IndexOf("1250") > -1)
                    {
                        charset = "Windows-1250";
                    }
                    else if (charset.ToLower().IndexOf("1252") > -1)
                    {
                        charset = "Windows-1252";
                    }
                    else if (charset.ToLower().IndexOf("iso8859") > -1)
                    {
                        charset = charset.Replace("8859", "-8859");
                    }
                    else if (charset.ToLower().StartsWith("8859"))
                    {
                        charset = charset.Replace("8859", "iso-8859");
                    }

                    try
                    {
                        encoding = Encoding.GetEncoding(charset);
                    }
                    catch
                    {
                    }
                }
            }

            return encoding;
        }

        internal static bool HasToEncodeHeader(string headerValue)
        {
            char[] charArray = headerValue.ToCharArray();

            for (int j = 0; j < charArray.Length; j++)
            {
                int intValue = (int)charArray[j];

                if (intValue > 127)
                {
                    return true;
                }
            }

            return false;
        }

        internal static bool HasToEncodeHeader(string headerName, string headerValue)
        {
            if (headerName.ToLower() != "received")
            {
                char[] charArray = headerValue.ToCharArray();

                for (int j = 0; j < charArray.Length; j++)
                {
                    int intValue = (int)charArray[j];

                    if (intValue > 127)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsEmailAddress(string word)
        {
            int spaceIndex = word.IndexOf(" ");

            if (spaceIndex > -1)
            {
                return false;
            }

            int tabIndex = word.IndexOf("\t");

            if (tabIndex > -1)
            {
                return false;
            }

            int startIndex = word.IndexOf("<");

            if (startIndex == -1)
            {
                return false;
            }

            int endIndex = word.IndexOf(">", startIndex);

            if (endIndex == -1)
            {
                return false;
            }

            int atIndex = word.IndexOf("@");

            if (atIndex == -1 || atIndex < startIndex || atIndex > endIndex)
            {
                return false;
            }

            return true;
        }

        internal static string EncodeHeader(string header, string charset, HeaderEncoding headerEncoding)
        {
            Encoding encoding = null;

            try
            {
                encoding = Encoding.GetEncoding(charset);
            }
            catch
            {
                encoding = Encoding.UTF8;
                charset = "utf-8";
            }

            string encodedHeader = "";

            if (headerEncoding == HeaderEncoding.QuotedPrintable)
            {
                string[] lines = SplitToLinesToEncode(header, 68);

                for (int i = 0; i < lines.Length; i++)
                {
                    if (HasToEncodeHeader(lines[i]))
                    {
                        string encodedLine = EncodeHeaderQuotedPrintable(lines[i], encoding);

                        encodedLine = "=?" + charset + "?" + EnumUtil.ParseHeaderEncoding(headerEncoding) + "?" + encodedLine + "?=";

                        if (encodedHeader.EndsWith(" ") || encodedHeader.EndsWith("\t"))
                        {
                            encodedHeader += encodedLine;
                        }
                        else
                        {
                            encodedHeader += " " + encodedLine;
                        }
                    }
                    else
                    {
                        if (encodedHeader.EndsWith(" ") || encodedHeader.EndsWith("\t"))
                        {
                            encodedHeader += lines[i];
                        }
                        else
                        {
                            encodedHeader += " " + lines[i];
                        }
                    }
                }

            }
            else if (headerEncoding == HeaderEncoding.Binary)
            {
                string[] lines = SplitToLinesToEncode(header, 68);

                for (int i = 0; i < lines.Length; i++)
                {
                    if (HasToEncodeHeader(lines[i]))
                    {
                        byte[] headerBuffer = encoding.GetBytes(lines[i]);
                        string encodedLine = Convert.ToBase64String(headerBuffer, 0, headerBuffer.Length);

                        encodedLine = "=?" + charset + "?" + EnumUtil.ParseHeaderEncoding(headerEncoding) + "?" + encodedLine + "?=";

                        if (encodedHeader.EndsWith(" ") || encodedHeader.EndsWith("\t"))
                        {
                            encodedHeader += encodedLine;
                        }
                        else
                        {
                            encodedHeader += " " + encodedLine;
                        }
                    }
                    else
                    {
                        if (encodedHeader.EndsWith(" ") || encodedHeader.EndsWith("\t"))
                        {
                            encodedHeader += lines[i];
                        }
                        else
                        {
                            encodedHeader += " " + lines[i];
                        }
                    }
                }
            }

            return encodedHeader;
        }

        internal static string EncodeHeaderQuotedPrintable(string text, Encoding encoding)
        {
            byte[] inputByte = encoding.GetBytes(text);
            char[] inputChar = encoding.GetChars(inputByte);

            StringBuilder quotedPrintable = new StringBuilder(inputByte.Length);

            for (int i = 0; i < inputChar.Length; i++)
            {
                int b = inputChar[i];

                if ((b >= 33 && b <= 60) || (b >= 62 && b <= 94) || (b >= 96 && b <= 126))
                {
                    quotedPrintable.Append(inputChar[i]);
                }
                else if (b == 95 || b > 127)
                {
                    string inputString = inputChar[i].ToString();
                    byte[] newInputByte = encoding.GetBytes(inputString);

                    for (int k = 0; k < newInputByte.Length; k++)
                    {
                        byte b1 = newInputByte[k];

                        string hexNumber = b1.ToString("X");

                        if (hexNumber.Length == 1)
                        {
                            hexNumber = "0" + hexNumber;
                        }

                        hexNumber = "=" + hexNumber;

                        quotedPrintable.Append(hexNumber);
                    }
                }
                else
                {
                    string hexNumber = b.ToString("X");

                    if (hexNumber.Length == 1)
                    {
                        hexNumber = "0" + hexNumber;
                    }

                    hexNumber = "=" + hexNumber;

                    quotedPrintable.Append(hexNumber);
                }
            }

            string quotedPrintableString = quotedPrintable.ToString();
            return quotedPrintableString;
        }

        internal static string EncodeBodyQuotedPrintable(string text, Encoding encoding)
        {
            byte[] inputByte = encoding.GetBytes(text);
            char[] inputChar = encoding.GetChars(inputByte);

            StringBuilder quotedPrintable = new StringBuilder(inputByte.Length);

            for (int i = 0; i < inputChar.Length; i++)
            {
                int b = inputChar[i];

                if ((b >= 33 && b <= 60) || (b >= 62 && b <= 126) || b == 9 || b == 32)
                {
                    quotedPrintable.Append(inputChar[i]);
                }
                else if (b > 127)
                {
                    string inputString = inputChar[i].ToString();
                    byte[] newInputByte = encoding.GetBytes(inputString);

                    for (int k = 0; k < newInputByte.Length; k++)
                    {
                        byte b1 = newInputByte[k];

                        string hexNumber = b1.ToString("X");

                        if (hexNumber.Length == 1)
                        {
                            hexNumber = "0" + hexNumber;
                        }

                        hexNumber = "=" + hexNumber;

                        quotedPrintable.Append(hexNumber);
                    }
                }
                else
                {
                    string hexNumber = b.ToString("X");

                    if (hexNumber.Length == 1)
                    {
                        hexNumber = "0" + hexNumber;
                    }

                    hexNumber = "=" + hexNumber;

                    quotedPrintable.Append(hexNumber);
                }
            }

            string quotedPrintableString = quotedPrintable.ToString();
            return quotedPrintableString;
        }

        internal static string DecodeHeader(string encodedWord, ref HeaderEncoding headerEncoding)
        {
            int encodedWordStartIndex = encodedWord.IndexOf("=?");

            if (encodedWordStartIndex == -1)
            {
                return encodedWord;
            }

            int encodedWordEndIndex = encodedWord.IndexOf("?=", encodedWordStartIndex);

            if (encodedWordEndIndex == -1)
            {
                return encodedWord;
            }

            if (encodedWordEndIndex < encodedWordStartIndex)
            {
                return encodedWord;
            }

            //remove SPACE between encoded words
            encodedWord = encodedWord.Replace("?= =?", "?==?");
            encodedWord = encodedWord.Replace("?=  =?", "?==?");
            encodedWord = encodedWord.Replace("?=   =?", "?==?");
            encodedWord = encodedWord.Replace("?=    =?", "?==?");
            encodedWord = encodedWord.Replace("?=     =?", "?==?");
            encodedWord = encodedWord.Replace("?=      =?", "?==?");
            encodedWord = encodedWord.Replace("?=       =?", "?==?");
            encodedWord = encodedWord.Replace("?=        =?", "?==?");

            while(encodedWord.IndexOf("?==?") > 0)
            {
                bool replaced = false;

                int startIndex = encodedWord.IndexOf("?==?");

                int firstQuestionMarkIndex = encodedWord.IndexOf("?", startIndex + 4);

                if (firstQuestionMarkIndex > startIndex)
                {
                    int secondQuestionMarkIndex = encodedWord.IndexOf("?", firstQuestionMarkIndex + 1);

                    if(secondQuestionMarkIndex > firstQuestionMarkIndex)
                    {
                        encodedWord = encodedWord.Substring(0, startIndex) + encodedWord.Substring(secondQuestionMarkIndex + 1);
                        replaced = true;
                    }
                }

                if(!replaced)
                {
                    break;
                }
            }

            string decodedHeader = encodedWord.Substring(0, encodedWordStartIndex);

            while (true)
            {
                int firstQuestionMarkIndex = encodedWord.IndexOf("?", encodedWordStartIndex + 2);
                int secondQuestionMarkIndex = encodedWord.IndexOf("?", firstQuestionMarkIndex + 1);

                if (firstQuestionMarkIndex > -1 && secondQuestionMarkIndex > -1)
                {
                    encodedWordEndIndex = encodedWord.IndexOf("?=", secondQuestionMarkIndex + 1);

                    if (encodedWordEndIndex == -1)
                    {
                        encodedWordEndIndex = encodedWord.Length - 2;
                    }

                    String headerCharSet = encodedWord.Substring(encodedWordStartIndex + 2, firstQuestionMarkIndex - encodedWordStartIndex - 2);
                    string encodingMethod = encodedWord.Substring(firstQuestionMarkIndex + 1, 1);

                    string encodedText = null;

                    if (encodedWordEndIndex > -1)
                    {
                        encodedText = encodedWord.Substring(secondQuestionMarkIndex + 1, encodedWordEndIndex - secondQuestionMarkIndex - 1);
                    }
                    else
                    {
                        encodedText = encodedWord.Substring(secondQuestionMarkIndex + 1);
                    }
                    

                    if (encodingMethod == "Q" || encodingMethod == "q")
                    {
                        headerEncoding = HeaderEncoding.QuotedPrintable;
                        Encoding encoding = null;

                        try
                        {
                            encoding = Encoding.GetEncoding(headerCharSet);
                        }
                        catch
                        {
                            encoding = Encoding.UTF8;
                            headerCharSet = "utf-8";
                        }

                        string decodedWord = DecodeQuotedPrintable(encodedText, encoding);

                        decodedWord = decodedWord.Replace("_", " ");

                        decodedHeader += decodedWord;
                    }
                    else if (encodingMethod == "B" || encodingMethod == "b")
                    {
                        headerEncoding = HeaderEncoding.Binary;
                        Encoding encoding = null;

                        try
                        {
                            encoding = Encoding.GetEncoding(headerCharSet);
                        }
                        catch
                        {
                            encoding = Encoding.UTF8;
                            headerCharSet = "utf-8";
                        }

                        string decodedWord = DecodeBase64(encodedText, encoding);

                        decodedWord = decodedWord.Replace("_", " ");

                        decodedHeader += decodedWord;
                    }
                }

                encodedWordStartIndex = encodedWord.IndexOf("=?", encodedWordEndIndex + 2);

                if (encodedWordStartIndex == -1)
                {
                    decodedHeader += encodedWord.Substring(encodedWordEndIndex + 2, encodedWord.Length - encodedWordEndIndex - 2);
                    break;
                }

                int lastEncodingWordEndIndex = encodedWordEndIndex;

                encodedWordEndIndex = encodedWord.IndexOf("?=", encodedWordStartIndex);

                if (encodedWordEndIndex == -1)
                {
                    decodedHeader += encodedWord.Substring(lastEncodingWordEndIndex + 2, encodedWordStartIndex - lastEncodingWordEndIndex - 2);
                    break;
                }
                else
                {
                    decodedHeader += encodedWord.Substring(lastEncodingWordEndIndex + 2, encodedWordStartIndex - lastEncodingWordEndIndex - 2);
                }
            }

            return decodedHeader;
        }

        internal static string SplitQuotedPrintableBody(string quotedPrintable)
        {
            StringBuilder builder = new StringBuilder(quotedPrintable.Length);

            string[] lines = ToLinesWithCrlf(quotedPrintable);

            for (int k = 0; k < lines.Length; k++)
            {
                if (lines[k].Length < 76)
                {
                    builder.Append(lines[k]);
                }
                else
                {
                    string[] splitted = Split(lines[k]);

                    string line = "";

                    for (int i = 0; i < splitted.Length; i++)
                    {
                        if ((line.Length + splitted[i].Length) < 76)
                        {
                            line += splitted[i];
                        }
                        else
                        {
                            if (line.Length > 0)
                            {
                                line += "=\r\n";
                                builder.Append(line);
                                line = "";
                            }

                            line += splitted[i];
                        }
                    }

                    if (line.Length > 0)
                    {
                        line += "=\r\n";
                        builder.Append(line);
                    }
                }
            }

            return builder.ToString();
        }

        internal static string SplitBase64Body(string input)
        {
            int inputLength = input.Length;
            int numberOfInserts = (int)(input.Length / 76);
            int capacity = inputLength + (numberOfInserts * 2);

            StringBuilder builder = new StringBuilder(capacity);

            int nextPosition = 76;
            int inputNextPosition = 0;

            while (nextPosition < capacity + 1)
            {
                string temp = input.Substring(inputNextPosition, 76);
                builder.Append(temp + "\r\n");
                nextPosition += 78;
                inputNextPosition += 76;
            }

            string last = input.Substring(inputNextPosition);
            builder.Append(last);

            return builder.ToString();
        }

        internal static string DecodeBase64(string encodedText, Encoding encoding)
        {
            byte[] decodedBuffer = DecodeBase64(encodedText);
            return encoding.GetString(decodedBuffer, 0, decodedBuffer.Length);
        }

        internal static byte[] DecodeBase64(string encodedText)
        {
            if (encodedText == null || encodedText.Length == 0)
            {
                return new byte[0];
            }

            encodedText = encodedText.Replace("\r\n", String.Empty);
            encodedText = encodedText.TrimEnd('.');

            byte[] decodedBuffer = new byte[0];

            try
            {
                decodedBuffer = Convert.FromBase64String(encodedText);
            }
            catch
            {
                //try to receover
                if (!encodedText.EndsWith("==") && encodedText.EndsWith("="))
                {
                    encodedText += "=";
                }
                else if (!encodedText.EndsWith("==") && !encodedText.EndsWith("="))
                {
                    encodedText += "==";
                }

                encodedText = encodedText.Replace(" ", String.Empty);
            }

            try
            {
                decodedBuffer = Convert.FromBase64String(encodedText);
            }
            catch
            {
            }

            return decodedBuffer;
        }

        internal static byte[] DecodeBase64Attachment(string encodedText)
        {
            if (encodedText == null || encodedText.Length == 0)
            {
                return new byte[0];
            }

            encodedText = encodedText.TrimEnd('.');

            byte[] decodedBuffer = new byte[0];

            try
            {
                decodedBuffer = Convert.FromBase64String(encodedText);
            }
            catch
            {
                //try to recover
                if (!encodedText.EndsWith("==") && encodedText.EndsWith("="))
                {
                    encodedText += "=";
                }
                else if (!encodedText.EndsWith("==") && !encodedText.EndsWith("="))
                {
                    encodedText += "==";
                }

                encodedText = encodedText.Replace(" ", String.Empty);

                try
                {
                    decodedBuffer = Convert.FromBase64String(encodedText);
                }
                catch
                {
                }
            }

            return decodedBuffer;
        }

        internal static string DecodeQuotedPrintable(string encodedText, Encoding encoding)
        {
            try
            {
                byte[] decodedBuffer = DecodeQuotedPrintable(encodedText);
                return encoding.GetString(decodedBuffer, 0, decodedBuffer.Length);
            }
            catch (OverflowException)
            {
                return encodedText;
            }
            catch (Exception ex)
            {
                throw new MessageFormatException(ex.Message, ex);
            }
        }

        internal static byte[] DecodeQuotedPrintable(string encodedText)
        {
            byte[] returnBuffer = null;

            encodedText = encodedText.Replace("=09", "\t");
            encodedText = encodedText.Replace("=20", " ");
            encodedText = encodedText.Replace("=\r\n", "");

            byte[] inputByte = Encoding.UTF8.GetBytes(encodedText);
            byte[] outputByte = new byte[inputByte.Length];

            int outputByteCount = 0;

            int newLineIndex = 0;
            int lineBegin = 0;
            byte[] line = null;

            for (int i = 0; i < inputByte.Length; i++)
            {
                if (inputByte[i] == 13)
                {
                    continue;
                }
                else if (inputByte[i] == 10)
                {
                    newLineIndex = i + 1;
                }
                else if (i < inputByte.Length - 1)
                {
                    continue;
                }
                else
                {
                    newLineIndex = inputByte.Length;
                }

                if (newLineIndex > 1)
                {
                    line = new byte[newLineIndex - lineBegin];
                    System.Array.Copy(inputByte, lineBegin, line, 0, newLineIndex - lineBegin);
                    lineBegin = newLineIndex;
                }
                else
                {
                    line = new byte[inputByte.Length - lineBegin];
                    System.Array.Copy(inputByte, lineBegin, line, 0, inputByte.Length - lineBegin);
                    lineBegin = inputByte.Length;
                }

                for (int j = 0; j < line.Length; j++)
                {
                    byte currentInteger = line[j];

                    if (currentInteger > 126)
                    {
                        continue;
                    }

                    if (currentInteger == 61)
                    {
                        if (j < line.Length - 2)
                        {
                            int firstNumber = line[j + 1];
                            int secondNumber = line[j + 2];

                            if ((firstNumber > 47 && firstNumber < 58) || (firstNumber > 64 && firstNumber < 71) || (firstNumber > 96 && firstNumber < 103) &&
                                (secondNumber > 47 && secondNumber < 58) || (secondNumber > 64 && secondNumber < 71) || (secondNumber > 96 && secondNumber < 103))
                            {
                                string hexNumber = Convert.ToString((char)firstNumber) + Char.ToString((char)secondNumber);

                                try
                                {
                                    byte intNumber = Byte.Parse(hexNumber, System.Globalization.NumberStyles.HexNumber);
                                    outputByte[outputByteCount++] = intNumber;
                                }
                                catch
                                {
                                }

                                j += 2;
                            }
                        }
                    }
                    else
                    {
                        outputByte[outputByteCount++] = currentInteger;
                    }
                }
            }

            returnBuffer = new byte[outputByteCount];
            System.Array.Copy(outputByte, 0, returnBuffer, 0, outputByteCount);

            return returnBuffer;
        }

        internal static IList<Mailbox> ParseMailboxes(string input)
        {
            IList<Mailbox> mailboxes = new List<Mailbox>();

            int separator1Index = input.IndexOf(";");
            int separator2Index = input.IndexOf(",");

            int separatorIndex = separator1Index;

            if ((separator1Index == -1 && separator2Index > -1) || (separator1Index > -1 && separator2Index > -1 && separator2Index < separator1Index))
            {
                separatorIndex = separator2Index;
            }

            while (separatorIndex > -1)
            {
                int openQuoteIndex = input.IndexOf("\"");
                int closeQuoteIndex = input.IndexOf("\"", openQuoteIndex + 1);

                if (openQuoteIndex < separatorIndex && separatorIndex < closeQuoteIndex)
                {
                    separator1Index = input.IndexOf(";", closeQuoteIndex + 1);
                    separator2Index = input.IndexOf(",", closeQuoteIndex + 1);

                    separatorIndex = separator1Index;

                    if ((separator1Index == -1 && separator2Index > -1) || (separator1Index > -1 && separator2Index > -1 && separator2Index < separator1Index))
                    {
                        separatorIndex = separator2Index;
                    }
                }
                else
                {
                    string mailboxString = input.Substring(0, separatorIndex);
                    input = input.Substring(separatorIndex + 1);

                    Mailbox mailbox = new Mailbox(mailboxString);
                    mailboxes.Add(mailbox);

                    separator1Index = input.IndexOf(";");
                    separator2Index = input.IndexOf(",");

                    separatorIndex = separator1Index;

                    if ((separator1Index == -1 && separator2Index > -1) || (separator1Index > -1 && separator2Index > -1 && separator2Index < separator1Index))
                    {
                        separatorIndex = separator2Index;
                    }
                }
            }

            Mailbox lastMailbox = new Mailbox(input);
            mailboxes.Add(lastMailbox);

            return mailboxes;
        }

        private static string[] ToLinesWithCrlf(string input)
        {
            IList<String> list = new List<String>();

            int crlfIndex = input.IndexOf("\r\n");
            int startPosition = 0;
            string line = "";

            while (crlfIndex > -1)
            {
                line = input.Substring(startPosition, crlfIndex + 2 - startPosition);
                startPosition = crlfIndex + 2;

                list.Add(line);

                crlfIndex = input.IndexOf("\r\n", startPosition);
            }

            line = input.Substring(startPosition);
            list.Add(line);

            string[] lines = new string[list.Count];

            for (int i = 0; i < list.Count; i++)
            {
                lines[i] = list[i];
            }

            return lines;
        }

        private static string[] SplitToLinesToEncode(string input, int length)
        {
            IList<String> list = new List<String>();

            string[] splitted = Split(input);

            string line = "";

            for (int i = 0; i < splitted.Length; i++)
            {
                if (IsEmailAddress(splitted[i]))
                {
                    if (line.Length > 0)
                    {
                        list.Add(line);
                        line = "";
                    }

                    list.Add(splitted[i]);
                }
                else if ((line.Length + splitted[i].Length) < length)
                {
                    line += splitted[i];
                }
                else
                {
                    if (line.Length > 0)
                    {
                        list.Add(line);
                        line = "";
                    }

                    line += splitted[i];
                }
            }

            if (line.Length > 0)
            {
                list.Add(line);
            }

            string[] lines = new string[list.Count];

            for (int i = 0; i < list.Count; i++)
            {
                lines[i] = list[i];
            }

            return lines;
        }

        internal static string SplitHeaderLines(string input, int length)
        {
            if (input.Length <= length)
            {
                return input;
            }

            IList<String> list = new List<String>();

            string[] splitted = Split(input);

            string line = "";

            for (int i = 0; i < splitted.Length; i++)
            {
                if ((line.Length + splitted[i].Length) < length)
                {
                    line += splitted[i];
                }
                else
                {
                    if (line.Length > 0)
                    {
                        list.Add(line);
                        line = "";
                    }

                    line += splitted[i];
                }
            }

            if (line.Length > 0)
            {
                list.Add(line);
            }

            string splittedHeader = "";

            for (int i = 0; i < list.Count; i++)
            {
                splittedHeader += list[i];

                if (i < list.Count - 1)
                {
                    splittedHeader += "\r\n ";
                }
            }

            return splittedHeader;
        }


        private static string[] Split(string input)
        {
            if (input.Length > 2 * 1024 * 1024) //do not split body part over 2 MB
            {
                return new string[] { input };
            }

            IList<String> list = new List<String>();
            char[] chars = input.ToCharArray();

            string word = "";

            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] != ' ' && chars[i] != '\t')
                {
                    if (i == 0)
                    {
                        word += chars[i];
                    }
                    else
                    {
                        if (chars[i - 1] == ' ' || chars[i - 1] == '\t')
                        {
                            if (word.Length > 0)
                            {
                                list.Add(word);
                            }

                            word = "";
                            word += chars[i];
                        }
                        else
                        {
                            word += chars[i];
                        }
                    }
                }
                else
                {
                    if (word.Length > 0)
                    {
                        list.Add(word);
                    }

                    word = "";
                    word += chars[i];
                }
            }

            if (word.Length > 0)
            {
                list.Add(word);
            }

            string[] splitted = new string[list.Count];

            for (int i = 0; i < list.Count; i++)
            {
                splitted[i] = list[i];
            }

            return splitted;
        }

        internal static int IndexOfBuffer(byte[] buffer, byte[] subbuffer)
        {
            return IndexOfBuffer(buffer, subbuffer, 0);
        }

        internal static int IndexOfBuffer(byte[] buffer, byte[] subbuffer, int startIndex)
        {
            bool found = false;

            for (int i = startIndex; i <= buffer.Length - subbuffer.Length; i++)
            {
                for (int j = 0; j < subbuffer.Length; j++)
                {
                    if (buffer[i + j] == subbuffer[j])
                    {
                        found = true;
                    }
                    else
                    {
                        found = false;
                        break;
                    }
                }

                if (found)
                {
                    return i;
                }
            }

            return -1;
        }

        internal static string Unfolding(string input)
        {
            string output = input.Replace("\r\n ", " ");

            output = output.Replace("\r\n\t", " ");

            return output;
        }

        internal static string EncodeEscapeCharacters(string input)
        {
            if (input == null)
            {
                return null;
            }

            string output = input.Replace("&", "&amp;"); //must be the first
            output = output.Replace("<", "&lt;");
            output = output.Replace(">", "&gt;");
            output = output.Replace("'", "&apos;");
            output = output.Replace("\"", "&quot;");

            return output;
        }

        internal static string DecodeEscapeCharacters(string input)
        {
            if (input == null)
            {
                return null;
            }

            string output = input.Replace("&amp;", "&"); //must be the first
            output = output.Replace("&lt;", "<");
            output = output.Replace("&gt;", ">");
            output = output.Replace("&apos;", "'");
            output = output.Replace("&quot;", "\"");

            return output;
        }

        internal static byte[] RecoverMessage(byte[] buffer)
        {
            MemoryStream correctBuffer = new MemoryStream();

            for (int index = 0; index < buffer.Length; index++)
            {
                switch (buffer[index])
                {
                    case 13:
                        if (index == buffer.Length - 1)
                        {
                            correctBuffer.WriteByte(13);
                            correctBuffer.WriteByte(10);
                        }
                        else if (buffer[index + 1] != 10)
                        {
                            correctBuffer.WriteByte(13);
                            correctBuffer.WriteByte(10);
                        }
                        else
                        {
                            correctBuffer.WriteByte(13);
                            correctBuffer.WriteByte(buffer[index + 1]);
                            index = index + 1;
                        }

                        break;

                    case 10:
                        if (index == 0)
                        {
                            correctBuffer.WriteByte(13);
                            correctBuffer.WriteByte(10);
                        }
                        else if (buffer[index - 1] != 13)
                        {
                            correctBuffer.WriteByte(13);
                            correctBuffer.WriteByte(10);
                        }
                        else
                        {
                            correctBuffer.WriteByte(10);
                        }

                        break;

                    default:

                        correctBuffer.WriteByte(buffer[index]);
                        break;
                }
            }

            if (correctBuffer.Length == 2)
            {
                correctBuffer.WriteByte(13);
                correctBuffer.WriteByte(10);
            }

            return correctBuffer.ToArray();
        }

        internal static byte[] RemoveCRLF(byte[] input)
        {
            MemoryStream outputMemoryStream = new MemoryStream(input.Length);

            for (int index = 0; index < input.Length - 1; index++)
            {
                if (input[index] != 13 && input[index + 1] != 10)
                {
                    outputMemoryStream.WriteByte(input[index]);
                }
                else
                {
                    index++;
                }
            }

            //copy last
            if (input.Length > 1)
            {
                if (input[input.Length-2] != 13 && input[input.Length-1] != 10)
                {
                    outputMemoryStream.WriteByte(input[input.Length-1]);
                }
            }

            return outputMemoryStream.ToArray();
        }

        internal static string NewGuid()
        {
            return NextRandom() + "-" + NextRandom() + "-" + NextRandom();
        }

        internal static int NextRandom()
        {
            return random.Next(Int32.MaxValue);
        }

        internal static DateTime ParseRFC822Date(string adate)
        {
            string tmp;
            string[] resp;
            string dayName;
            string dpart;
            string hour, minute;
            string timeZone;
            System.DateTime dt = System.DateTime.Now;

            tmp = Regex.Replace(adate, "(\\([^(].*\\))", "");

            // strip extra white spaces
            tmp = Regex.Replace(tmp, "\\s+", " ");
            tmp = Regex.Replace(tmp, "^\\s+", "");
            tmp = Regex.Replace(tmp, "\\s+$", "");

            // extract week name part
            resp = tmp.Split(new char[] { ',' }, 2);
            if (resp.Length == 2)
            {
                // there's week name
                dayName = resp[0];
                tmp = resp[1];
            }
            else dayName = "";

            try
            {
                // extract date and time
                int pos = tmp.LastIndexOf(" ");

                if (pos < 1)
                {
                    throw new FormatException("Probably not a date");
                }

                dpart = tmp.Substring(0, pos - 1);
                timeZone = tmp.Substring(pos + 1);
                dt = Convert.ToDateTime(dpart);

                // check weekDay name
                // this must be done before convert to GMT 
                if (dayName != string.Empty)
                {
                    if ((dt.DayOfWeek == DayOfWeek.Friday && dayName != "Fri") ||
                        (dt.DayOfWeek == DayOfWeek.Monday && dayName != "Mon") ||
                        (dt.DayOfWeek == DayOfWeek.Saturday && dayName != "Sat") ||
                        (dt.DayOfWeek == DayOfWeek.Sunday && dayName != "Sun") ||
                        (dt.DayOfWeek == DayOfWeek.Thursday && dayName != "Thu") ||
                        (dt.DayOfWeek == DayOfWeek.Tuesday && dayName != "Tue") ||
                        (dt.DayOfWeek == DayOfWeek.Wednesday && dayName != "Wed")
                        )
                        throw new FormatException("Invalid week of day");
                }

                // adjust to localtime
                if (Regex.IsMatch(timeZone, "[+\\-][0-9][0-9][0-9][0-9]"))
                {
                    // it's a modern ANSI style timezone
                    int factor = 0;
                    hour = timeZone.Substring(1, 2);
                    minute = timeZone.Substring(3, 2);
                    if (timeZone.Substring(0, 1) == "+") factor = 1;
                    else if (timeZone.Substring(0, 1) == "-") factor = -1;
                    else throw new FormatException("incorrect time zone");
                    dt = dt.AddHours(factor * Convert.ToInt32(hour));
                    dt = dt.AddMinutes(factor * Convert.ToInt32(minute));
                }
                else
                {
                    // it's a old style military time zone ?
                    switch (timeZone)
                    {
                        case "A": dt = dt.AddHours(1); break;
                        case "B": dt = dt.AddHours(2); break;
                        case "C": dt = dt.AddHours(3); break;
                        case "D": dt = dt.AddHours(4); break;
                        case "E": dt = dt.AddHours(5); break;
                        case "F": dt = dt.AddHours(6); break;
                        case "G": dt = dt.AddHours(7); break;
                        case "H": dt = dt.AddHours(8); break;
                        case "I": dt = dt.AddHours(9); break;
                        case "K": dt = dt.AddHours(10); break;
                        case "L": dt = dt.AddHours(11); break;
                        case "M": dt = dt.AddHours(12); break;
                        case "N": dt = dt.AddHours(-1); break;
                        case "O": dt = dt.AddHours(-2); break;
                        case "P": dt = dt.AddHours(-3); break;
                        case "Q": dt = dt.AddHours(-4); break;
                        case "R": dt = dt.AddHours(-5); break;
                        case "S": dt = dt.AddHours(-6); break;
                        case "T": dt = dt.AddHours(-7); break;
                        case "U": dt = dt.AddHours(-8); break;
                        case "V": dt = dt.AddHours(-9); break;
                        case "W": dt = dt.AddHours(-10); break;
                        case "X": dt = dt.AddHours(-11); break;
                        case "Y": dt = dt.AddHours(-12); break;
                        case "Z":
                        case "UT":
                        case "GMT": break;    // It's UTC
                        case "BST": dt = dt.AddHours(1); break;
                        case "EST": dt = dt.AddHours(5); break;
                        case "EDT": dt = dt.AddHours(4); break;
                        case "CST": dt = dt.AddHours(6); break;
                        case "CDT": dt = dt.AddHours(5); break;
                        case "MST": dt = dt.AddHours(7); break;
                        case "MDT": dt = dt.AddHours(6); break;
                        case "PST": dt = dt.AddHours(8); break;
                        case "PDT": dt = dt.AddHours(7); break;
                    }
                }
            }
            catch (Exception)
            {
                //throw new FormatException(string.Format("Invalid date:{0}:{1}", e.Message, adate));
            }
            return dt;
        }
    }
}