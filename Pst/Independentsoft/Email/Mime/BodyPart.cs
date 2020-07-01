using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class BodyPart.
    /// </summary>
    public class BodyPart
    {
        private HeaderList headers = new HeaderList();
        private BodyPartList bodyParts = new BodyPartList();
        private string body;
        private Message embeddedMessage;
        private HeaderEncoding headerEncoding = HeaderEncoding.QuotedPrintable;
        private string headerCharSet = "utf-8";

        /// <summary>
        /// Initializes a new instance of the <see cref="BodyPart"/> class.
        /// </summary>
        public BodyPart()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BodyPart"/> class.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        public BodyPart(Attachment attachment)
        {
            if (attachment != null)
            {
                if (attachment.ContentType != null)
                {
                    this.ContentType = attachment.ContentType;
                }

                this.ContentID = attachment.ContentID;
                this.ContentLocation = attachment.ContentLocation;
                this.ContentTransferEncoding = ContentTransferEncoding.Base64;

                if (attachment.ContentDisposition != null)
                {
                    ContentDisposition contentDisposition = attachment.ContentDisposition;

                    if (attachment.ContentID != null)
                    {
                        contentDisposition = new ContentDisposition(ContentDispositionType.Inline);
                    }

                    if (attachment.Name != null)
                    {
                        Parameter nameParameter = new Parameter("filename", "\"" + attachment.Name + "\"");
                        contentDisposition.Parameters.Add(nameParameter);
                    }

                    this.ContentDisposition = contentDisposition;
                }
                else
                {
                    ContentDisposition contentDisposition = new ContentDisposition(ContentDispositionType.Attachment);

                    if (attachment.ContentID != null)
                    {
                        contentDisposition = new ContentDisposition(ContentDispositionType.Inline);
                    }

                    if (attachment.Name != null)
                    {
                        Parameter nameParameter = new Parameter("filename", "\"" + attachment.Name + "\"");
                        contentDisposition.Parameters.Add(nameParameter);
                    }

                    this.ContentDisposition = contentDisposition;
                }

                byte[] attachmentBuffer = attachment.GetBytes();

                if (attachmentBuffer != null && this.ContentType != null && this.ContentType.Type == "text")
                {
                    System.Text.Encoding encoding = Util.GetEncoding(this.ContentType.CharSet);

                    body = encoding.GetString(attachmentBuffer, 0, attachmentBuffer.Length);
                }
                else if (attachmentBuffer != null)
                {
                    body = Convert.ToBase64String(attachmentBuffer, 0, attachmentBuffer.Length);
                    body = Util.SplitBase64Body(body);
                }
            }
        }

        internal BodyPart(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            int headerIndex = Util.IndexOfBuffer(buffer, Util.CrlfBuffer);

            if (headerIndex > -1)
            {
                byte[] headerBuffer = new byte[headerIndex];
                System.Array.Copy(buffer, 0, headerBuffer, 0, headerBuffer.Length);

                ParseHeader(headerBuffer);
            }

            /////////////// This part of code is same in Message and BodyPart classes

            if (buffer.Length > headerIndex + 4)
            {
                byte[] bodyBuffer = new byte[buffer.Length - headerIndex - 4];
                System.Array.Copy(buffer, headerIndex + 4, bodyBuffer, 0, bodyBuffer.Length);

                ContentType contentType = this.ContentType;
                ContentDisposition contentDisposition = this.ContentDisposition;
                ContentTransferEncoding contentTransferEncoding = this.ContentTransferEncoding;

                if (contentType == null)
                {
                    body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);
                }
                else if ((contentType.Type != null && contentType.Type.ToLower() == "text") ||
                         (contentType.Type != null && contentType.Type.ToLower() == "text" && contentDisposition != null && contentDisposition.Type == ContentDispositionType.Inline))
                {
                    Parameter charsetParameter = contentType.Parameters["charset"];

                    if (charsetParameter != null && charsetParameter.Value != null)
                    {
                        Encoding encoding = Util.GetEncoding(charsetParameter.Value);
                        body = encoding.GetString(bodyBuffer, 0, bodyBuffer.Length);

                        if (contentTransferEncoding == ContentTransferEncoding.QuotedPrintable)
                        {
                            body = Util.DecodeQuotedPrintable(body, encoding);
                        }
                        else if (contentTransferEncoding == ContentTransferEncoding.Base64)
                        {
                            body = Util.DecodeBase64(body, encoding);
                        }
                        else
                        {
                            body = encoding.GetString(bodyBuffer, 0, bodyBuffer.Length);
                        }
                    }
                    else
                    {
                        body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);

                        if (contentTransferEncoding == ContentTransferEncoding.QuotedPrintable)
                        {
                            body = Util.DecodeQuotedPrintable(body, System.Text.Encoding.Default);
                        }
                        else if (contentTransferEncoding == ContentTransferEncoding.Base64)
                        {
                            body = Util.DecodeBase64(body, System.Text.Encoding.Default);
                        }
                    }
                }
                else if (contentType.Type != null && contentType.Type.ToLower() == "multipart")
                {
                    Parameter boundaryParameter = contentType.Parameters["boundary"];

                    if (boundaryParameter != null && boundaryParameter.Value != null)
                    {
                        string boundary = boundaryParameter.Value;

                        IList<byte[]> bodyPartsBufferList = Util.GetBodyPartsBuffers(bodyBuffer, boundary);

                        for (int i = 0; i < bodyPartsBufferList.Count; i++)
                        {
                            byte[] bodyPartBuffer = bodyPartsBufferList[i];

                            BodyPart bodyPart = new BodyPart(bodyPartBuffer);
                            bodyParts.Add(bodyPart);
                        }
                    }
                }
                else if (contentType.Type != null && contentType.Type.ToLower() == "message")
                {
                    if (contentType.SubType != null && contentType.SubType.ToLower() == "rfc822")
                    {
                        embeddedMessage = new Message(bodyBuffer);
                    }
                    else if (contentType.SubType != null && contentType.SubType.ToLower() == "delivery-status")
                    {
                        byte[] encodedBodyBuffer = new byte[0];

                        if (contentTransferEncoding == ContentTransferEncoding.Base64)
                        {
                            String base64String = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);

                            if (base64String != null && base64String.Length > 0)
                            {
                                encodedBodyBuffer = Convert.FromBase64String(base64String);
                            }
                        }

                        if (encodedBodyBuffer.Length > 0)
                        {
                            try
                            {
                                embeddedMessage = new Message(encodedBodyBuffer);
                                embeddedMessage.Subject = contentType.Parameters["name"] != null ? contentType.Parameters["name"].Value : "ATT";
                            }
                            catch (MessageFormatException)
                            {
                                body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);
                            }
                        }
                    }
                    else if (contentType.SubType != null && contentType.SubType == "partial")
                    {
                        body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);
                    }
                }
                else //contentType = image, application, video, audio, ...
                {
                    if (contentTransferEncoding == ContentTransferEncoding.QuotedPrintable)
                    {
                        body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);
                        body = body.Replace("=09", "\t");
                        body = body.Replace("=0D", "\r");
                        body = body.Replace("=0A", "\n");
                        body = body.Replace("=20", " ");
                        body = body.Replace("=\r\n", String.Empty);
                    }
                    else if (contentTransferEncoding == ContentTransferEncoding.Base64)
                    {
                        bodyBuffer = Util.RemoveCRLF(bodyBuffer);

                        body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);             
                    }
                    else
                    {
                        body = System.Text.Encoding.Default.GetString(bodyBuffer, 0, bodyBuffer.Length);
                    }
                }
            }

            ///////////////////////////////////////////////////////////////////////////////////////
        }

        private void ParseHeader(byte[] headerBuffer)
        {
            string headerAsciiString = System.Text.Encoding.Default.GetString(headerBuffer, 0, headerBuffer.Length);

            headerAsciiString = Util.Unfolding(headerAsciiString);

            string[] lines = headerAsciiString.Split(Util.CrlfSeparator);

            string currentHeaderName = null;
            string currentHeaderValue = null;

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Length > 0)
                {
                    if (!lines[i].StartsWith(" ") && !lines[i].StartsWith("\t"))
                    {
                        if (currentHeaderName != null && currentHeaderValue != null)
                        {
                            currentHeaderValue = currentHeaderValue.TrimStart();
                            currentHeaderValue = Util.DecodeHeader(currentHeaderValue, ref headerEncoding);

                            if (Util.IsStandardKey(currentHeaderName))
                            {
                                currentHeaderName = Util.GetCorrectedStandardKey(currentHeaderName);
                            }

                            Header header = new Header(currentHeaderName, currentHeaderValue);
                            headers.Add(header);

                            currentHeaderName = null;
                            currentHeaderValue = null;
                        }

                        int colonIndex = lines[i].IndexOf(":");

                        if (colonIndex > -1)
                        {
                            currentHeaderName = lines[i].Substring(0, colonIndex);

                            if (currentHeaderName.ToLower() == "content-disposition")
                            {
                                currentHeaderValue = lines[i].Substring(colonIndex + 1);

                                //parse Content-Disposition
                                ContentDisposition cd = new ContentDisposition(currentHeaderValue);
                                currentHeaderValue = cd.ToString();
                            }
                            else
                            {
                                currentHeaderValue = lines[i].Substring(colonIndex + 1);
                            }
                        }
                    }
                    else
                    {
                        lines[i] = lines[i].Replace("\t", " ");
                        currentHeaderValue += lines[i];
                    }
                }
            }

            if (currentHeaderName != null && currentHeaderValue != null)
            {
                currentHeaderValue = currentHeaderValue.TrimStart();
                currentHeaderValue = Util.DecodeHeader(currentHeaderValue, ref headerEncoding);

                if (Util.IsStandardKey(currentHeaderName))
                {
                    currentHeaderName = Util.GetCorrectedStandardKey(currentHeaderName);
                }

                Header header = new Header(currentHeaderName, currentHeaderValue);
                headers.Add(header);
            }
        }

        internal Attachment[] GetAttachments(bool includeEmbedded)
        {
            IList<Attachment> attachmentList = new List<Attachment>();

            if (embeddedMessage != null)
            {
                string attachmentName = null;
                string charset = null;

                if (embeddedMessage.ContentType != null)
                {
                    Parameter nameParameter = embeddedMessage.ContentType.Parameters["name"];

                    if (nameParameter != null && nameParameter.Value != null)
                    {
                        attachmentName = nameParameter.Value;
                    }

                    Parameter charsetParameter = embeddedMessage.ContentType.Parameters["charset"];

                    if (charsetParameter != null)
                    {
                        charset = charsetParameter.Value;
                    }
                }

                if (attachmentName == null && embeddedMessage.ContentDisposition != null)
                {
                    Parameter filenameParameter = embeddedMessage.ContentDisposition.Parameters["filename"];

                    if (filenameParameter != null && filenameParameter.Value != null)
                    {
                        attachmentName = filenameParameter.Value;
                    }
                }

                if (attachmentName == null)
                {
                    attachmentName = embeddedMessage.Subject + ".eml";
                }

                Encoding encoding = Util.GetEncoding(charset);

                byte[] attachmentBuffer = encoding.GetBytes(embeddedMessage.ToString());

                Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                attachment.ContentID = this.ContentID;
                attachment.ContentLocation = this.ContentLocation;
                attachment.ContentDisposition = this.ContentDisposition;
                attachment.ContentDescription = this.ContentDescription;
                attachmentList.Add(attachment);
            } 
            else if (ContentType != null && ContentType.Type != null && ContentType.Type.ToLower() == "message" && ContentType.SubType != null && ContentType.SubType.ToLower() == "rfc822")
            {
                string attachmentName = null;
                string charset = null;

                if (this.ContentType != null)
                {
                    Parameter nameParameter = this.ContentType.Parameters["name"];

                    if (nameParameter != null && nameParameter.Value != null)
                    {
                        attachmentName = nameParameter.Value;
                    }

                    Parameter charsetParameter = this.ContentType.Parameters["charset"];

                    if (charsetParameter != null)
                    {
                        charset = charsetParameter.Value;
                    }
                }

                if (attachmentName == null && this.ContentDisposition != null)
                {
                    Parameter filenameParameter = this.ContentDisposition.Parameters["filename"];

                    if (filenameParameter != null && filenameParameter.Value != null)
                    {
                        attachmentName = filenameParameter.Value;
                    }
                }

                Encoding encoding = Util.GetEncoding(charset);

                byte[] attachmentBuffer = encoding.GetBytes(this.ToString());

                Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                attachment.ContentID = this.ContentID;
                attachment.ContentLocation = this.ContentLocation;
                attachment.ContentDisposition = this.ContentDisposition;
                attachment.ContentDescription = this.ContentDescription;
                attachmentList.Add(attachment);
            }
            else if (body != null && ContentType != null && ContentType.Type != null && ContentType.Type.ToLower() != "text")
            {
                string attachmentName = null;

                if (this.ContentType != null)
                {
                    Parameter nameParameter = this.ContentType.Parameters["name"];

                    if (nameParameter != null && nameParameter.Value != null)
                    {
                        attachmentName = nameParameter.Value;
                    }
                }

                if (attachmentName == null && this.ContentDisposition != null)
                {
                    Parameter filenameParameter = this.ContentDisposition.Parameters["filename"];

                    if (filenameParameter != null && filenameParameter.Value != null)
                    {
                        attachmentName = filenameParameter.Value;
                    }
                }

                if (this.ContentTransferEncoding == ContentTransferEncoding.Base64)
                {
                    byte[] attachmentBuffer = Util.DecodeBase64Attachment(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachment.ContentType = new ContentType(ContentType.Type, ContentType.SubType);
                    attachment.ContentID = this.ContentID;
                    attachment.ContentLocation = this.ContentLocation;
                    attachment.ContentDisposition = this.ContentDisposition;
                    attachment.ContentDescription = this.ContentDescription;
                    attachmentList.Add(attachment);
                }
                else
                {
                    byte[] attachmentBuffer = System.Text.Encoding.Default.GetBytes(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachment.ContentType = new ContentType(ContentType.Type, ContentType.SubType);
                    attachment.ContentID = this.ContentID;
                    attachment.ContentLocation = this.ContentLocation;
                    attachment.ContentDisposition = this.ContentDisposition;
                    attachment.ContentDescription = this.ContentDescription;
                    attachmentList.Add(attachment);
                }
            }
            else if (body != null && ContentType != null && ContentType.Type != null && ContentType.Type.ToLower() == "text" && ContentType.Parameters["name"] != null)
            {
                string attachmentName = null;

                if (this.ContentType != null)
                {
                    Parameter nameParameter = this.ContentType.Parameters["name"];

                    if (nameParameter != null && nameParameter.Value != null)
                    {
                        attachmentName = nameParameter.Value;
                    }
                }

                if (attachmentName == null && this.ContentDisposition != null)
                {
                    Parameter filenameParameter = this.ContentDisposition.Parameters["filename"];

                    if (filenameParameter != null && filenameParameter.Value != null)
                    {
                        attachmentName = filenameParameter.Value;
                    }
                }

                byte[] attachmentBuffer = System.Text.Encoding.Default.GetBytes(body);

                Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                attachment.ContentType = new ContentType(ContentType.Type, ContentType.SubType);
                attachment.ContentID = this.ContentID;
                attachment.ContentLocation = this.ContentLocation;
                attachment.ContentDisposition = this.ContentDisposition;
                attachment.ContentDescription = this.ContentDescription;
                attachmentList.Add(attachment);
                
            }
            else if (body != null && ContentDisposition != null && ContentDisposition.ToString().ToLower() != "inline")
            {
                string attachmentName = null;

                if (this.ContentType != null)
                {
                    Parameter nameParameter = this.ContentType.Parameters["name"];

                    if (nameParameter != null && nameParameter.Value != null)
                    {
                        attachmentName = nameParameter.Value;
                    }
                }

                if (attachmentName == null && this.ContentDisposition != null)
                {
                    Parameter filenameParameter = this.ContentDisposition.Parameters["filename"];

                    if (filenameParameter != null && filenameParameter.Value != null)
                    {
                        attachmentName = filenameParameter.Value;
                    }
                }

                if (this.ContentTransferEncoding == ContentTransferEncoding.Base64)
                {
                    byte[] attachmentBuffer = Util.DecodeBase64Attachment(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachment.ContentID = this.ContentID;
                    attachment.ContentLocation = this.ContentLocation;
                    attachment.ContentDisposition = this.ContentDisposition;
                    attachment.ContentDescription = this.ContentDescription;
                    attachmentList.Add(attachment);
                }
                else
                {
                    byte[] attachmentBuffer = System.Text.Encoding.Default.GetBytes(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachment.ContentID = this.ContentID;
                    attachment.ContentLocation = this.ContentLocation;
                    attachment.ContentDisposition = this.ContentDisposition;
                    attachment.ContentDescription = this.ContentDescription;
                    attachmentList.Add(attachment);
                }
            }
            else if (bodyParts.Count > 0)
            {
                for (int i = 0; i < bodyParts.Count; i++)
                {
                    if (bodyParts[i] != null && (includeEmbedded || bodyParts[i].ContentID == null) && bodyParts[i].ContentType != null && ContentType.Type.ToLower() != "text")
                    {
                        Attachment[] bodyPartAttachments = bodyParts[i].GetAttachments(includeEmbedded);

                        for (int j = 0; j < bodyPartAttachments.Length; j++)
                        {
                            attachmentList.Add(bodyPartAttachments[j]);
                        }
                    }
                }
            }

            Attachment[] attachments = new Attachment[attachmentList.Count];

            for (int i = 0; i < attachmentList.Count; i++)
            {
                attachments[i] = attachmentList[i];
            }

            return attachments;
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false);
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
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
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public void Save(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            byte[] buffer = GetBytes();
            stream.Write(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Gets the bytes.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetBytes()
        {
            byte[] buffer = null;
            MemoryStream memoryStream = (MemoryStream)GetStream();

            using (memoryStream)
            {
                buffer = memoryStream.ToArray();
            }

            return buffer;
        }

        /// <summary>
        /// Gets the stream.
        /// </summary>
        /// <returns>Stream.</returns>
        public Stream GetStream()
        {
            return GetStreamImplementation();
        }

        private Stream GetStreamImplementation()
        {
            MemoryStream memoryStream = new MemoryStream();

            string headerString = "";

            //Content type
            ContentType contentType = this.ContentType;
            string boundary = null;

            if (contentType != null && contentType.Type != null && contentType.Type.ToLower() == "multipart")
            {
                contentType.Parameters.Remove("boundary");

                boundary = "----=_NextPart_" + Util.NewGuid();
                contentType.Parameters.Add(new Parameter("boundary", boundary));

                this.ContentType = contentType;
            }
            else if (contentType == null && bodyParts.Count > 0)
            {
                contentType = new ContentType("multipart", "mixed");

                boundary = "----=_NextPart_" + Util.NewGuid();
                contentType.Parameters.Add(new Parameter("boundary", boundary));

                this.ContentType = contentType;
            }
            else if (contentType == null)
            {
                contentType = new ContentType("text", "plain");
                this.ContentType = contentType;
            }

            //Header
            for (int i = 0; i < headers.Count; i++)
            {
                Header header = headers[i];

                if (header != null && header.Name != null && header.Value != null)
                {
                    if (Util.HasToEncodeHeader(header.Name, header.Value))
                    {
                        string doNotEncodePrefix = "";
                        string doNotEncodeSuffix = "";

                        string headerValueToEncode = header.Value;

                        //do not encode content-type and content-disposition headers
                        if (header.Name.ToLower() == "content-type" || header.Name.ToLower() == "content-disposition")
                        {
                            int semiColonIndex = headerValueToEncode.IndexOf(";");

                            if (semiColonIndex > -1)
                            {
                                doNotEncodePrefix = headerValueToEncode.Substring(0, semiColonIndex + 1);
                                headerValueToEncode = headerValueToEncode.Substring(semiColonIndex + 1, header.Value.Length - semiColonIndex - 1);
                            }

                            int fileNameIndex = headerValueToEncode.IndexOf("filename=");

                            if (fileNameIndex > -1)
                            {
                                int nextSemiColonIndex = headerValueToEncode.IndexOf(";", fileNameIndex);

                                if (nextSemiColonIndex == -1)
                                {
                                    nextSemiColonIndex = headerValueToEncode.Length;
                                }

                                doNotEncodePrefix += headerValueToEncode.Substring(0, 10);
                                headerValueToEncode = headerValueToEncode.Substring(10, nextSemiColonIndex - 10);

                                headerValueToEncode = headerValueToEncode.Trim(new char[] { '"' });
                            }

                            int nameIndex = headerValueToEncode.IndexOf("name=");

                            if (nameIndex > -1)
                            {
                                int nextSemiColonIndex = headerValueToEncode.IndexOf(";", nameIndex);

                                if (nextSemiColonIndex == -1)
                                {
                                    nextSemiColonIndex = headerValueToEncode.Length;
                                }

                                doNotEncodePrefix += headerValueToEncode.Substring(0, 6);
                                headerValueToEncode = headerValueToEncode.Substring(6, nextSemiColonIndex - 6);

                                headerValueToEncode = headerValueToEncode.Trim(new char[] { '"' });
                            }
                        }
                        
                        string encodedHeader = Util.EncodeHeader(headerValueToEncode, this.headerCharSet, headerEncoding);

                        encodedHeader = doNotEncodePrefix + encodedHeader + doNotEncodeSuffix;

                        string headerRowString = header.Name + ": " + encodedHeader;

                        headerRowString = Util.SplitHeaderLines(headerRowString, 76);

                        headerString += headerRowString + "\r\n";
                        
                    }
                    else
                    {
                        string headerRowString = header.Name + ": " + header.Value;

                        headerRowString = Util.SplitHeaderLines(headerRowString, 76);

                        headerString += headerRowString + "\r\n";
                    }
                }
            }

            headerString += "\r\n";

            byte[] headerBuffer = System.Text.Encoding.Default.GetBytes(headerString);
            memoryStream.Write(headerBuffer, 0, headerBuffer.Length);

            if (bodyParts.Count > 0)
            {
                string delimiter = "--" + boundary;
                string lastDelimiter = "--" + boundary + "--";

                for (int i = 0; i < bodyParts.Count; i++)
                {
                    string delimiterString = "\r\n" + delimiter + "\r\n";
                    byte[] delimiterBuffer = System.Text.Encoding.Default.GetBytes(delimiterString);
                    memoryStream.Write(delimiterBuffer, 0, delimiterBuffer.Length);

                    BodyPart bodyPart = bodyParts[i];
                    byte[] bodyPartBuffer = bodyPart.GetBytes();
                    memoryStream.Write(bodyPartBuffer, 0, bodyPartBuffer.Length);
                }

                string lastDelimiterString = "\r\n" + lastDelimiter + "\r\n";
                byte[] lastDelimiterBuffer = System.Text.Encoding.Default.GetBytes(lastDelimiterString);
                memoryStream.Write(lastDelimiterBuffer, 0, lastDelimiterBuffer.Length);

            }
            else if (contentType != null && contentType.Type != null && contentType.Type.ToLower() == "message" && contentType.SubType != null && contentType.SubType.ToLower() == "rfc822" && embeddedMessage != null)
            {
                byte[] embeddedMessageBuffer = embeddedMessage.GetBytes();
                memoryStream.Write(embeddedMessageBuffer, 0, embeddedMessageBuffer.Length);
            }
            else if (contentType != null && contentType.Type != null && contentType.Type.ToLower() == "text" && body != null)
            {
                Encoding encoding = System.Text.Encoding.Default;
                ContentTransferEncoding contentTransferEncoding = ContentTransferEncoding.SevenBit;

                if (contentType != null)
                {
                    Parameter charsetParameter = contentType.Parameters["charset"];

                    if (charsetParameter != null && charsetParameter.Value != null)
                    {
                        encoding = Util.GetEncoding(charsetParameter.Value);
                    }
                }

                Header contentTransferEncodingHeader = headers[StandardHeader.ContentTransferEncoding];

                if (contentTransferEncodingHeader != null && contentTransferEncodingHeader.Value != null)
                {
                    contentTransferEncoding = EnumUtil.ParseContentTransferEncoding(contentTransferEncodingHeader.Value);
                }

                if (contentTransferEncoding == ContentTransferEncoding.QuotedPrintable)
                {
                    string encodedBody = Util.EncodeBodyQuotedPrintable(body, encoding);

                    encodedBody = Util.SplitQuotedPrintableBody(encodedBody);

                    byte[] encodedBodyBuffer = System.Text.Encoding.Default.GetBytes(encodedBody);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
                else if (contentTransferEncoding == ContentTransferEncoding.Base64)
                {
                    byte[] bodyBuffer = encoding.GetBytes(body);

                    string encodedBody = Convert.ToBase64String(bodyBuffer, 0, bodyBuffer.Length);

                    if (encodedBody.IndexOf("\r\n") == -1)
                    {
                        encodedBody = Util.SplitBase64Body(encodedBody);
                    }

                    byte[] encodedBodyBuffer = System.Text.Encoding.Default.GetBytes(encodedBody);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
                else
                {
                    byte[] encodedBodyBuffer = encoding.GetBytes(body);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
            }
            else if (body != null)
            {
                ContentTransferEncoding contentTransferEncoding = ContentTransferEncoding.SevenBit;

                Header contentTransferEncodingHeader = headers[StandardHeader.ContentTransferEncoding];

                if (contentTransferEncodingHeader != null && contentTransferEncodingHeader.Value != null)
                {
                    contentTransferEncoding = EnumUtil.ParseContentTransferEncoding(contentTransferEncodingHeader.Value);
                }

                if (contentTransferEncoding == ContentTransferEncoding.QuotedPrintable)
                {
                    string encodedBody = Util.SplitQuotedPrintableBody(body);

                    byte[] encodedBodyBuffer = System.Text.Encoding.Default.GetBytes(encodedBody);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
                else if (contentTransferEncoding == ContentTransferEncoding.Base64)
                {
                    string encodedBody = body;

                    if (encodedBody.IndexOf("\r\n") == -1)
                    {
                        encodedBody = Util.SplitBase64Body(encodedBody);
                    }

                    byte[] encodedBodyBuffer = System.Text.Encoding.Default.GetBytes(encodedBody);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
                else
                {
                    byte[] encodedBodyBuffer = System.Text.Encoding.Default.GetBytes(body);
                    memoryStream.Write(encodedBodyBuffer, 0, encodedBodyBuffer.Length);
                }
            }

            return memoryStream;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            byte[] buffer = GetBytes();
            string stringValue = System.Text.Encoding.UTF8.GetString(buffer, 0, buffer.Length);
            return stringValue;
        }

        #region Properties

        /// <summary>
        /// Gets the headers.
        /// </summary>
        /// <value>The headers.</value>
        public HeaderList Headers
        {
            get
            {
                return headers;
            }
        }

        /// <summary>
        /// Gets the body parts.
        /// </summary>
        /// <value>The body parts.</value>
        public BodyPartList BodyParts
        {
            get
            {
                return bodyParts;
            }
        }

        /// <summary>
        /// Gets or sets the body.
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
        /// Gets or sets the embedded message.
        /// </summary>
        /// <value>The embedded message.</value>
        public Message EmbeddedMessage
        {
            get
            {
                return embeddedMessage;
            }
            set
            {
                embeddedMessage = value;
            }
        }

        /// <summary>
        /// Gets or sets the header encoding.
        /// </summary>
        /// <value>The header encoding.</value>
        public HeaderEncoding HeaderEncoding
        {
            get
            {
                return headerEncoding;
            }
            set
            {
                headerEncoding = value;
            }
        }

        /// <summary>
        /// Gets or sets the header character set.
        /// </summary>
        /// <value>The header character set.</value>
        public string HeaderCharSet
        {
            get
            {
                return headerCharSet;
            }
            set
            {
                headerCharSet = value;
            }
        }

        /// <summary>
        /// Gets or sets the type of the content.
        /// </summary>
        /// <value>The type of the content.</value>
        public ContentType ContentType
        {
            get
            {
                Header contentTypeHeader = headers[StandardHeader.ContentType];

                if (contentTypeHeader != null && contentTypeHeader.Value != null)
                {
                    ContentType contentType = new ContentType(contentTypeHeader.Value);
                    return contentType;
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
                    headers.Remove(StandardHeader.ContentType);

                    Header contentTypeHeader = new Header(StandardHeader.ContentType, value.ToString());
                    headers.Add(contentTypeHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ContentType);
                }
            }
        }

        /// <summary>
        /// Gets or sets the content transfer encoding.
        /// </summary>
        /// <value>The content transfer encoding.</value>
        public ContentTransferEncoding ContentTransferEncoding
        {
            get
            {
                Header contentTransferEncodingHeader = headers[StandardHeader.ContentTransferEncoding];

                if (contentTransferEncodingHeader != null && contentTransferEncodingHeader.Value != null)
                {
                    ContentTransferEncoding contentTransferEncoding = EnumUtil.ParseContentTransferEncoding(contentTransferEncodingHeader.Value);
                    return contentTransferEncoding;
                }
                else
                {
                    return ContentTransferEncoding.SevenBit;
                }
            }
            set
            {
                headers.Remove(StandardHeader.ContentTransferEncoding);

                Header contentTransferEncodingHeader = new Header(StandardHeader.ContentTransferEncoding, EnumUtil.ParseContentTransferEncoding(value));
                headers.Add(contentTransferEncodingHeader);
            }
        }

        /// <summary>
        /// Gets or sets the content disposition.
        /// </summary>
        /// <value>The content disposition.</value>
        public ContentDisposition ContentDisposition
        {
            get
            {
                Header contentDispositionHeader = headers[StandardHeader.ContentDisposition];

                if (contentDispositionHeader != null && contentDispositionHeader.Value != null)
                {
                    ContentDisposition contentDisposition = new ContentDisposition(contentDispositionHeader.Value);
                    return contentDisposition;
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
                    headers.Remove(StandardHeader.ContentDisposition);

                    Header contentDispositionHeader = new Header(StandardHeader.ContentDisposition, value.ToString());
                    headers.Add(contentDispositionHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ContentDisposition);
                }
            }
        }

        /// <summary>
        /// Gets or sets the content description.
        /// </summary>
        /// <value>The content description.</value>
        public string ContentDescription
        {
            get
            {
                Header contentDescriptionHeader = headers[StandardHeader.ContentDescription];

                if (contentDescriptionHeader != null && contentDescriptionHeader.Value != null)
                {
                    return contentDescriptionHeader.Value;
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
                    headers.Remove(StandardHeader.ContentDescription);

                    Header contentDescriptionHeader = new Header(StandardHeader.ContentDescription, value);
                    headers.Add(contentDescriptionHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ContentDescription);
                }
            }
        }

        /// <summary>
        /// Gets or sets the content identifier.
        /// </summary>
        /// <value>The content identifier.</value>
        public string ContentID
        {
            get
            {
                Header contentIDHeader = headers[StandardHeader.ContentID];

                if (contentIDHeader != null && contentIDHeader.Value != null)
                {
                    return contentIDHeader.Value;
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
                    headers.Remove(StandardHeader.ContentID);

                    Header contentIDHeader = new Header(StandardHeader.ContentID, value);
                    headers.Add(contentIDHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ContentID);
                }
            }
        }

        /// <summary>
        /// Gets or sets the content location.
        /// </summary>
        /// <value>The content location.</value>
        public string ContentLocation
        {
            get
            {
                Header contentLocationHeader = headers[StandardHeader.ContentLocation];

                if (contentLocationHeader != null && contentLocationHeader.Value != null)
                {
                    return contentLocationHeader.Value;
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
                    headers.Remove(StandardHeader.ContentLocation);

                    Header contentLocationHeader = new Header(StandardHeader.ContentLocation, value);
                    headers.Add(contentLocationHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ContentLocation);
                }
            }
        }


        #endregion
    }
}