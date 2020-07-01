using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class Message.
    /// </summary>
    public class Message
    {
        private HeaderList headers = new HeaderList();
        private BodyPartList bodyParts = new BodyPartList();
        private string body;
        private Message embeddedMessage;
        private HeaderEncoding headerEncoding = HeaderEncoding.QuotedPrintable;
        private string headerCharSet = "utf-8";

        //mailbox collections
        private IList<Mailbox> to = new List<Mailbox>();
        private IList<Mailbox> cc = new List<Mailbox>();
        private IList<Mailbox> bcc = new List<Mailbox>();
        private IList<Mailbox> replyTo = new List<Mailbox>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class.
        /// </summary>
        public Message()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public Message(string filePath)
            : this()
        {
            Open(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public Message(Stream stream)
            : this()
        {
            Open(stream);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public Message(byte[] buffer)
            : this()
        {
            Open(buffer);
        }

        /// <summary>
        /// Opens the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public void Open(string filePath)
        {
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            using (fileStream)
            {
                Open(fileStream);
            }
        }

        /// <summary>
        /// Opens the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public void Open(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            byte[] buffer = new byte[32768];
            int count = -1;

            MemoryStream memoryStream = new MemoryStream();

            using (memoryStream)
            {
                while ((count = stream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    memoryStream.Write(buffer, 0, count);
                }

                Open(memoryStream.ToArray());
            }
        }

        /// <summary>
        /// Opens the specified buffer.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public void Open(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            int headerIndex = Util.IndexOfBuffer(buffer, Util.CrlfBuffer);

            if (headerIndex == -1)
            {
                //try to recover message
                buffer = Util.RecoverMessage(buffer);
                headerIndex = Util.IndexOfBuffer(buffer, Util.CrlfBuffer);
            }

            if (headerIndex == -1)
            {
                throw new MessageFormatException("Invalid message format.");
            }
            else
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
                string currentLine = lines[i];

                if (currentLine.Length > 0)
                {
                    if (!currentLine.StartsWith(" ") && !currentLine.StartsWith("\t"))
                    {
                        int colonIndex = currentLine.IndexOf(":");

                        if (colonIndex > -1)
                        {
                            currentHeaderName = currentLine.Substring(0, colonIndex);
                            currentHeaderValue = currentLine.Substring(colonIndex + 1);
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

                            currentHeaderName = null;
                            currentHeaderValue = null;
                        }
                    }
                    else
                    {
                        currentLine = currentLine.TrimStart(Util.TabAndSpaceSeparator);
                        currentHeaderValue += currentLine;
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

            headers = GetMailboxCollections(headers);
        }

        private HeaderList GetMailboxCollections(HeaderList headers)
        {
            Header toHeader = headers[StandardHeader.To];
            Header ccHeader = headers[StandardHeader.Cc];
            Header bccHeader = headers[StandardHeader.Bcc];
            Header replyToHeader = headers[StandardHeader.ReplyTo];

            if (toHeader != null && toHeader.Value != null)
            {
                to = Util.ParseMailboxes(toHeader.Value);
                headers.Remove(StandardHeader.To);
            }

            if (ccHeader != null && ccHeader.Value != null)
            {
                cc = Util.ParseMailboxes(ccHeader.Value);
                headers.Remove(StandardHeader.Cc);
            }

            if (bccHeader != null && bccHeader.Value != null)
            {
                bcc = Util.ParseMailboxes(bccHeader.Value);
                headers.Remove(StandardHeader.Bcc);
            }

            if (replyToHeader != null && replyToHeader.Value != null)
            {
                replyTo = Util.ParseMailboxes(replyToHeader.Value);
                headers.Remove(StandardHeader.ReplyTo);
            }

            return headers;
        }

        private HeaderList AddMailboxCollections(HeaderList headers)
        {
            string toString = "";
            string ccString = "";
            string bccString = "";
            string replyToString = "";

            for (int i = 0; i < to.Count; i++)
            {
                if (to[i] != null)
                {
                    toString += to[i].ToString();
                }

                if (i < to.Count - 1)
                {
                    toString += ", ";
                }
            }

            for (int i = 0; i < cc.Count; i++)
            {
                if (cc[i] != null)
                {
                    ccString += cc[i].ToString();
                }

                if (i < cc.Count - 1)
                {
                    ccString += ", ";
                }
            }

            for (int i = 0; i < bcc.Count; i++)
            {
                if (bcc[i] != null)
                {
                    bccString += bcc[i].ToString();
                }

                if (i < bcc.Count - 1)
                {
                    bccString += ", ";
                }
            }

            for (int i = 0; i < replyTo.Count; i++)
            {
                if (replyTo[i] != null)
                {
                    replyToString += replyTo[i].ToString();
                }

                if (i < replyTo.Count - 1)
                {
                    replyToString += ", ";
                }
            }

            if (toString.Length > 0)
            {
                Header toHeader = new Header(StandardHeader.To, toString);

                headers.Remove(StandardHeader.To);
                headers.Add(toHeader);
            }

            if (ccString.Length > 0)
            {
                Header ccHeader = new Header(StandardHeader.Cc, ccString);

                headers.Remove(StandardHeader.Cc);
                headers.Add(ccHeader);
            }

            if (bccString.Length > 0)
            {
                Header bccHeader = new Header(StandardHeader.Bcc, bccString);

                headers.Remove(StandardHeader.Bcc);
                headers.Add(bccHeader);
            }

            if (replyToString.Length > 0)
            {
                Header replyToHeader = new Header(StandardHeader.ReplyTo, replyToString);

                headers.Remove(StandardHeader.ReplyTo);
                headers.Add(replyToHeader);
            }

            return headers;
        }

        /// <summary>
        /// Gets the attachments.
        /// </summary>
        /// <returns>Attachment[][].</returns>
        public Attachment[] GetAttachments()
        {
            return GetAttachments(false);
        }

        /// <summary>
        /// Gets the attachments.
        /// </summary>
        /// <param name="includeEmbedded">if set to <c>true</c> [include embedded].</param>
        /// <returns>Attachment[][].</returns>
        public Attachment[] GetAttachments(bool includeEmbedded)
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
                attachmentList.Add(attachment);

            }
            else if (body != null && ((ContentDisposition != null && ContentDisposition.ToString().ToLower() != "inline") || (ContentType != null && ContentType.Type != null && ContentType.Type.ToLower() != "text")))
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
                    attachmentList.Add(attachment);
                }
                else
                {
                    byte[] attachmentBuffer = System.Text.Encoding.Default.GetBytes(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachmentList.Add(attachment);
                }
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
                    attachmentList.Add(attachment);
                }
                else
                {
                    byte[] attachmentBuffer = System.Text.Encoding.Default.GetBytes(body);

                    Attachment attachment = new Attachment(attachmentBuffer, attachmentName);
                    attachmentList.Add(attachment);
                }
            }
            else if (bodyParts.Count > 0)
            {
                for (int i = 0; i < bodyParts.Count; i++)
                {
                    Attachment[] bodyPartAttachments = bodyParts[i].GetAttachments(includeEmbedded);

                    for (int j = 0; j < bodyPartAttachments.Length; j++)
                    {
                        attachmentList.Add(bodyPartAttachments[j]);
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
        /// Gets the name of the file.
        /// </summary>
        /// <returns>System.String.</returns>
        public string GetFileName()
        {
            if (Subject == null)
            {
                return "Message-" + Util.NextRandom();
            }
            else
            {
                string fileName = Subject;

                fileName = fileName.Replace(":", "_");
                fileName = fileName.Replace("*", "_");
                fileName = fileName.Replace("\\", "_");
                fileName = fileName.Replace("/", "_");
                fileName = fileName.Replace("?", "_");
                fileName = fileName.Replace("\"", "_");
                fileName = fileName.Replace("<", "_");
                fileName = fileName.Replace(">", "_");
                fileName = fileName.Replace("|", "_");
                fileName = fileName.Replace("", "_");
                fileName = fileName.Replace("\r", "_");
                fileName = fileName.Replace("\n", "_");
                fileName = fileName.Replace("\t", "_");

                if (fileName.Length > 128)
                {
                    fileName = fileName.Substring(0, 128);
                }

                return fileName + ".eml";
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
                contentType.Parameters.Add(new Parameter("boundary", "\"" + boundary + "\""));

                this.ContentType = contentType;

                if (this.MimeVersion == null)
                {
                    this.MimeVersion = "1.0";
                }
            }
            else if (contentType == null && bodyParts.Count > 0)
            {
                contentType = new ContentType("multipart", "mixed");

                boundary = "----=_NextPart_" + Util.NewGuid();
                contentType.Parameters.Add(new Parameter("boundary", "\"" + boundary + "\""));

                this.ContentType = contentType;

                if (this.MimeVersion == null)
                {
                    this.MimeVersion = "1.0";
                }
            }
            else if (contentType == null)
            {
                contentType = new ContentType("text", "plain");
                this.ContentType = contentType;
            }

            headers = AddMailboxCollections(headers);

            for (int i = 0; i < headers.Count; i++)
            {
                Header header = headers[i];

                if (header != null && header.Name != null && header.Value != null)
                {
                    if (Util.HasToEncodeHeader(header.Name, header.Value))
                    {
                        string encodedHeader = Util.EncodeHeader(header.Value, this.headerCharSet, headerEncoding);

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

            /////////////// This part of code is same in Message and BodyPart classes

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

                    encodedBody = Util.SplitBase64Body(encodedBody);

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
                    string encodedBody = Util.SplitBase64Body(body);

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

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>The subject.</value>
        public string Subject
        {
            get
            {
                Header subjectHeader = headers[StandardHeader.Subject];

                if (subjectHeader != null && subjectHeader.Value != null)
                {
                    return subjectHeader.Value;
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
                    headers.Remove(StandardHeader.Subject);

                    Header subjectHeader = new Header(StandardHeader.Subject, value);
                    headers.Add(subjectHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.Subject);
                }
            }
        }

        /// <summary>
        /// Gets or sets the MIME version.
        /// </summary>
        /// <value>The MIME version.</value>
        public string MimeVersion
        {
            get
            {
                Header mimeVersionHeader = headers[StandardHeader.MimeVersion];

                if (mimeVersionHeader != null && mimeVersionHeader.Value != null)
                {
                    return mimeVersionHeader.Value;
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
                    headers.Remove(StandardHeader.MimeVersion);

                    Header mimeVersionHeader = new Header(StandardHeader.MimeVersion, value);
                    headers.Add(mimeVersionHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.MimeVersion);
                }
            }
        }

        /// <summary>
        /// Gets or sets the comments.
        /// </summary>
        /// <value>The comments.</value>
        public string Comments
        {
            get
            {
                Header commentsHeader = headers[StandardHeader.Comments];

                if (commentsHeader != null && commentsHeader.Value != null)
                {
                    return commentsHeader.Value;
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
                    headers.Remove(StandardHeader.Comments);

                    Header commentsHeader = new Header(StandardHeader.Comments, value);
                    headers.Add(commentsHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.Comments);
                }
            }
        }

        /// <summary>
        /// Gets or sets the keywords.
        /// </summary>
        /// <value>The keywords.</value>
        public string Keywords
        {
            get
            {
                Header keywordsHeader = headers[StandardHeader.Keywords];

                if (keywordsHeader != null && keywordsHeader.Value != null)
                {
                    return keywordsHeader.Value;
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
                    headers.Remove(StandardHeader.Keywords);

                    Header keywordsHeader = new Header(StandardHeader.Keywords, value);
                    headers.Add(keywordsHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.Keywords);
                }
            }
        }

        /// <summary>
        /// Gets or sets the message identifier.
        /// </summary>
        /// <value>The message identifier.</value>
        public string MessageID
        {
            get
            {
                Header messageIDHeader = headers[StandardHeader.MessageID];

                if (messageIDHeader != null && messageIDHeader.Value != null)
                {
                    return messageIDHeader.Value;
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
                    headers.Remove(StandardHeader.MessageID);

                    Header messageIDHeader = new Header(StandardHeader.MessageID, value);
                    headers.Add(messageIDHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.MessageID);
                }
            }
        }

        /// <summary>
        /// Gets or sets the resent message identifier.
        /// </summary>
        /// <value>The resent message identifier.</value>
        public string ResentMessageID
        {
            get
            {
                Header resentMessageIDHeader = headers[StandardHeader.ResentMessageID];

                if (resentMessageIDHeader != null && resentMessageIDHeader.Value != null)
                {
                    return resentMessageIDHeader.Value;
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
                    headers.Remove(StandardHeader.ResentMessageID);

                    Header resentMessageIDHeader = new Header(StandardHeader.ResentMessageID, value);
                    headers.Add(resentMessageIDHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ResentMessageID);
                }
            }
        }

        /// <summary>
        /// Gets or sets the date.
        /// </summary>
        /// <value>The date.</value>
        public DateTime Date
        {
            get
            {
                Header dateHeader = headers[StandardHeader.Date];

                if (dateHeader != null && dateHeader.Value != null)
                {
                    string dateString = dateHeader.Value;
                    int commentStartIndex = dateString.IndexOf("(");

                    if (commentStartIndex > -1)
                    {
                        dateString = dateString.Substring(0, commentStartIndex);
                    }

                    DateTime dateValue = DateTime.MinValue;

                    try
                    {
                        dateValue = DateTime.Parse(dateString, System.Globalization.CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        dateValue = Util.ParseRFC822Date(dateString);
                    }

                    return dateValue;
                }
                else
                {
                    return DateTime.MinValue;
                }
            }
            set
            {
                if (value.CompareTo(DateTime.MinValue) > 0)
                {
                    headers.Remove(StandardHeader.Date);

                    string dateTimeStringValue = value.ToString("ddd, dd MMM yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    string timeZoneStringValue = value.ToString("zzzz", System.Globalization.CultureInfo.InvariantCulture);

                    dateTimeStringValue = dateTimeStringValue + " " + timeZoneStringValue.Replace(":", "");

                    Header dateHeader = new Header(StandardHeader.Date, dateTimeStringValue);
                    headers.Add(dateHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.Date);
                }
            }
        }

        /// <summary>
        /// Gets or sets the resent date.
        /// </summary>
        /// <value>The resent date.</value>
        public DateTime ResentDate
        {
            get
            {
                Header resentDateHeader = headers[StandardHeader.ResentDate];

                if (resentDateHeader != null && resentDateHeader.Value != null)
                {
                    string dateString = resentDateHeader.Value;
                    int commentStartIndex = dateString.IndexOf("(");

                    if (commentStartIndex > -1)
                    {
                        dateString = dateString.Substring(0, commentStartIndex);
                    }

                    DateTime dateValue = DateTime.MinValue;

                    try
                    {
                        dateValue = DateTime.Parse(dateString, System.Globalization.CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                    }

                    return dateValue;
                }
                else
                {
                    return DateTime.MinValue;
                }
            }
            set
            {
                if (value.CompareTo(DateTime.MinValue) > 0)
                {
                    headers.Remove(StandardHeader.ResentDate);

                    string dateTimeStringValue = value.ToString("ddd, dd MMM yyyy HH:mm:ss");
                    string timeZoneStringValue = value.ToString("zzzz");

                    dateTimeStringValue = dateTimeStringValue + " " + timeZoneStringValue.Replace(":", "");

                    Header resentDateHeader = new Header(StandardHeader.ResentDate, dateTimeStringValue);
                    headers.Add(resentDateHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ResentDate);
                }
            }
        }

        /// <summary>
        /// Gets or sets the references.
        /// </summary>
        /// <value>The references.</value>
        public string References
        {
            get
            {
                Header referencesHeader = headers[StandardHeader.References];

                if (referencesHeader != null && referencesHeader.Value != null)
                {
                    return referencesHeader.Value;
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
                    headers.Remove(StandardHeader.References);

                    Header referencesHeader = new Header(StandardHeader.References, value);
                    headers.Add(referencesHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.References);
                }
            }
        }

        /// <summary>
        /// Gets to.
        /// </summary>
        /// <value>To.</value>
        public IList<Mailbox> To
        {
            get
            {
                return to;
            }
        }

        /// <summary>
        /// Gets the cc.
        /// </summary>
        /// <value>The cc.</value>
        public IList<Mailbox> Cc
        {
            get
            {
                return cc;
            }
        }

        /// <summary>
        /// Gets the BCC.
        /// </summary>
        /// <value>The BCC.</value>
        public IList<Mailbox> Bcc
        {
            get
            {
                return bcc;
            }
        }

        /// <summary>
        /// Gets the reply to.
        /// </summary>
        /// <value>The reply to.</value>
        public IList<Mailbox> ReplyTo
        {
            get
            {
                return replyTo;
            }
        }

        /// <summary>
        /// Gets or sets from.
        /// </summary>
        /// <value>From.</value>
        public Mailbox From
        {
            get
            {
                Header fromHeader = headers[StandardHeader.From];

                if (fromHeader != null && fromHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(fromHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.From);

                    Header fromHeader = new Header(StandardHeader.From, value.ToString());
                    headers.Add(fromHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.From);
                }
            }
        }

        /// <summary>
        /// Gets or sets the sender.
        /// </summary>
        /// <value>The sender.</value>
        public Mailbox Sender
        {
            get
            {
                Header senderHeader = headers[StandardHeader.Sender];

                if (senderHeader != null && senderHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(senderHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.Sender);

                    Header senderHeader = new Header(StandardHeader.Sender, value.ToString());
                    headers.Add(senderHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.Sender);
                }
            }
        }

        /// <summary>
        /// Gets or sets the resent from.
        /// </summary>
        /// <value>The resent from.</value>
        public Mailbox ResentFrom
        {
            get
            {
                Header resentFromHeader = headers[StandardHeader.ResentFrom];

                if (resentFromHeader != null && resentFromHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(resentFromHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.ResentFrom);

                    Header resentFromHeader = new Header(StandardHeader.ResentFrom, value.ToString());
                    headers.Add(resentFromHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ResentFrom);
                }
            }
        }

        /// <summary>
        /// Gets or sets the resent sender.
        /// </summary>
        /// <value>The resent sender.</value>
        public Mailbox ResentSender
        {
            get
            {
                Header resentSenderHeader = headers[StandardHeader.ResentSender];

                if (resentSenderHeader != null && resentSenderHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(resentSenderHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.ResentSender);

                    Header resentSenderHeader = new Header(StandardHeader.ResentSender, value.ToString());
                    headers.Add(resentSenderHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ResentSender);
                }
            }
        }

        /// <summary>
        /// Gets or sets the in reply to.
        /// </summary>
        /// <value>The in reply to.</value>
        public Mailbox InReplyTo
        {
            get
            {
                Header inReplyToHeader = headers[StandardHeader.InReplyTo];

                if (inReplyToHeader != null && inReplyToHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(inReplyToHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.InReplyTo);

                    Header inReplyToHeader = new Header(StandardHeader.InReplyTo, value.ToString());
                    headers.Add(inReplyToHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.InReplyTo);
                }
            }
        }

        /// <summary>
        /// Gets or sets the return path.
        /// </summary>
        /// <value>The return path.</value>
        public Mailbox ReturnPath
        {
            get
            {
                Header returnPathHeader = headers[StandardHeader.ReturnPath];

                if (returnPathHeader != null && returnPathHeader.Value != null)
                {
                    Mailbox mailbox = new Mailbox(returnPathHeader.Value);
                    return mailbox;
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
                    headers.Remove(StandardHeader.ReturnPath);

                    Header returnPathHeader = new Header(StandardHeader.ReturnPath, value.ToString());
                    headers.Add(returnPathHeader);
                }
                else
                {
                    headers.Remove(StandardHeader.ReturnPath);
                }
            }
        }

        #endregion
    }
}