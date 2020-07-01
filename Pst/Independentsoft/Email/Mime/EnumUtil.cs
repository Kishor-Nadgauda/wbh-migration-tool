using System;

namespace Independentsoft.Email.Mime
{
    internal class EnumUtil
    {
        internal static HeaderEncoding ParseHeaderEncoding(string encoding)
        {
            if (encoding == null)
            {
                return HeaderEncoding.Binary;
            }

            encoding = encoding.ToUpper();

            if (encoding == "Q")
            {
                return HeaderEncoding.QuotedPrintable;
            }
            else
            {
                return HeaderEncoding.Binary;
            }
        }

        internal static string ParseHeaderEncoding(HeaderEncoding encoding)
        {
            if (encoding == HeaderEncoding.QuotedPrintable)
            {
                return "Q";
            }
            else
            {
                return "B";
            }
        }

        internal static ContentTransferEncoding ParseContentTransferEncoding(string encoding)
        {
            if (encoding == null)
            {
                return ContentTransferEncoding.SevenBit;
            }

            encoding = encoding.ToLower();

            if (encoding == "base64")
            {
                return ContentTransferEncoding.Base64;
            }
            else if (encoding == "quoted-printable")
            {
                return ContentTransferEncoding.QuotedPrintable;
            }
            else if (encoding == "8bit")
            {
                return ContentTransferEncoding.EightBit;
            }
            else if (encoding == "binary")
            {
                return ContentTransferEncoding.Binary;
            }
            else
            {
                return ContentTransferEncoding.SevenBit;
            }
        }

        internal static string ParseContentTransferEncoding(ContentTransferEncoding encoding)
        {
            if (encoding == ContentTransferEncoding.Base64)
            {
                return "base64";
            }
            else if (encoding == ContentTransferEncoding.QuotedPrintable)
            {
                return "quoted-printable";
            }
            else if (encoding == ContentTransferEncoding.EightBit)
            {
                return "8bit";
            }
            else if (encoding == ContentTransferEncoding.Binary)
            {
                return "binary";
            }
            else
            {
                return "7bit";
            }
        }

        internal static string ParseContentDispositionType(ContentDispositionType dispositionType)
        {
            if (dispositionType == ContentDispositionType.Inline)
            {
                return "inline";
            }
            else
            {
                return "attachment";
            }
        }

        internal static ContentDispositionType ParseContentDispositionType(string dispositionType)
        {
            if (dispositionType == null)
            {
                return ContentDispositionType.Attachment;
            }

            dispositionType = dispositionType.ToLower().Trim();

            if (dispositionType == "inline")
            {
                return ContentDispositionType.Inline;
            }
            else
            {
                return ContentDispositionType.Attachment;
            }
        }

        internal static string ParseStandardHeader(StandardHeader header)
        {
            if (header == StandardHeader.ResentDate)
            {
                return "Resent-Date";
            }
            else if (header == StandardHeader.ResentFrom)
            {
                return "Resent-From";
            }
            else if (header == StandardHeader.ResentSender)
            {
                return "Resent-Sender";
            }
            else if (header == StandardHeader.ResentTo)
            {
                return "Resent-To";
            }
            else if (header == StandardHeader.ResentCc)
            {
                return "Resent-Cc";
            }
            else if (header == StandardHeader.ResentBcc)
            {
                return "Resent-Bcc";
            }
            else if (header == StandardHeader.ResentMessageID)
            {
                return "Resent-Message-ID";
            }
            else if (header == StandardHeader.From)
            {
                return "From";
            }
            else if (header == StandardHeader.Sender)
            {
                return "Sender";
            }
            else if (header == StandardHeader.ReplyTo)
            {
                return "Reply-To";
            }
            else if (header == StandardHeader.To)
            {
                return "To";
            }
            else if (header == StandardHeader.Cc)
            {
                return "Cc";
            }
            else if (header == StandardHeader.Bcc)
            {
                return "Bcc";
            }
            else if (header == StandardHeader.MessageID)
            {
                return "Message-ID";
            }
            else if (header == StandardHeader.InReplyTo)
            {
                return "In-Reply-To";
            }
            else if (header == StandardHeader.References)
            {
                return "References";
            }
            else if (header == StandardHeader.Subject)
            {
                return "Subject";
            }
            else if (header == StandardHeader.Comments)
            {
                return "Comments";
            }
            else if (header == StandardHeader.Keywords)
            {
                return "Keywords";
            }
            else if (header == StandardHeader.Date)
            {
                return "Date";
            }
            else if (header == StandardHeader.ReturnPath)
            {
                return "Return-Path";
            }
            else if (header == StandardHeader.Received)
            {
                return "Received";
            }
            else if (header == StandardHeader.MimeVersion)
            {
                return "MIME-Version";
            }
            else if (header == StandardHeader.ContentType)
            {
                return "Content-Type";
            }
            else if (header == StandardHeader.ContentID)
            {
                return "Content-ID";
            }
            else if (header == StandardHeader.ContentTransferEncoding)
            {
                return "Content-Transfer-Encoding";
            }
            else if (header == StandardHeader.ContentDescription)
            {
                return "Content-Description";
            }
            else if (header == StandardHeader.ContentDisposition)
            {
                return "Content-Disposition";
            }
            else if (header == StandardHeader.ContentLocation)
            {
                return "Content-Location";
            }
            else
            {
                return "ContentLength";
            }
        }

        internal static StandardHeader ParseStandardHeader(string header)
        {
            if (header == "Resent-Date")
            {
                return StandardHeader.ResentDate;
            }
            else if (header == "Resent-From")
            {
                return StandardHeader.ResentFrom;
            }
            else if (header == "Resent-Sender")
            {
                return StandardHeader.ResentSender;
            }
            else if (header == "Resent-To")
            {
                return StandardHeader.ResentTo;
            }
            else if (header == "Resent-Cc")
            {
                return StandardHeader.ResentCc;
            }
            else if (header == "Resent-Bcc")
            {
                return StandardHeader.ResentBcc;
            }
            else if (header == "Resent-Message-ID")
            {
                return StandardHeader.ResentMessageID;
            }
            else if (header == "From")
            {
                return StandardHeader.From;
            }
            else if (header == "Sender")
            {
                return StandardHeader.Sender;
            }
            else if (header == "Reply-To")
            {
                return StandardHeader.ReplyTo;
            }
            else if (header == "To")
            {
                return StandardHeader.To;
            }
            else if (header == "Cc")
            {
                return StandardHeader.Cc;
            }
            else if (header == "Bcc")
            {
                return StandardHeader.Bcc;
            }
            else if (header == "Message-ID")
            {
                return StandardHeader.MessageID;
            }
            else if (header == "In-Reply-To")
            {
                return StandardHeader.InReplyTo;
            }
            else if (header == "References")
            {
                return StandardHeader.References;
            }
            else if (header == "Subject")
            {
                return StandardHeader.Subject;
            }
            else if (header == "Comments")
            {
                return StandardHeader.Comments;
            }
            else if (header == "Keywords")
            {
                return StandardHeader.Keywords;
            }
            else if (header == "Date")
            {
                return StandardHeader.Date;
            }
            else if (header == "Return-Path")
            {
                return StandardHeader.ReturnPath;
            }
            else if (header == "Received")
            {
                return StandardHeader.Received;
            }
            else if (header == "MIME-Version")
            {
                return StandardHeader.MimeVersion;
            }
            else if (header == "Content-Type")
            {
                return StandardHeader.ContentType;
            }
            else if (header == "Content-ID")
            {
                return StandardHeader.ContentID;
            }
            else if (header == "Content-Transfer-Encoding")
            {
                return StandardHeader.ContentTransferEncoding;
            }
            else if (header == "Content-Description")
            {
                return StandardHeader.ContentDescription;
            }
            else if (header == "Content-Disposition")
            {
                return StandardHeader.ContentDisposition;
            }
            else if (header == "Content-Location")
            {
                return StandardHeader.ContentLocation;
            }
            else
            {
                return StandardHeader.ContentLength;
            }
        }
    }
}

