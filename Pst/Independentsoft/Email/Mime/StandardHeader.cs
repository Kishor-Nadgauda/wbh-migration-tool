using System;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Contains the standard message headers defined in RFC 2822.
    /// </summary>
    public enum StandardHeader
    {
        /// <summary>
        /// Resent-Date header field.
        /// </summary>
        ResentDate,

        /// <summary>
        /// Resent-From header field.
        /// </summary>
        ResentFrom,

        /// <summary>
        /// Resent-Sender header field.
        /// </summary>
        ResentSender,

        /// <summary>
        /// Resent-To header field.
        /// </summary>
        ResentTo,

        /// <summary>
        /// Resent-Cc header field.
        /// </summary>
        ResentCc,

        /// <summary>
        /// Resent-Bcc header field.
        /// </summary>
        ResentBcc,

        /// <summary>
        /// Resent-Msg-ID header field.
        /// </summary>
        ResentMessageID,

        /// <summary>
        /// From header field.
        /// </summary>
        From,

        /// <summary>
        /// Sender header field.
        /// </summary>
        Sender,

        /// <summary>
        /// Reply-To header field.
        /// </summary>
        ReplyTo,

        /// <summary>
        /// To header field.
        /// </summary>
        To,

        /// <summary>
        /// Cc header field.
        /// </summary>
        Cc,

        /// <summary>
        /// Bcc header field.
        /// </summary>
        Bcc,

        /// <summary>
        /// Message-ID header field.
        /// </summary>
        MessageID,

        /// <summary>
        /// In-Reply-To header field.
        /// </summary>
        InReplyTo,

        /// <summary>
        /// References header field.
        /// </summary>
        References,

        /// <summary>
        /// Subject header field.
        /// </summary>
        Subject,

        /// <summary>
        /// Comments header field.
        /// </summary>
        Comments,

        /// <summary>
        /// Keywords header field.
        /// </summary>
        Keywords,

        /// <summary>
        /// Date header field.
        /// </summary>
        Date,

        /// <summary>
        /// Return-Path header field.
        /// </summary>
        ReturnPath,

        /// <summary>
        /// Received header field.
        /// </summary>
        Received,

        /// <summary>
        /// MIME-Version header field.
        /// </summary>
        MimeVersion,

        /// <summary>
        /// Content-Type header field.
        /// </summary>
        ContentType,

        /// <summary>
        /// Content-ID header field.
        /// </summary>
        ContentID,

        /// <summary>
        /// Content-Transfer-Encoding header field.
        /// </summary>
        ContentTransferEncoding,

        /// <summary>
        /// Content-Description header field.
        /// </summary>
        ContentDescription,

        /// <summary>
        /// Content-Disposition header field.
        /// </summary>
        ContentDisposition,

        /// <summary>
        /// Content-Location header field.
        /// </summary>
        ContentLocation,

        /// <summary>
        /// Content-Length header field.
        /// </summary>
        ContentLength
    }
}
