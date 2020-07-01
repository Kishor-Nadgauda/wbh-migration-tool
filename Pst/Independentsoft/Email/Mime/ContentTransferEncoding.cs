using System;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Specifies the Content-Transfer-Encoding header information for an e-mail message attachment.
    /// </summary>
    public enum ContentTransferEncoding
    {
        /// <summary>
        /// 
        /// </summary>
        SevenBit,

        /// <summary>
        /// 
        /// </summary>
        EightBit,

        /// <summary>
        /// 
        /// </summary>
        Base64,

        /// <summary>
        /// 
        /// </summary>
        Binary,

        /// <summary>
        /// 
        /// </summary>
        QuotedPrintable,

        /// <summary>
        /// 
        /// </summary>
        IetfToken,

        /// <summary>
        /// 
        /// </summary>
        XToken
    }
}
