using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents type of an object.
    /// </summary>
    public enum ObjectType
    {
        /// <summary>
        /// Address book container object.
        /// </summary>
        AddressBookContainer,

        /// <summary>
        /// Address book object.
        /// </summary>
        AddressBook,

        /// <summary>
        /// Message attachment object.
        /// </summary>
        Attachment,

        /// <summary>
        /// Distribution list object.
        /// </summary>
        DistributionList,

        /// <summary>
        /// Folder object.
        /// </summary>
        Folder,

        /// <summary>
        /// Form object.
        /// </summary>
        Form,

        /// <summary>
        /// Messaging user object.
        /// </summary>
        MailUser,

        /// <summary>
        /// Message object.
        /// </summary>
        Message,

        /// <summary>
        /// Profile section object.
        /// </summary>
        ProfileSelection,

        /// <summary>
        /// Session object.
        /// </summary>
        Session,

        /// <summary>
        /// Status object.
        /// </summary>
        Status,

        /// <summary>
        /// Message store object.
        /// </summary>
        MessageStore,

        /// <summary>
        /// None
        /// </summary>
        None,
    }
}
