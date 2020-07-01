using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Contains a value used to associate an icon with a particular row of a table.
    /// </summary>
    public enum DisplayType
    {
        /// <summary>
        /// An automated agent, such as Quote-Of-The-Day or a weather chart display.
        /// </summary>
        Agent,
        
        /// <summary>
        /// A distribution list.
        /// </summary>
        DistributionList,
        
        /// <summary>
        /// Display default folder icon adjacent to folder.
        /// </summary>
        Folder,
        
        /// <summary>
        /// Display default folder link icon adjacent to folder rather than the default folder icon.
        /// </summary>
        FolderLink,
        
        /// <summary>
        /// Display icon for a folder with an application-specific distinction, such as a special type of public folder.
        /// </summary>
        FolderSpecial,
        
        /// <summary>
        /// A forum, such as a bulletin board service or a public or shared folder.
        /// </summary>
        Forum,
        
        /// <summary>
        /// A global address book.
        /// </summary>
        GlobalAddressBook,
        
        /// <summary>
        /// A local address book that you share with a small workgroup.
        /// </summary>
        LocalAddressBook,
        
        /// <summary>
        /// A typical messaging user.
        /// </summary>
        MailUser,
        
        /// <summary>
        /// Modifiable; the container should be denoted as modifiable in the user interface.
        /// </summary>
        Modifiable,
        
        /// <summary>
        /// A special alias defined for a large group, such as helpdesk, accounting, or blood-drive coordinator.
        /// </summary>
        Organization,
        
        /// <summary>
        /// A private, personally administered distribution list.
        /// </summary>
        PrivateDistributionList,
        
        /// <summary>
        /// A recipient known to be from a foreign or remote messaging system.
        /// </summary>
        RemoteMailUser,
        
        /// <summary>
        /// A wide area network address book.
        /// </summary>
        WideAreaNetworkAddressBook,
        
        /// <summary>
        /// Does not match any of the other settings.
        /// </summary>
        NotSpecific, 
        
        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
