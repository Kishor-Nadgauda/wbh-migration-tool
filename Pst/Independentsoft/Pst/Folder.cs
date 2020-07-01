using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Folder.
    /// </summary>
    public class Folder
    {
        private PstFileReader reader;
        private ItemDescriptor descriptor;
        private Table table;
        private string displayName;
        private string comment;
        private int itemCount;
        private int unreadItemCount; 
        private bool hasSubFolders;
        private string containerClass;
        private int childrenCount;
        private byte[] entryId;

        internal Folder(PstFileReader reader, ItemDescriptor descriptor, Table table)
        {
            Parse(reader, descriptor, table);
        }

        private void Parse(PstFileReader reader, ItemDescriptor descriptor, Table table)
        {
            this.reader = reader;
            this.descriptor = descriptor;
            this.table = table;

            if (reader.PstFile == null)
            {
              throw new DataStructureException("Parse - reader.PstFile was null");
            }

            if (reader.PstFile.MessageStore == null)
            {
              throw new DataStructureException("Parse - reader.PstFile.MessageStore was null");
            }

            //Compute EntryID
            if (reader.PstFile.MessageStore.RecordKey != null)
            {
                entryId = new byte[24];
                System.Array.Copy(reader.PstFile.MessageStore.RecordKey, 0, entryId, 4, 16);

                byte[] idBytes = BitConverter.GetBytes((uint)descriptor.Id);
                System.Array.Copy(idBytes, 0, entryId, 20, 4);
            }

            if (descriptor != null)
            {
                childrenCount = descriptor.Children.Count;
            }

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_NAME] != null)
            {
                displayName = table.Entries[MapiPropertyTag.PR_DISPLAY_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_COMMENT] != null)
            {
                comment = table.Entries[MapiPropertyTag.PR_COMMENT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CONTENT_COUNT] != null)
            {
                itemCount = table.Entries[MapiPropertyTag.PR_CONTENT_COUNT].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CONTENT_UNREAD] != null)
            {
                unreadItemCount = table.Entries[MapiPropertyTag.PR_CONTENT_UNREAD].GetIntegerValue();
            }

            if (table.Entries[MapiPropertyTag.PR_SUBFOLDERS] != null)
            {
                hasSubFolders = table.Entries[MapiPropertyTag.PR_SUBFOLDERS].GetBooleanValue();
            }

            if (table.Entries[MapiPropertyTag.PR_CONTAINER_CLASS] != null)
            {
                containerClass = table.Entries[MapiPropertyTag.PR_CONTAINER_CLASS].GetStringValue();
            }
        }

        /// <summary>
        /// Gets the folders.
        /// </summary>
        /// <param name="recursive">if set to <c>true</c> [recursive].</param>
        /// <returns>IList{Folder}.</returns>
        public IList<Folder> GetFolders(bool recursive)
        {
            if (recursive)
            {
                IList<Folder> allFolders = new List<Folder>();
                return GetFolders(allFolders, this);
            }
            else
            {
                return GetFolders();
            }
        }

        private static IList<Folder> GetFolders(IList<Folder> allFolders, Folder folder)
        {
            foreach (Folder subFolder in folder.GetFolders())
            {
                allFolders.Add(subFolder);

                GetFolders(allFolders, subFolder);
            }

            return allFolders;
        }

        /// <summary>
        /// Gets the folder.
        /// </summary>
        /// <param name="displayName">The display name.</param>
        /// <returns>Folder.</returns>
        public Folder GetFolder(string displayName)
        {
            IList<Folder> folders = GetFolders();

            for (int i = 0; i < folders.Count; i++)
            {
                if (folders[i].DisplayName == displayName)
                {
                    return folders[i];
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the folders.
        /// </summary>
        /// <returns>IList{Folder}.</returns>
        public IList<Folder> GetFolders()
        {
            IList<Folder> folders = new List<Folder>();

            for (int i = 0; i < descriptor.Children.Count; i++)
            {
                ItemDescriptor childDescriptor = descriptor.Children[i];
                    
                DataStructure childLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, childDescriptor.LocalDescriptorListId);
                DataStructure childDataNode = reader.PstFile.DataIndexTree.GetDataStructure(childDescriptor.TableId);

                if (childDataNode != null)
                {
                    LocalDescriptorList localDescriptorList = Util.GetLocalDescriptorList(reader, childLocalDescriptorListNode);

                    Table childTable = Util.GetTable(reader, childDataNode, localDescriptorList);

                    if (childTable != null && Util.IsFolder(childTable))
                    {
                        Folder child = new Folder(reader, childDescriptor, childTable);
                        folders.Add(child);
                    }
                }
            }

            return folders;
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <returns>IList{Item}.</returns>
        public IList<Item> GetItems()
        {
            return GetItems(0, Int32.MaxValue);
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <param name="endIndex">The end index.</param>
        /// <returns>IList{Item}.</returns>
        public IList<Item> GetItems(int endIndex)
        {
            return GetItems(0, endIndex);
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <param name="endIndex">The end index.</param>
        /// <param name="includeAttachments">if set to <c>true</c> [include attachments].</param>
        /// <returns>IList{Item}.</returns>
        public IList<Item> GetItems(int endIndex, bool includeAttachments)
        {
            return GetItems(0, endIndex, includeAttachments);
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <param name="startIndex">The start index.</param>
        /// <param name="endIndex">The end index.</param>
        /// <returns>IList{Item}.</returns>
        public IList<Item> GetItems(int startIndex, int endIndex)
        {
            return GetItems(startIndex, endIndex, true);
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <param name="startIndex">The start index.</param>
        /// <param name="endIndex">The end index.</param>
        /// <param name="includeAttachments">if set to <c>true</c> [include attachments].</param>
        /// <returns>IList{Item}.</returns>
        public IList<Item> GetItems(int startIndex, int endIndex, bool includeAttachments)
        {
            IList<Item> items = new List<Item>();

            if (startIndex < 0)
            {
                startIndex = 0;
            }

            if (endIndex > descriptor.Children.Count)
            {
                endIndex = descriptor.Children.Count;
            }

            for (int i = startIndex; i < endIndex; i++)
            {
                ItemDescriptor childDescriptor = descriptor.Children[i];

                DataStructure childLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, childDescriptor.LocalDescriptorListId);
                DataStructure childDataNode = reader.PstFile.DataIndexTree.GetDataStructure(childDescriptor.TableId);

                if (childDataNode != null)
                {
                    LocalDescriptorList localDescriptorList = Util.GetLocalDescriptorList(reader, childLocalDescriptorListNode);

                    Table childTable = Util.GetTable(reader, childDataNode, localDescriptorList);

                    if (childTable != null && !Util.IsFolder(childTable))
                    {
                        Item child = Util.GetItem(reader, localDescriptorList, childTable, childDescriptor.Id, childDescriptor.ParentItemDescriptorId, includeAttachments);
                        items.Add(child);
                    }
                }
            }

            return items;
        }

        #region Properties

        internal ItemDescriptor Descriptor
        {
            get
            {
                return descriptor;
            }
        }

        /// <summary>
        /// Gets the identifier.
        /// </summary>
        /// <value>The identifier.</value>
        public long Id
        {
            get
            {
                return (long)descriptor.Id;
            }
        }

        /// <summary>
        /// Gets the parent identifier.
        /// </summary>
        /// <value>The parent identifier.</value>
        public long ParentId
        {
            get
            {
                return (long)descriptor.ParentItemDescriptorId;
            }
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <value>The table.</value>
        public Table Table
        {
            get
            {
                return table;
            }
        }

        /// <summary>
        /// Gets the display name.
        /// </summary>
        /// <value>The display name.</value>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
        }

        /// <summary>
        /// Gets the comment.
        /// </summary>
        /// <value>The comment.</value>
        public string Comment
        {
            get
            {
                return comment;
            }
        }

        /// <summary>
        /// Gets the item count.
        /// </summary>
        /// <value>The item count.</value>
        public int ItemCount
        {
            get
            {
                return itemCount;
            }
        }

        /// <summary>
        /// Gets the unread item count.
        /// </summary>
        /// <value>The unread item count.</value>
        public int UnreadItemCount
        {
            get
            {
                return unreadItemCount;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance has sub folders.
        /// </summary>
        /// <value><c>true</c> if this instance has sub folders; otherwise, <c>false</c>.</value>
        public bool HasSubFolders
        {
            get
            {
                return hasSubFolders;
            }
        }

        /// <summary>
        /// Gets the container class.
        /// </summary>
        /// <value>The container class.</value>
        public string ContainerClass
        {
            get
            {
                return containerClass;
            }
        }

        /// <summary>
        /// Gets the children count.
        /// </summary>
        /// <value>The children count.</value>
        public int ChildrenCount
        {
            get
            {
                return childrenCount;
            }
        }

        /// <summary>
        /// Gets the entry identifier.
        /// </summary>
        /// <value>The entry identifier.</value>
        public byte[] EntryId
        {
            get
            {
                return entryId;
            }
        }

        #endregion
    }
}
