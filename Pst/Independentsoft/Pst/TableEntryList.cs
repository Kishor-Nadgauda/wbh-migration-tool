using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class TableEntryList.
    /// </summary>
    public class TableEntryList : List<TableEntry>
    {
        /// <summary>
        /// Gets the <see cref="TableEntry"/> with the specified property tag.
        /// </summary>
        /// <param name="propertyTag">The property tag.</param>
        /// <returns>TableEntry.</returns>
        public TableEntry this[PropertyTag propertyTag]
        {
            get
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (this[i].PropertyTag.Tag == propertyTag.Tag)
                    {
                        return this[i];
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the <see cref="TableEntry"/> with the specified tag.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="type">The type.</param>
        /// <returns>TableEntry.</returns>
        public TableEntry this[int tag, PropertyType type]
        {
            get
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (this[i].PropertyTag.Tag == tag && this[i].PropertyTag.Type == type)
                    {
                        return this[i];
                    }
                }

                return null;
            }
        }
    }
}
