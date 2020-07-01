using System;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class HeaderList.
    /// </summary>
    public class HeaderList : List<Header>
    {
        /// <summary>
        /// Removes the specified standard header.
        /// </summary>
        /// <param name="standardHeader">The standard header.</param>
        public void Remove(StandardHeader standardHeader)
        {
            string name = EnumUtil.ParseStandardHeader(standardHeader);
            Remove(name);
        }

        /// <summary>
        /// Removes the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        public void Remove(string name)
        {
            if (name != null)
            {
                for (int i = this.Count - 1; i >= 0; i--)
                {
                    if (this[i] != null && this[i].Name != null && this[i].Name.ToLower() == name.ToLower())
                    {
                        this.RemoveAt(i);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the <see cref="Header"/> with the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>Header.</returns>
        public Header this[StandardHeader name]
        {
            get
            {
                string headerName = EnumUtil.ParseStandardHeader(name);
                return this[headerName];
            }
        }

        /// <summary>
        /// Gets the <see cref="Header"/> with the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>Header.</returns>
        public Header this[string name]
        {
            get
            {
                if (name != null)
                {
                    for (int i = 0; i < this.Count; i++)
                    {
                        if (this[i] != null && this[i].Name != null && this[i].Name.ToLower() == name.ToLower())
                        {
                            return this[i];
                        }
                    }
                }

                return null;
            }
        }
    }
}
