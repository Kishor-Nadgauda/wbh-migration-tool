using System;
using System.Collections.Generic;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Class ExtendedPropertyList.
    /// </summary>
    public class ExtendedPropertyList : List<ExtendedProperty>
    {
        /// <summary>
        /// Gets the <see cref="ExtendedProperty"/> with the specified tag.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <returns>ExtendedProperty.</returns>
        public ExtendedProperty this[ExtendedPropertyTag tag]
        {
            get
            {
                if (tag != null)
                {
                    if (tag is ExtendedPropertyId)
                    {
                        ExtendedPropertyId extendedPropertyId = (ExtendedPropertyId)tag;

                        for (int i = 0; i < this.Count; i++)
                        {
                            ExtendedProperty current = this[i];

                            if (current.Tag is ExtendedPropertyId)
                            {
                                ExtendedPropertyId currentPropertyId = (ExtendedPropertyId)current.Tag;

                                if (currentPropertyId.Id == extendedPropertyId.Id)
                                {
                                    bool isGuidEqual = true;

                                    if (currentPropertyId.Guid != null && extendedPropertyId.Guid != null && currentPropertyId.Guid.Length == extendedPropertyId.Guid.Length)
                                    {
                                        for (int j = 0; j < currentPropertyId.Guid.Length; j++)
                                        {
                                            if (currentPropertyId.Guid[j] != extendedPropertyId.Guid[j])
                                            {
                                                isGuidEqual = false;
                                                break;
                                            }
                                        }
                                    }

                                    if (isGuidEqual)
                                    {
                                        return current;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        ExtendedPropertyName extendedPropertyName = (ExtendedPropertyName)tag;

                        for (int i = 0; i < this.Count; i++)
                        {
                            ExtendedProperty current = this[i];

                            if (current.Tag is ExtendedPropertyName)
                            {
                                ExtendedPropertyName currentPropertyName = (ExtendedPropertyName)current.Tag;

                                if (currentPropertyName.Name == extendedPropertyName.Name)
                                {
                                    bool isGuidEqual = true;

                                    if (currentPropertyName.Guid != null && extendedPropertyName.Guid != null && currentPropertyName.Guid.Length == extendedPropertyName.Guid.Length)
                                    {
                                        for (int j = 0; j < currentPropertyName.Guid.Length; j++)
                                        {
                                            if (currentPropertyName.Guid[j] != extendedPropertyName.Guid[j])
                                            {
                                                isGuidEqual = false;
                                                break;
                                            }
                                        }
                                    }

                                    if (isGuidEqual)
                                    {
                                        return current;
                                    }
                                }
                            }
                        }
                    }

                }

                return null;
            }
        }
    }
}
