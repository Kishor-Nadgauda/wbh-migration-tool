using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Class ExtendedPropertyTag.
    /// </summary>
    public abstract class ExtendedPropertyTag
    {
        internal byte[] guid;
        internal PropertyType type = PropertyType.String;

        #region Properties

        /// <summary>
        /// Gets or sets the unique identifier.
        /// </summary>
        /// <value>The unique identifier.</value>
        public byte[] Guid
        {
            get
            {
                return guid;
            }
            set
            {
                guid = value;
            }
        }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The type.</value>
        public PropertyType Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
            }
        }

        #endregion
    }
}
