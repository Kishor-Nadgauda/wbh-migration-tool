using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class ExtendedPropertyId.
    /// </summary>
    public class ExtendedPropertyId : ExtendedPropertyTag
    {
        private int id;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyId"/> class.
        /// </summary>
        public ExtendedPropertyId()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyId"/> class.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <param name="guid">The unique identifier.</param>
        public ExtendedPropertyId(int id, byte[] guid)
        {
            this.id = id;
            this.guid = guid;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyId"/> class.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <param name="guid">The unique identifier.</param>
        /// <param name="type">The type.</param>
        public ExtendedPropertyId(int id, byte[] guid, PropertyType type)
        {
            this.id = id;
            this.guid = guid;
            this.type = type;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            return id.ToString();
        }

        #region Properties

        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>The identifier.</value>
        public int Id
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }

        #endregion
    }
}