using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class ExtendedPropertyName.
    /// </summary>
    public class ExtendedPropertyName : ExtendedPropertyTag
    {
        private string name;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyName"/> class.
        /// </summary>
        public ExtendedPropertyName()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyName"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="guid">The unique identifier.</param>
        public ExtendedPropertyName(string name, byte[] guid)
        {
            this.name = name;
            this.guid = guid;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedPropertyName"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="guid">The unique identifier.</param>
        /// <param name="type">The type.</param>
        public ExtendedPropertyName(string name, byte[] guid, PropertyType type)
        {
            this.name = name;
            this.guid = guid;
            this.type = type;
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            return name;
        }

        #region Properties

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        #endregion
    }
}