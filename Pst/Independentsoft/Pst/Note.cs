using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Note.
    /// </summary>
    public class Note : Item
    {
        internal Note()
        {
        }

        #region Note

        /// <summary>
        /// Gets the height.
        /// </summary>
        /// <value>The height.</value>
        public long Height
        {
            get
            {
                return noteHeight;
            }
        }

        /// <summary>
        /// Gets the width.
        /// </summary>
        /// <value>The width.</value>
        public long Width
        {
            get
            {
                return noteWidth;
            }
        }

        /// <summary>
        /// Gets the top.
        /// </summary>
        /// <value>The top.</value>
        public long Top
        {
            get
            {
                return noteTop;
            }
        }

        /// <summary>
        /// Gets the left.
        /// </summary>
        /// <value>The left.</value>
        public long Left
        {
            get
            {
                return noteLeft;
            }
        }

        /// <summary>
        /// Gets the color.
        /// </summary>
        /// <value>The color.</value>
        public NoteColor Color
        {
            get
            {
                return noteColor;
            }
        }

        #endregion
    }
}
