using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Journal.
    /// </summary>
    public class Journal : Item
    {
        internal Journal()
        {
        }

        #region Properties

        /// <summary>
        /// Gets the start time.
        /// </summary>
        /// <value>The start time.</value>
        public DateTime StartTime
        {
            get
            {
                return journalStartTime;
            }
        }

        /// <summary>
        /// Gets the end time.
        /// </summary>
        /// <value>The end time.</value>
        public DateTime EndTime
        {
            get
            {
                return journalEndTime;
            }
        }

        /// <summary>
        /// Gets the type.
        /// </summary>
        /// <value>The type.</value>
        public string Type
        {
            get
            {
                return journalType;
            }
        }

        /// <summary>
        /// Gets the type description.
        /// </summary>
        /// <value>The type description.</value>
        public string TypeDescription
        {
            get
            {
                return journalTypeDescription;
            }
        }

        /// <summary>
        /// Gets the duration.
        /// </summary>
        /// <value>The duration.</value>
        public long Duration
        {
            get
            {
                return journalDuration;
            }
        }

        #endregion
    }
}
