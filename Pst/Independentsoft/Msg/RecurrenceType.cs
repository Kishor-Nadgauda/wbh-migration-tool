using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Specifies the recurrence pattern type.
    /// </summary>
    public enum RecurrenceType
    {
        /// <summary>
        /// Represents a daily recurrence pattern. 
        /// </summary>
        Daily,

        /// <summary>
        /// Represents a weekly recurrence pattern.
        /// </summary>
        Weekly,

        /// <summary>
        /// Represents a monthly recurrence pattern. 
        /// </summary>
        Monthly,

        /// <summary>
        /// Represents a MonthNth recurrence pattern.
        /// </summary>
        MonthNth,

        /// <summary>
        /// Represents a yearly recurrence pattern. 
        /// </summary>
        Yearly,

        /// <summary>
        /// Represents a YearNth recurrence pattern.
        /// </summary>
        YearNth,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
