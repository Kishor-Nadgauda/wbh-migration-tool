using System;
using System.Collections.Generic;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Class RecurrencePattern.
    /// </summary>
    public class RecurrencePattern
    {
        private RecurrencePatternFrequency frequency = RecurrencePatternFrequency.Daily;
        private RecurrencePatternType type = RecurrencePatternType.Day;
        private CalendarType calendarType = CalendarType.None;
        private int period;
        private IList<DayOfWeek> dayOfWeek = new List<DayOfWeek>();
        private DayOfWeekIndex dayOfWeekIndex = DayOfWeekIndex.None;
        private int dayOfMonth;
        private RecurrenceEndType endType = RecurrenceEndType.NeverEnd;
        private int occurenceCount;
        private DayOfWeek firstDayOfWeek = Independentsoft.Msg.DayOfWeek.Sunday;
        private int deletedInstanceCount;
        private IList<DateTime> deletedInstanceDates = new List<DateTime>();
        private int modifiedInstanceCount;
        private IList<DateTime> modifiedInstanceDates = new List<DateTime>();
        private DateTime startDate;
        private DateTime endDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecurrencePattern"/> class.
        /// </summary>
        public RecurrencePattern()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RecurrencePattern"/> class.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public RecurrencePattern(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            if(buffer.Length < 22)
            {
                return;
            }

            short frequencyValue = BitConverter.ToInt16(buffer, 4);
            this.frequency = EnumUtil.ParseRecurrencePatternFrequency(frequencyValue);

            short typeValue = BitConverter.ToInt16(buffer, 6);
            this.type = EnumUtil.ParseRecurrencePatternType(typeValue);

            short calendarTypeValue = BitConverter.ToInt16(buffer, 8);
            this.calendarType = EnumUtil.ParseCalendarType(calendarTypeValue);

            short firstDateTimeValue = BitConverter.ToInt16(buffer, 10);
            
            period = BitConverter.ToInt16(buffer, 14);

            int nextPosition = 22;

            if (this.type == RecurrencePatternType.Day)
            {
                //skip
            }
            else if (this.type == RecurrencePatternType.Week)
            {
                int dayOfWeekValue = BitConverter.ToInt32(buffer, nextPosition);
                nextPosition += 4;

                IList<Independentsoft.Msg.DayOfWeek> list = new List<Independentsoft.Msg.DayOfWeek>();

                if ((dayOfWeekValue & 0x00000001) == 0x00000001)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Sunday);
                }

                if ((dayOfWeekValue & 0x00000002) == 0x00000002)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Monday);
                }

                if ((dayOfWeekValue & 0x00000004) == 0x00000004)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Tuesday);
                }

                if ((dayOfWeekValue & 0x00000008) == 0x00000008)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Wednesday);
                }

                if ((dayOfWeekValue & 0x00000010) == 0x00000010)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Thursday);
                }

                if ((dayOfWeekValue & 0x00000020) == 0x00000020)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Friday);
                }

                if ((dayOfWeekValue & 0x00000040) == 0x00000040)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Saturday);
                }

                dayOfWeek = new DayOfWeek[list.Count];

                for (int i = 0; i < list.Count; i++)
                {
                    dayOfWeek[i] = (Independentsoft.Msg.DayOfWeek)list[i];
                }
            }
            else if (this.type == RecurrencePatternType.Month || this.type == RecurrencePatternType.HijriMonth)
            {
                dayOfMonth = BitConverter.ToInt32(buffer, nextPosition);
                nextPosition += 4;
            }
            else if (this.type == RecurrencePatternType.MonthEnd || this.type == RecurrencePatternType.HijriMonthEnd || this.type == RecurrencePatternType.MonthNth || this.type == RecurrencePatternType.HijriMonthNth)
            {
                if (buffer.Length < 50)
                {
                    return;
                }

                int dayOfWeekValue = BitConverter.ToInt32(buffer, nextPosition);
                nextPosition += 4;

                int dayOfWeekIndexValue = BitConverter.ToInt32(buffer, nextPosition);
                nextPosition += 4;

                IList<Independentsoft.Msg.DayOfWeek> list = new List<Independentsoft.Msg.DayOfWeek>();

                if ((dayOfWeekValue & 0x00000001) == 0x00000001)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Sunday);
                }

                if ((dayOfWeekValue & 0x00000002) == 0x00000002)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Monday);
                }

                if ((dayOfWeekValue & 0x00000004) == 0x00000004)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Tuesday);
                }

                if ((dayOfWeekValue & 0x00000008) == 0x00000008)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Wednesday);
                }

                if ((dayOfWeekValue & 0x00000010) == 0x00000010)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Thursday);
                }

                if ((dayOfWeekValue & 0x00000020) == 0x00000020)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Friday);
                }

                if ((dayOfWeekValue & 0x00000040) == 0x00000040)
                {
                    list.Add(Independentsoft.Msg.DayOfWeek.Saturday);
                }

                dayOfWeek = new DayOfWeek[list.Count];

                for (int i = 0; i < list.Count; i++)
                {
                    dayOfWeek[i] = (Independentsoft.Msg.DayOfWeek)list[i];
                }

                if (dayOfWeekIndexValue == 0x00000001)
                {
                    dayOfWeekIndex = DayOfWeekIndex.First;
                }
                else if (dayOfWeekIndexValue == 0x00000002)
                {
                    dayOfWeekIndex = DayOfWeekIndex.Second;
                }
                else if (dayOfWeekIndexValue == 0x00000003)
                {
                    dayOfWeekIndex = DayOfWeekIndex.Third;
                }
                else if (dayOfWeekIndexValue == 0x00000004)
                {
                    dayOfWeekIndex = DayOfWeekIndex.Fourth;
                }
                else if (dayOfWeekIndexValue == 0x00000005)
                {
                    dayOfWeekIndex = DayOfWeekIndex.Last;
                }
            }

            int endTypeValue = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            this.endType = EnumUtil.ParseRecurrenceEndType(endTypeValue);
            
            occurenceCount = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            int firstDayOfWeekValue = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            this.firstDayOfWeek = EnumUtil.ParseDayOfWeek(firstDayOfWeekValue);

            deletedInstanceCount = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            if (deletedInstanceCount > 0)
            {
                deletedInstanceDates = new DateTime[deletedInstanceCount];

                for (int i = 0; i < deletedInstanceCount; i++)
                {
                    if (buffer.Length < nextPosition + 4)
                    {
                        return;
                    }

                    int minutes = BitConverter.ToInt32(buffer, nextPosition);
                    nextPosition += 4;

                    deletedInstanceDates[i] = Util.GetDateTime(minutes);
                }
            }

            modifiedInstanceCount = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            if (modifiedInstanceCount > 0)
            {
                modifiedInstanceDates = new DateTime[modifiedInstanceCount];

                for (int i = 0; i < modifiedInstanceCount; i++)
                {
                    if (buffer.Length < nextPosition + 4)
                    {
                        return;
                    }

                    int minutes = BitConverter.ToInt32(buffer, nextPosition);
                    nextPosition += 4;

                    modifiedInstanceDates[i] = Util.GetDateTime(minutes);
                }
            }

            if (buffer.Length < nextPosition + 4)
            {
                return;
            }

            int startDateMinutes = BitConverter.ToInt32(buffer, nextPosition);
            nextPosition += 4;

            this.startDate = Util.GetDateTime(startDateMinutes);

            int endDateMinutes = BitConverter.ToInt32(buffer, nextPosition);

            this.endDate = Util.GetDateTime(endDateMinutes);
        }

        #region Properties

        /// <summary>
        /// Gets the frequency.
        /// </summary>
        /// <value>The frequency.</value>
        public RecurrencePatternFrequency Frequency
        {
            get
            {
                return frequency;
            }
        }

        /// <summary>
        /// Gets the type.
        /// </summary>
        /// <value>The type.</value>
        public RecurrencePatternType Type
        {
            get
            {
                return type;
            }
        }

        /// <summary>
        /// Gets the type of the calendar.
        /// </summary>
        /// <value>The type of the calendar.</value>
        public CalendarType CalendarType
        {
            get
            {
                return calendarType;
            }
        }

        /// <summary>
        /// Gets the period.
        /// </summary>
        /// <value>The period.</value>
        public int Period
        {
            get
            {
                return period;
            }
        }

        /// <summary>
        /// Gets the day of week.
        /// </summary>
        /// <value>The day of week.</value>
        public IList<Independentsoft.Msg.DayOfWeek> DayOfWeek
        {
            get
            {
                return dayOfWeek;
            }
        }

        /// <summary>
        /// Gets the index of the day of week.
        /// </summary>
        /// <value>The index of the day of week.</value>
        public DayOfWeekIndex DayOfWeekIndex
        {
            get
            {
                return dayOfWeekIndex;
            }
        }

        /// <summary>
        /// Gets the day of month.
        /// </summary>
        /// <value>The day of month.</value>
        public int DayOfMonth
        {
            get
            {
                return dayOfMonth;
            }
        }

        /// <summary>
        /// Gets the end type.
        /// </summary>
        /// <value>The end type.</value>
        public RecurrenceEndType EndType
        {
            get
            {
                return endType;
            }
        }

        /// <summary>
        /// Gets the occurence count.
        /// </summary>
        /// <value>The occurence count.</value>
        public int OccurenceCount
        {
            get
            {
                return occurenceCount;
            }
        }

        /// <summary>
        /// Gets the first day of week.
        /// </summary>
        /// <value>The first day of week.</value>
        public Independentsoft.Msg.DayOfWeek FirstDayOfWeek
        {
            get
            {
                return firstDayOfWeek;
            }
        }

        /// <summary>
        /// Gets the deleted instance count.
        /// </summary>
        /// <value>The deleted instance count.</value>
        public int DeletedInstanceCount
        {
            get
            {
                return deletedInstanceCount;
            }
        }

        /// <summary>
        /// Gets the deleted instance dates.
        /// </summary>
        /// <value>The deleted instance dates.</value>
        public IList<DateTime> DeletedInstanceDates 
        {
            get
            {
                return deletedInstanceDates;
            }
        }

        /// <summary>
        /// Gets the modified instance count.
        /// </summary>
        /// <value>The modified instance count.</value>
        public int ModifiedInstanceCount
        {
            get
            {
                return modifiedInstanceCount;
            }
        }

        /// <summary>
        /// Gets the modified instance dates.
        /// </summary>
        /// <value>The modified instance dates.</value>
        public IList<DateTime> ModifiedInstanceDates
        {
            get
            {
                return modifiedInstanceDates;
            }
        }

        /// <summary>
        /// Gets the start date.
        /// </summary>
        /// <value>The start date.</value>
        public DateTime StartDate
        {
            get
            {
                return startDate;
            }
        }

        /// <summary>
        /// Gets the end date.
        /// </summary>
        /// <value>The end date.</value>
        public DateTime EndDate
        {
            get
            {
                return endDate;
            }
        }

        #endregion
    }
}
