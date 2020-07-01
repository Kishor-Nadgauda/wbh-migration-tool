using System;
using System.Globalization;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Class ExtendedProperty.
    /// </summary>
    public class ExtendedProperty
    {
        private ExtendedPropertyTag tag;
        private byte[] value;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        public ExtendedProperty()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        public ExtendedProperty(ExtendedPropertyTag tag)
        {
            this.tag = tag;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        /// <exception cref="System.ArgumentNullException">
        /// tag
        /// or
        /// value
        /// </exception>
        public ExtendedProperty(ExtendedPropertyTag tag, string value)
        {
            if (tag == null)
            {
                throw new ArgumentNullException("tag");
            }

            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            this.tag = tag;

            if (tag.Type == PropertyType.String)
            {
                this.value = System.Text.Encoding.Unicode.GetBytes(value);
            }
            else
            {
                this.value = System.Text.Encoding.UTF8.GetBytes(value);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, short value)
        {
            this.tag = tag;
            this.value = BitConverter.GetBytes(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, int value)
        {
            this.tag = tag;
            this.value = BitConverter.GetBytes(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, long value)
        {
            this.tag = tag;
            this.value = BitConverter.GetBytes(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, double value)
        {
            this.tag = tag;
            this.value = BitConverter.GetBytes(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">if set to <c>true</c> [value].</param>
        public ExtendedProperty(ExtendedPropertyTag tag, bool value)
        {
            this.tag = tag;
            this.value = BitConverter.GetBytes(value);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, DateTime value)
        {
            this.tag = tag;

            DateTime year1601 = new DateTime(1601, 1, 1);
            TimeSpan timeSpan = value.ToUniversalTime().Subtract(year1601);

            long ticks = timeSpan.Ticks;

            byte[] ticksBytes = BitConverter.GetBytes(ticks);

            this.value = ticksBytes;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        /// <exception cref="System.ArgumentNullException">
        /// tag
        /// or
        /// value
        /// </exception>
        public ExtendedProperty(ExtendedPropertyTag tag, string[] value)
        {
            if (tag == null)
            {
                throw new ArgumentNullException("tag");
            }

            if (value == null)
            {
                throw new ArgumentNullException("value");
            }

            this.tag = tag;

            if (value != null && value.Length > 0)
            {
                System.IO.MemoryStream memoryStream = new System.IO.MemoryStream();

                using (memoryStream)
                {
                    uint count = (uint)value.Length;
                    byte[] countBuffer = BitConverter.GetBytes(count);

                    memoryStream.Write(countBuffer, 0, countBuffer.Length);

                    for (int i = 0; i < count; i++)
                    {
                        if (tag.Type == PropertyType.String)
                        {
                            byte[] buffer = System.Text.Encoding.Unicode.GetBytes(value[i]);

                            byte[] lengthBuffer = BitConverter.GetBytes(buffer.Length);

                            memoryStream.Write(lengthBuffer, 0, lengthBuffer.Length);
                            memoryStream.Write(buffer, 0, buffer.Length);
                        }
                        else
                        {
                            byte[] buffer = System.Text.Encoding.UTF8.GetBytes(value[i]);

                            byte[] lengthBuffer = BitConverter.GetBytes(buffer.Length);

                            memoryStream.Write(lengthBuffer, 0, lengthBuffer.Length);
                            memoryStream.Write(buffer, 0, buffer.Length);
                        }
                    }

                    this.value = memoryStream.ToArray();
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="value">The value.</param>
        public ExtendedProperty(ExtendedPropertyTag tag, byte[] value)
        {
            this.tag = tag;
            this.value = value;
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <returns>System.String.</returns>
        public string GetStringValue()
        {
            if(value != null)
            {
                if (tag.Type == PropertyType.String)
                {
                    return System.Text.Encoding.Unicode.GetString(value, 0, value.Length);
                }
                else if (tag.Type == PropertyType.String8)
                {
                    return System.Text.Encoding.UTF8.GetString(value, 0, value.Length);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the string array value.
        /// </summary>
        /// <returns>System.String[][].</returns>
        public string[] GetStringArrayValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.MultipleString)
                {
                    return Util.GetStringArrayValues(value, System.Text.Encoding.Unicode);
                }
                else if (tag.Type == PropertyType.MultipleString8)
                {
                    return Util.GetStringArrayValues(value, System.Text.Encoding.UTF8);
                }
                else
                {
                    return new string[0];
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the boolean value.
        /// </summary>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public bool GetBooleanValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Boolean && value.Length == 2)
                {
                    return BitConverter.ToBoolean(value,0);
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the short value.
        /// </summary>
        /// <returns>System.Int16.</returns>
        public short GetShortValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Integer16 && value.Length == 2)
                {
                    return BitConverter.ToInt16(value, 0);
                }
                else
                {
                    return Int16.MinValue;
                }
            }
            else
            {
                return Int16.MinValue;
            }
        }

        /// <summary>
        /// Gets the integer value.
        /// </summary>
        /// <returns>System.Int32.</returns>
        public int GetIntegerValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Integer32 && value.Length == 4)
                {
                    return BitConverter.ToInt32(value, 0);
                }
                else
                {
                    return Int32.MinValue;
                }
            }
            else
            {
                return Int32.MinValue;
            }
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <returns>System.Int64.</returns>
        public long GetLongValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Integer64 && value.Length == 8)
                {
                    return BitConverter.ToInt64(value, 0);
                }
                else
                {
                    return Int64.MinValue;
                }
            }
            else
            {
                return Int64.MinValue;
            }
        }

        /// <summary>
        /// Gets the float value.
        /// </summary>
        /// <returns>System.Single.</returns>
        public float GetFloatValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Floating32 && value.Length == 4)
                {
                    return BitConverter.ToSingle(value, 0);
                }
                else
                {
                    return float.MinValue;
                }
            }
            else
            {
                return float.MinValue;
            }
        }

        /// <summary>
        /// Gets the double value.
        /// </summary>
        /// <returns>System.Double.</returns>
        public double GetDoubleValue()
        {
            if (value != null)
            {
                if (tag.Type == PropertyType.Floating64 && value.Length == 8)
                {
                    return BitConverter.ToDouble(value, 0);
                }
                else
                {
                    return Double.MinValue;
                }
            }
            else
            {
                return Double.MinValue;
            }
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <returns>DateTime.</returns>
        public DateTime GetDateTimeValue()
        {
            if (value != null)
            {
                if ((tag.Type == PropertyType.Time || tag.Type == PropertyType.FloatingTime) && value.Length == 8)
                {
                    uint dateTimeLow = BitConverter.ToUInt32(value, 0);
                    ulong dateTimeHigh = BitConverter.ToUInt32(value, 4);

                    if (dateTimeHigh > 0)
                    {
                        long ticks = dateTimeLow + (long)(dateTimeHigh << 32);

                        DateTime year1601 = new DateTime(1601, 1, 1);

                        try
                        {
                            DateTime dateTimeValue = year1601.AddTicks(ticks).ToLocalTime();
                            return dateTimeValue;
                        }
                        catch (Exception) //ignore wrong dates
                        {
                        }
                    }

                    return DateTime.MinValue;
                }
                else
                {
                    return DateTime.MinValue;
                }
            }
            else
            {
                return DateTime.MinValue;
            }
        }

        #region Properties

        /// <summary>
        /// Gets or sets the tag.
        /// </summary>
        /// <value>The tag.</value>
        public ExtendedPropertyTag Tag
        {
            get
            {
                return tag;
            }
            set
            {
                tag = value;
            }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>The value.</value>
        public byte[] Value
        {
            get
            {
                return value;
            }
            set
            {
                this.value = value;
            }
        }

        #endregion

    }
}
