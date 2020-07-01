using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class TableEntry.
    /// </summary>
    public class TableEntry
    {
        private PropertyTag propertyTag;
        private byte[] valueBuffer;
        private System.Text.Encoding pstFileEncoding = System.Text.Encoding.UTF8;
        
        internal TableEntry(System.Text.Encoding pstFileEncoding)
        {
            this.pstFileEncoding = pstFileEncoding;
        }

        /// <summary>
        /// Gets the boolean value.
        /// </summary>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public bool GetBooleanValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 2)
            {
                bool boolValue = BitConverter.ToBoolean(valueBuffer, 0);
                return boolValue;
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
            if (valueBuffer != null && valueBuffer.Length == 2)
            {
                short shortValue = BitConverter.ToInt16(valueBuffer, 0);
                return shortValue;
            }
            else
            {
                return Int16.MinValue;
            }
        }

        /// <summary>
        /// Gets the short array value.
        /// </summary>
        /// <returns>System.Int16[][].</returns>
        public short[] GetShortArrayValue()
        {
            if (valueBuffer != null)
            {
                short[] shortArray = new short[valueBuffer.Length / 2];

                for (int i = 0; i < shortArray.Length; i++)
                {
                    shortArray[i] = BitConverter.ToInt16(valueBuffer, i * 2);
                }

                return shortArray;
            }
            else
            {
                return new short[0];
            }
        }

        /// <summary>
        /// Gets the integer value.
        /// </summary>
        /// <returns>System.Int32.</returns>
        public int GetIntegerValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 4)
            {
                int intValue = BitConverter.ToInt32(valueBuffer, 0);
                return intValue;
            }
            else
            {
                return Int32.MinValue;
            }
        }

        /// <summary>
        /// Gets the integer array value.
        /// </summary>
        /// <returns>System.Int32[][].</returns>
        public int[] GetIntegerArrayValue()
        {
            if (valueBuffer != null)
            {
                int[] intArray = new int[valueBuffer.Length / 4];

                for (int i = 0; i < intArray.Length; i++)
                {
                    intArray[i] = BitConverter.ToInt32(valueBuffer, i * 4);
                }

                return intArray;
            }
            else
            {
                return new int[0];
            }
        }

        /// <summary>
        /// Gets the float value.
        /// </summary>
        /// <returns>System.Single.</returns>
        public float GetFloatValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 4)
            {
                float floatValue = BitConverter.ToSingle(valueBuffer, 0);
                return floatValue;
            }
            else
            {
                return Single.MinValue;
            }
        }

        /// <summary>
        /// Gets the float array value.
        /// </summary>
        /// <returns>System.Single[][].</returns>
        public float[] GetFloatArrayValue()
        {
            if (valueBuffer != null)
            {
                float[] floatArray = new float[valueBuffer.Length / 4];

                for (int i = 0; i < floatArray.Length; i++)
                {
                    floatArray[i] = BitConverter.ToSingle(valueBuffer, i * 4);
                }

                return floatArray;
            }
            else
            {
                return new float[0];
            }
        }

        /// <summary>
        /// Gets the long value.
        /// </summary>
        /// <returns>System.Int64.</returns>
        public long GetLongValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 8)
            {
                long longValue = BitConverter.ToInt64(valueBuffer, 0);
                return longValue;
            }
            else
            {
                return Int64.MinValue;
            }
        }

        /// <summary>
        /// Gets the long array value.
        /// </summary>
        /// <returns>System.Int64[][].</returns>
        public long[] GetLongArrayValue()
        {
            if (valueBuffer != null)
            {
                long[] longArray = new long[valueBuffer.Length / 8];

                for (int i = 0; i < longArray.Length; i++)
                {
                    longArray[i] = BitConverter.ToInt64(valueBuffer, i * 8);
                }

                return longArray;
            }
            else
            {
                return new long[0];
            }
        }

        /// <summary>
        /// Gets the double value.
        /// </summary>
        /// <returns>System.Double.</returns>
        public double GetDoubleValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 8)
            {
                double doubleValue = BitConverter.ToDouble(valueBuffer, 0);
                return doubleValue;
            }
            else
            {
                return Double.MinValue;
            }
        }

        /// <summary>
        /// Gets the double array value.
        /// </summary>
        /// <returns>System.Double[][].</returns>
        public double[] GetDoubleArrayValue()
        {
            if (valueBuffer != null)
            {
                double[] doubleArray = new double[valueBuffer.Length / 8];

                for (int i = 0; i < doubleArray.Length; i++)
                {
                    doubleArray[i] = BitConverter.ToDouble(valueBuffer, i * 8);
                }

                return doubleArray;
            }
            else
            {
                return new double[0];
            }
        }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <returns>System.String.</returns>
        public string GetStringValue()
        {
            if (valueBuffer != null)
            {
                if (propertyTag.Type == PropertyType.String)
                {
                    string stringValue = Util.GetCorrectedStringValue(propertyTag, valueBuffer, System.Text.Encoding.Unicode);
                    return stringValue;
                }
                else if (propertyTag.Type == PropertyType.String8)
                {
                    string stringValue = Util.GetCorrectedStringValue(propertyTag, valueBuffer, pstFileEncoding);
                    return stringValue;
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
            if (valueBuffer != null)
            {
                if (propertyTag.Type == PropertyType.StringArray)
                {
                    string[] stringArrayValue = Util.GetStringArrayValue(valueBuffer, System.Text.Encoding.Unicode);
                    return stringArrayValue;
                }
                else if (propertyTag.Type == PropertyType.String8Array)
                {
                    string[] stringArrayValue = Util.GetStringArrayValue(valueBuffer, pstFileEncoding);
                    return stringArrayValue;
                }
                else
                {
                    return new string[0];
                }
            }
            else
            {
                return new string[0];
            }
        }

        /// <summary>
        /// Gets the binary value.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetBinaryValue()
        {
            return valueBuffer;
        }

        /// <summary>
        /// Gets the binary array value.
        /// </summary>
        /// <returns>IList{System.Byte[]}.</returns>
        public IList<byte[]> GetBinaryArrayValue()
        {
            if (valueBuffer != null)
            {
                IList<byte[]> binaryArray = Util.GetBinaryArrayValues(valueBuffer);

                return binaryArray;
            }
            else
            {
                return new List<byte[]>();
            }
        }

        /// <summary>
        /// Gets the unique identifier value.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetGuidValue()
        {
            return valueBuffer;
        }

        /// <summary>
        /// Gets the unique identifier array value.
        /// </summary>
        /// <returns>IList{System.Byte[]}.</returns>
        public IList<byte[]> GetGuidArrayValue()
        {
            if (valueBuffer != null)
            {
                IList<byte[]> guidArray = new List<byte[]>();

                for (int i = 0; i < valueBuffer.Length / 16; i++)
                {
                    byte[] guid = new byte[16];
                    System.Array.Copy(valueBuffer, i * 16, guid, 0, guid.Length);

                    guidArray.Add(guid);
                }

                return guidArray;
            }
            else
            {
                return new List<byte[]>();
            }
        }

        /// <summary>
        /// Gets the date time value.
        /// </summary>
        /// <returns>DateTime.</returns>
        public DateTime GetDateTimeValue()
        {
            if (valueBuffer != null && valueBuffer.Length == 8)
            {
                uint dateTimeLow = BitConverter.ToUInt32(valueBuffer, 0);
                ulong dateTimeHigh = BitConverter.ToUInt32(valueBuffer, 4);

                if (dateTimeHigh > 0)
                {
                    long ticks = dateTimeLow + (long)(dateTimeHigh << 32);

                    DateTime year1601 = new DateTime(1601, 1, 1);

                    try
                    {
                      if (year1601.Ticks + ticks < 0)
                      {
                        return DateTime.MinValue;
                      }
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

        #region Properties

        /// <summary>
        /// Gets or sets the property tag.
        /// </summary>
        /// <value>The property tag.</value>
        public PropertyTag PropertyTag
        {
            get
            {
                return propertyTag;
            }
            set
            {
                this.propertyTag = value; 
            }
        }

        /// <summary>
        /// Gets or sets the value buffer.
        /// </summary>
        /// <value>The value buffer.</value>
        public byte[] ValueBuffer
        {
            get
            {
                return valueBuffer;
            }
            set
            {
                this.valueBuffer = value;
            }
        }

        #endregion
    }
}
