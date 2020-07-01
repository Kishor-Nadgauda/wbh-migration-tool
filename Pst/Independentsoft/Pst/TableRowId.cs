using System;

namespace Independentsoft.Pst
{
    internal class TableRowId
    {
        private uint dwRowId;
        private uint dwRowIndex;

        internal TableRowId()
        {
        }

        internal TableRowId(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            if (buffer.Length == 8)
            {
                dwRowId = BitConverter.ToUInt32(buffer, 0);
                dwRowIndex = BitConverter.ToUInt32(buffer, 4);
            }
            else
            {
                dwRowId = BitConverter.ToUInt32(buffer, 0);
                dwRowIndex = BitConverter.ToUInt16(buffer, 4);
            }
        }

        #region Properties

        internal uint Id
        {
            get
            {
                return dwRowId;
            }
        }


        internal uint Index
        {
            get
            {
                return dwRowIndex;
            }
        }
        
        #endregion
    }
}
