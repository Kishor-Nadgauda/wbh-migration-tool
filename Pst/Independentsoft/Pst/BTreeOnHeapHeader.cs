using System;

namespace Independentsoft.Pst
{
    internal class BTreeOnHeapHeader
    {
        private byte bTypeBTH;
        private byte cbKey;
        private byte cbEnt;
        private byte bIdxLevels;
        private HeapId hidRoot;

        internal BTreeOnHeapHeader()
        {
        }

        internal BTreeOnHeapHeader(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            bTypeBTH = buffer[0];
            cbKey = buffer[1];
            cbEnt = buffer[2];
            bIdxLevels = buffer[3];

            byte[] hidRootBuffer = new byte[4];
            System.Array.Copy(buffer, 4, hidRootBuffer, 0, 4);

            hidRoot = new HeapId(hidRootBuffer);
        }

        #region Properties

        internal byte CbKey
        {
            get
            {
                return cbKey;
            }
        }

        internal byte CbEnt
        {
            get
            {
                return cbEnt;
            }
        }

        internal byte BIdxLevels
        {
            get
            {
                return bIdxLevels;
            }
        }

        internal HeapId HidRoot
        {
            get
            {
                return hidRoot;
            }
        }

        #endregion
    }
}
