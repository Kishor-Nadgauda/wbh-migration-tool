using System;
using System.IO;
using System.Text;

namespace Independentsoft.Pst
{
    internal class PstFileReader : BinaryReader
    {
        private PstFile pstFile;

        internal PstFileReader(PstFile pstFile, Stream stream) : base(stream)
        {
            this.pstFile = pstFile;
        }

        internal PstFileReader(PstFile pstFile, Stream stream, Encoding encoding) : base(stream, encoding)
        {
            this.pstFile = pstFile;
        }

        #region Properties

        internal PstFile PstFile
        {
            get
            {
                return pstFile;
            }
            set
            {
                pstFile = value;
            }
        }

        #endregion
    }
}
