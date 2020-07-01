using System;

namespace Independentsoft.Pst
{
    internal class TableIndexItem
    {
        private int startOffset;
        private int endOffset;

        internal TableIndexItem()
        {
        }

        internal TableIndexItem Clone()
        {
            TableIndexItem clone = new TableIndexItem();

            clone.startOffset = this.startOffset;
            clone.endOffset = this.endOffset;

            return clone;
        }

        #region Properties

        internal int StartOffset
        {
            get
            {
                return startOffset;
            }
            set
            {
                startOffset = value;
            }
        }

        internal int EndOffset
        {
            get
            {
                return endOffset;
            }
            set
            {
                endOffset = value;
            }
        }

        #endregion
    }
}
