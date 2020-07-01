using System;

namespace Independentsoft.IO.StructuredStorage
{
    internal class EnumUtil
    {
        internal static byte ParseColor(Color color)
        {
            if (color == Color.Red)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        internal static Color ParseColor(byte color)
        {
            if (color == 0)
            {
                return Color.Red;
            }
            else
            {
                return Color.Black;
            }
        }

        internal static byte ParseDirectoryEntryType(DirectoryEntryType type)
        {
            if (type == DirectoryEntryType.Root)
            {
                return 5;
            }
            else if (type == DirectoryEntryType.Property)
            {
                return 4;
            }
            else if (type == DirectoryEntryType.LockBytes)
            {
                return 3;
            }
            else if (type == DirectoryEntryType.Stream)
            {
                return 2;
            }
            else if (type == DirectoryEntryType.Storage)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        internal static DirectoryEntryType ParseDirectoryEntryType(byte type)
        {
            if (type == 5)
            {
                return DirectoryEntryType.Root;
            }
            else if (type == 4)
            {
                return DirectoryEntryType.Property;
            }
            else if (type == 3)
            {
                return DirectoryEntryType.LockBytes;
            }
            else if (type == 2)
            {
                return DirectoryEntryType.Stream;
            }
            else if (type == 1)
            {
                return DirectoryEntryType.Storage;
            }
            else
            {
                return DirectoryEntryType.Invalid;
            }
        }
    }
}
