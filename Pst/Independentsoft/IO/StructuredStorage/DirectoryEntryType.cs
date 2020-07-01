using System;

namespace Independentsoft.IO.StructuredStorage
{
    internal enum DirectoryEntryType
    {
        Invalid,
        Storage,
        Stream,
        LockBytes,
        Property,
        Root
    }
}
