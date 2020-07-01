using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Enum PropertyType
    /// </summary>
    public enum PropertyType
    {
        //Fixed length properties
        Integer16,
        Integer32,
        Floating32,
        Floating64,
        Currency,
        FloatingTime,
        ErrorCode,
        Boolean,
        Integer64,
        Time,

        //Variable length properties
        String,
        Binary,
        String8,
        Guid,
        Object,

        //Fixed length multi-valued properties
        MultipleInteger16,
        MultipleInteger32,
        MultipleFloating32,
        MultipleFloating64,
        MultipleCurrency,
        MultipleFloatingTime,
        MultipleInteger64,
        MultipleGuid,
        MultipleTime,

        //Variable length multi-valued properties
        MultipleString,
        MultipleBinary,
        MultipleString8
    }
}
