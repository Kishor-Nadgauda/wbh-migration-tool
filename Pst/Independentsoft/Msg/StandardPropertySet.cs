using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Class StandardPropertySet.
    /// </summary>
    public class StandardPropertySet
    {
        public static readonly byte[] Mapi = new byte[16] { 40, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] PublicStrings = new byte[16] { 41, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] InternetHeaders = new byte[16] { 134, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Appointment = new byte[16] { 2, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Task = new byte[16] { 3, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Address = new byte[16] { 4, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Common = new byte[16] { 8, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Note = new byte[16] { 14, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
        public static readonly byte[] Journal = new byte[16] { 10, 32, 6, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };

    }
}
