using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class NameToIdMap
    {
        private IDictionary<uint, NamedProperty> map = new Dictionary<uint, NamedProperty>();

        internal NameToIdMap(Table table)
        {
            Parse(table);
        }

        internal void Parse(Table table)
        {
            TableEntry guidsTableEntry   = table.GetEntry(0x0002, PropertyType.Binary);
            TableEntry entriesTableEntry = table.GetEntry(0x0003, PropertyType.Binary);
            TableEntry stringsTableEntry = table.GetEntry(0x0004, PropertyType.Binary);

            byte[] guids = null;
            byte[] entries = null;
            byte[] strings = null;

            if (guidsTableEntry != null)
            {
                guids = guidsTableEntry.ValueBuffer;
            }

            if (entriesTableEntry != null)
            {
                entries = entriesTableEntry.ValueBuffer;
            }

            if (stringsTableEntry != null)
            {
                strings = stringsTableEntry.ValueBuffer;
            }

            if (entries != null)
            {
                for (int i = 0; i < entries.Length; i += 8)
                {
                    uint entryValueOfReference = BitConverter.ToUInt32(entries, i);
                    ushort entryId = BitConverter.ToUInt16(entries, i + 4);
                    ushort entryNumber = BitConverter.ToUInt16(entries, i + 6);

                    ushort propertyIndex = (ushort)(entryId >> 1);
                    ushort hasStringValue = (ushort)(entryId << 15);

                    NamedProperty namedProperty = new NamedProperty();

                    if (entryId == 5) //PS_PUBLIC_STRINGS
                    {
                        namedProperty.Guid = new byte[16] { 41, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70 };
                    }
                    else if (entryId > 5)
                    {
                        try
                        {
                            ushort guidIndex = (ushort)((entryId >> 1) - 3);

                            byte[] guid = new byte[16];
                            System.Array.Copy(guids, guidIndex * 16, guid, 0, 16);

                            namedProperty.Guid = guid;
                        }
                        catch
                        {
                        }
                    }

                    if (hasStringValue == 0) //numerical named property
                    {
                        namedProperty.Id = entryValueOfReference;
                        namedProperty.Type = NamedPropertyType.Numerical;
                    }
                    else //string named property
                    {
                        namedProperty.Id = entryValueOfReference;
                        namedProperty.Type = NamedPropertyType.String;

                        try
                        {
                            int stringOffset = (int)entryValueOfReference;
                            uint nameLength = BitConverter.ToUInt32(strings, stringOffset);
                            
                            namedProperty.Name = System.Text.Encoding.Unicode.GetString(strings, stringOffset + 4, (int)nameLength);
                        }
                        catch
                        {
                        }
                    }

                    uint id = (uint)(entryNumber + 0x8000);

                    if (id <= 0xFFFE)
                    {
                        map.Add(id, namedProperty);
                    }

                }
            }
        }

        internal long GetId(uint namedPropertyId)
        {
            foreach(uint key in map.Keys)
            {
                NamedProperty namedProperty = map[key];

                if (namedProperty.Id == namedPropertyId)
                {
                    return key;
                }
            }

            return -1;
        }

        internal long GetId(string name, byte[] guid)
        {
            if (name != null && guid != null)
            {
                foreach (uint key in map.Keys)
                {
                    NamedProperty namedProperty = map[key];

                    if (namedProperty.Name == name && namedProperty.Guid.Length == guid.Length)
                    {
                        for (int i = 0; i < guid.Length; i++)
                        {
                            if (namedProperty.Guid[i] != guid[i])
                            {
                                return -1;
                            }
                        }

                        return key;
                    }
                }
            }

            return -1;
        }

        #region Properties

        internal IDictionary<uint, NamedProperty> Map
        {
            get
            {
                return map;
            }
        }

        #endregion
    }
}
