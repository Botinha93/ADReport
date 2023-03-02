using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Security;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.IO;

namespace ADReport
{
    public static class DirectoryEntryExtension
    {
        [ComImport, Guid("9068270b-0939-11d1-8be1-00c04fd8d503"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
        internal interface IAdsLargeInteger
        {
            long HighPart
            {
                [SuppressUnmanagedCodeSecurity]
                get; [SuppressUnmanagedCodeSecurity]
                set;
            }

            long LowPart
            {
                [SuppressUnmanagedCodeSecurity]
                get; [SuppressUnmanagedCodeSecurity]
                set;
            }
        }
        public static String Property(this DirectoryEntry de, String Property)
        {
            if (de.Properties[Property].Count > 0)
            {
                if (Convert.ToString(de.Properties[Property].Value).Contains("ComObjec"))
                {
                    try
                    {
                        IAdsLargeInteger largeInt = (IAdsLargeInteger)de.Properties[Property][0];
                        long datelong = (((long)largeInt.HighPart) << 32) + largeInt.LowPart;
                        return DateTime.FromFileTime(datelong).ToString("dd/MM/yyyy");
                    }
                    catch
                    {
                        return "UmparsableObject";
                    }
                }
                else
                {
                    return Convert.ToString(de.Properties[Property].Value);
                }
            }
            return String.Empty;
        }
        public static Dictionary<string, string> GetProperties(this DirectoryEntry de)
        {
            String[] propertis = new String[1];
            Dictionary<string, string> dic = new Dictionary<string, string>();
            var en = de.Properties.GetEnumerator();
            while (en.MoveNext()) {
                dic.Add(Convert.ToString(en.Key), de.Property(Convert.ToString(en.Key)));
            }
            return dic;
        }
        public static Dictionary<string, string> GetProperties(this DirectoryEntry de, String[] properties)
        {
            if (properties.Length > 0)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                foreach (var i in properties)
                {
                    dic.Add(i, de.Property(i));
                }
                return dic;
            }
            else
            {
                return de.GetProperties();
            }

        }

    }
}

