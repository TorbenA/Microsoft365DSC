using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace Microsoft365DSC.Utilities
{
    public static class Utilities
    {
        /// <summary>
        /// Method to update special characters in strings.
        /// This method handles the conversion of special characters similar to Update-M365DSCSpecialCharacters.
        /// This function updates special characters in a string to be escaped in a DSC configuration.
        /// The function replaces the following characters:
        ///     - 0x201C = “
        ///     - 0x201D = ”
        ///     - 0x201E = „
        /// </summary>
        /// <param name="input">The input string to process</param>
        /// <returns>The processed string with special characters updated</returns>
        public static string UpdateSpecialCharacters(string input)
        {
            input = input.Replace(((char)0x201C).ToString(), "`" + ((char)0x201C).ToString());
            input = input.Replace(((char)0x201D).ToString(), "`" + ((char)0x201D).ToString());
            input = input.Replace(((char)0x201E).ToString(), "`" + ((char)0x201E).ToString());
            return input;
        }

        public static Array FilterHashtablesByResourceAndKey(IEnumerable<object> hashtables, string resourceName, string key, string keyValue)
        {
            List<Hashtable> results = new List<Hashtable>();
            foreach (Hashtable entry in hashtables.Cast<Hashtable>())
            {
                if (entry.ContainsKey("ResourceName") && entry["ResourceName"].ToString() == resourceName &&
                    entry.ContainsKey(key) && entry[key].ToString() == keyValue)
                {
                    results.Add(entry);
                }
            }
            return results.ToArray();
        }

        public static object FilterCimClassesByName(IEnumerable<object> schemaObjects, string className)
        {
            List<object> results = new List<object>();
            foreach (var entry in schemaObjects.Cast<PSObject>())
            {
                dynamic dyn = entry as dynamic;
                string name = dyn.ClassName;
                if (name == className)
                {
                    results.Add(entry);
                }
            }
            return results.ToArray();
        }
    }
}
