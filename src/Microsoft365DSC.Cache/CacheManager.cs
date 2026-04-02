using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Microsoft365DSC.Cache
{
    /// <summary>
    /// Provides a centralized, thread-safe cache for expensive data structures
    /// shared across PowerShell and C# layers. Loading data here avoids repeated
    /// deserialization and the boxing overhead that occurs when passing rich
    /// objects from PowerShell to C#.
    /// </summary>
    public static class CacheManager
    {
        private static readonly object SchemaLock = new();
        private static List<object> _schema;

        /// <summary>
        /// Gets a value indicating whether the M365DSC schema has been loaded.
        /// </summary>
        public static bool IsSchemaLoaded
        {
            get
            {
                lock (SchemaLock)
                {
                    return _schema != null;
                }
            }
        }

        /// <summary>
        /// Gets the cached schema as an <see cref="IEnumerable{Object}"/> of Hashtables.
        /// Returns <c>null</c> if the schema has not been loaded yet.
        /// </summary>
        public static IEnumerable<object> Schema
        {
            get
            {
                lock (SchemaLock)
                {
                    return _schema;
                }
            }
        }

        /// <summary>
        /// Loads the SchemaDefinition.json file from <paramref name="filePath"/>,
        /// deserializes it into a list of <see cref="Hashtable"/> trees, and caches
        /// the result. Subsequent calls are no-ops unless <paramref name="force"/>
        /// is <c>true</c>.
        /// </summary>
        /// <param name="filePath">Absolute path to SchemaDefinition.json.</param>
        /// <param name="force">When <c>true</c>, reloads even if already cached.</param>
        public static void LoadSchema(string filePath, bool force = false)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(nameof(filePath));

            lock (SchemaLock)
            {
                if (_schema != null && !force)
                    return;

                string json = File.ReadAllText(filePath);
                using JsonDocument doc = JsonDocument.Parse(json);
                _schema = ConvertJsonArray(doc.RootElement);
            }
        }

        /// <summary>
        /// Clears the cached schema, allowing it to be reloaded on next access.
        /// Primarily useful for testing.
        /// </summary>
        public static void ClearSchema()
        {
            lock (SchemaLock)
            {
                _schema = null;
            }
        }

        #region JSON to Hashtable conversion

        /// <summary>
        /// Converts a <see cref="JsonElement"/> array into a <see cref="List{Object}"/>
        /// of Hashtables, preserving the structure expected by ResourceComparer's
        /// GetProperty / GetStringProperty helpers.
        /// </summary>
        private static List<object> ConvertJsonArray(JsonElement element)
        {
            var list = new List<object>();
            foreach (JsonElement item in element.EnumerateArray())
            {
                list.Add(ConvertJsonElement(item));
            }
            return list;
        }

        private static object ConvertJsonElement(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    var table = new Hashtable(StringComparer.OrdinalIgnoreCase);
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        table[prop.Name] = ConvertJsonElement(prop.Value);
                    }
                    return table;

                case JsonValueKind.Array:
                    var items = new List<object>();
                    foreach (JsonElement child in element.EnumerateArray())
                    {
                        items.Add(ConvertJsonElement(child));
                    }
                    return items;

                case JsonValueKind.String:
                    return element.GetString();

                case JsonValueKind.Number:
                    if (element.TryGetInt64(out long longVal))
                        return longVal;
                    return element.GetDouble();

                case JsonValueKind.True:
                    return true;

                case JsonValueKind.False:
                    return false;

                case JsonValueKind.Null:
                default:
                    return null;
            }
        }

        #endregion
    }
}
