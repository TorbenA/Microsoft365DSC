using Microsoft.Management.Infrastructure;
using Microsoft365DSC.Converter;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Microsoft365DSC.Compare
{
    public static class SimpleObjectComparer
    {
        public static Dictionary<string, object> Compare(
            Hashtable currentValues,
            object desiredValues,
            ICollection valuesToCheck,
            Hashtable? includedDrifts = null,
            bool noEventMessage = false,
            bool noDriftReset = false,
            List<string>? excludedProperties = null)
        {
            // Contains the strings to write to the event log
            Hashtable driftedParameters = new();
            // Contains the drift 
            Hashtable driftObject = new()
            {
                { "DriftInfo", new List<Hashtable>() },
                { "CurrentValues", new Hashtable() },
                { "DesiredValues", new Hashtable() }
            };

            // The final return value for the method
            bool returnValue = false;

            if (includedDrifts is not null && includedDrifts.Keys.Count > 0)
            {
                driftedParameters = includedDrifts;
                foreach (DictionaryEntry existingDrift in includedDrifts)
                {
                    string propertyName = (string)existingDrift.Key;
                    string value = (string)existingDrift.Value;
                    int startIndex = value.IndexOf("</CurrentValue>");
                    string currentValue = value.Substring(0, startIndex)
                        .Replace("<CurrentValue>", "");
                    string desiredValue = value.Substring(startIndex + 15, value.Length - (startIndex + 15))
                        .Replace("<DesiredValue>", "")
                        .Replace("</DesiredValue>", "");

                    AddDriftInfo(driftObject, propertyName, currentValue, desiredValue);
                }
            }

            string desiredTypeName = desiredValues.GetType().Name;
            // Match for Hashtable, CimInstance and PSBoundParametersDictionary, which inherits from Dictionary<string, object>
            if (desiredValues is not Hashtable && desiredValues is not CimInstance && desiredValues is not Dictionary<string, object>)
            {
                throw new ArgumentException($"Property 'DesiredValues' must be either a Hashtable or CimInstance. Type detected was {desiredValues.GetType().FullName}");
            }

            if (desiredValues is CimInstance && valuesToCheck is null)
            {
                throw new ArgumentException("If 'DesiredValues' is a CimInstance, then property 'ValuesToCheck' must contain a value");
            }

            List<string> keyList = valuesToCheck.Cast<string>().ToList();
            Hashtable desiredValuesHashtable = ComplexObjectConverter.ToHashtable(desiredValues);

            // Add default Ensure value if it is not present in the DesiredValues but present in the CurrentValues
            if (!keyList.Contains("Ensure") && !keyList.Contains("IsSingleInstance") && currentValues.ContainsKey("Ensure"))
            {
                keyList.Add("Ensure");
                if (desiredValuesHashtable is not null && !desiredValuesHashtable.ContainsKey("Ensure"))
                {
                    desiredValuesHashtable.Add("Ensure", "Present");
                }
            }

            List<string> propertiesToExclude = ["Verbose", "Credential", "ApplicationId", "CertificateThumbprint", "CertificatePath", "CertificatePassword", "TenantId", "ApplicationSecret", "ManagedIdentity", "AccessTokens"];
            if (excludedProperties is not null)
            {
                propertiesToExclude.AddRange(excludedProperties);
                propertiesToExclude = propertiesToExclude.Distinct().ToList();
            }

            foreach (string key in keyList)
            {
                if (propertiesToExclude.Contains(key))
                {
                    continue;
                }

                if (!currentValues.ContainsKey(key) ||
                    !(currentValues[key]?.Equals(desiredValuesHashtable[key]) ?? false) ||
                        (desiredValuesHashtable.ContainsKey(key) && 
                            desiredValuesHashtable[key] is not null && desiredValuesHashtable[key] is Array))
                {
                    if (desiredValuesHashtable.ContainsKey(key))
                    {
                        object? desiredValue = desiredValuesHashtable[key];
                        Type desiredType;
                        if (desiredValue is null)
                        {
                            desiredType = currentValues[key]?.GetType() ?? typeof(object);
                        }
                        else
                        {
                            desiredType = desiredValue.GetType();
                        }

                        if (desiredType.IsArray)
                        {
                            if (!currentValues.ContainsKey(key) || currentValues[key] is null)
                            {
                                AddDriftInfo(driftObject, key, null, desiredValue);
                                AddDriftedParameter(driftedParameters, key, "null", "[Array]");
                                returnValue = false;
                                continue;
                            }
                            
                            if (desiredType.Name.Equals("CimInstance[]"))
                            {
                                // Do nothing because it's already handled previously through Compare-M365DSCComplexObject -> ComplexObjectComparer
                                // Complex properties are removed and thus not handled by this function
                                // However, since we might miss something during developing, throw an error
                                throw new NotSupportedException($"Comparing CimInstances with {typeof(SimpleObjectComparer).Name} is not supported.");
                            }

                            Array desiredValuesArray = desiredValue is null
                                ? Array.CreateInstance(desiredType, 0)
                                : (Array)desiredValue;
                            Array currentValuesArray = (Array)currentValues[key];
                            var arrayDifferences = ArrayComparer.CompareArrays(currentValuesArray, desiredValuesArray);
                            if (arrayDifferences.Count > 0)
                            {
                                string currentValuesString = string.Join(", ", currentValuesArray.Cast<string>());
                                string desiredValuesString = string.Join(", ", desiredValuesArray.Cast<string>());
                                AddDriftInfo(driftObject, key, currentValuesString, desiredValuesString);
                                AddDriftedParameter(driftedParameters, key, currentValuesString, desiredValuesString);
                                returnValue = false;
                            }
                            continue;
                        }

                        switch (desiredValue)
                        {
                            case string desiredValueString:
                                string currentValueString = currentValues[key] as string;
                                if (!string.IsNullOrEmpty(currentValueString))
                                {
                                    currentValueString = ReplaceLineBreaks(currentValueString);
                                }
                                desiredValueString = ReplaceLineBreaks(desiredValueString);

                                if (string.IsNullOrEmpty(currentValueString) && string.IsNullOrEmpty(desiredValueString))
                                {
                                    // Do nothing - Strings are equally empty
                                    continue;
                                }

                                if (!string.IsNullOrEmpty(currentValueString) && !string.IsNullOrEmpty(desiredValueString)
                                    && string.Equals(desiredValueString, currentValueString, StringComparison.Ordinal))
                                {
                                    // Do nothing - Strings are the same
                                    continue;
                                }

                                AddDriftInfo(driftObject, key, currentValueString, desiredValueString);
                                AddDriftedParameter(driftedParameters, key, currentValueString, desiredValueString);
                                returnValue = false;
                                break;

                            case int desiredValueInt32:
                                AddDriftInfo(driftObject, key, currentValues[key] as string ?? string.Empty, desiredValueInt32.ToString());
                                AddDriftedParameter(driftedParameters, key, currentValues[key] as string ?? string.Empty, desiredValueInt32.ToString());
                                returnValue = false;
                                break;

                            case uint desiredValueUint32:
                                AddDriftInfo(driftObject, key, currentValues[key] as string ?? string.Empty, desiredValueUint32.ToString());
                                AddDriftedParameter(driftedParameters, key, currentValues[key] as string ?? string.Empty, desiredValueUint32.ToString());
                                returnValue = false;
                                break;

                            case short desiredValueInt16:
                                AddDriftInfo(driftObject, key, currentValues[key] as string ?? string.Empty, desiredValueInt16.ToString());
                                AddDriftedParameter(driftedParameters, key, currentValues[key] as string ?? string.Empty, desiredValueInt16.ToString());
                                returnValue = false;
                                break;

                            case bool desiredValueBool:
                                AddDriftInfo(driftObject, key, currentValues[key] as string ?? string.Empty, desiredValueBool.ToString());
                                AddDriftedParameter(driftedParameters, key, currentValues[key] as string ?? string.Empty, desiredValueBool.ToString());
                                returnValue = false;
                                break;

                            case float desiredValueSingle:
                                AddDriftInfo(driftObject, key, currentValues[key] as string ?? string.Empty, desiredValueSingle.ToString());
                                AddDriftedParameter(driftedParameters, key, currentValues[key] as string ?? string.Empty, desiredValueSingle.ToString());
                                returnValue = false;
                                break;

                            case Hashtable:
                                throw new NotSupportedException($"Comparing Hashtables with {typeof(SimpleObjectComparer).Name} is not supported.");

                            default:
                                throw new NotSupportedException($"Comparing {desiredType.FullName} with {typeof(SimpleObjectComparer).Name} is not supported.");
                        }
                    }
                }
            }

            return new()
            {
                { "TestResult", returnValue },
                { "DriftObject", driftObject },
                { "DriftedParameters", driftedParameters }
            };
        }

        private static void AddDriftInfo(Hashtable driftObject, string propertyName, object currentValue, object desiredValue)
        {
            ((List<Hashtable>)driftObject["DriftInfo"]).Add(new Hashtable()
            {
                { "PropertyName", propertyName },
                { "CurrentValue", currentValue },
                { "DesiredValue", desiredValue }
            });
        }

        private static void AddDriftedParameter(Hashtable driftedParameters, string propertyName, string currentValue, string desiredValue)
        {
            string eventValue = $"<CurrentValue>{currentValue}</CurrentValue>";
            eventValue += $"<DesiredValue>{desiredValue}</DesiredValue>";
            driftedParameters.Add(propertyName, eventValue);
        }

        private static string ReplaceLineBreaks(string text)
        {
            return text.Replace("\r\n", "\n");
        }
    }
}
