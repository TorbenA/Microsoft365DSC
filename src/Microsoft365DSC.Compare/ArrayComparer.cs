using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft365DSC.Compare
{
    internal class ArrayComparer
    {
        public static List<CompareObjectModel> CompareArrays(Array currentArray, Array desiredArray, bool passThru = false)
        {
            List<CompareObjectModel> compareResults = [];

            if (currentArray == null || desiredArray == null)
            {
                throw new ArgumentException("Both currentValue and desiredValue must be of type Array and cannot be null.");
            }
            HashSet<object> currentSet = new (currentArray as IEnumerable<object>);
            HashSet<object> desiredSet = new (desiredArray as IEnumerable<object>);
            
            // Find items in desired but not in current
            foreach (var item in desiredSet)
            {
                if (!currentSet.Contains(item))
                {
                    compareResults.Add(new CompareObjectModel
                    {
                        SideIndicator = "<=",
                        InputObject = item
                    });
                    continue;
                }

                if (passThru)
                {
                    compareResults.Add(new CompareObjectModel
                    {
                        SideIndicator = "==",
                        InputObject = item
                    });
                }
            }
            // Find items in current but not in desired
            foreach (var item in currentSet)
            {
                if (!desiredSet.Contains(item))
                {
                    compareResults.Add(new CompareObjectModel
                    {
                        SideIndicator = "=>",
                        InputObject = item
                    });
                    continue;
                }

                if (passThru)
                {
                    compareResults.Add(new CompareObjectModel
                    {
                        SideIndicator = "==",
                        InputObject = item
                    });
                }
            }

            return compareResults;
        }
    }
}
