using ExcelDna.Integration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonExcel
{
    public static class ArrayFunctions
    {
        [ExcelFunction("Sort Array", Category = "Array",Name ="Sort2", IsExceptionSafe = true)]
        public static object Sort(
            [ExcelArgument(Description = "Excel input range", Name ="Input")] object[,] range,
            [ExcelArgument(Description = "Sort column", Name = "Sort Index")] int sortIndex = 0)
        {
            try
            {
                SortedList list = new SortedList();

                for (int i = range.GetLowerBound(0); i <= range.GetUpperBound(0); i++)
                {
                    if (!list.Contains(range[i, sortIndex]))
                        list.Add(range[i, sortIndex], range[i, 0]);
                }

                object[,] array = new object[list.Count, 1];
                int j = 0;
                foreach (var e in list.Values)
                {
                    array[j, 0] = e;
                    j++;
                }
                return array;
            }
            catch (Exception ex)
            {
                return ex.Message;  // ExcelError.ExcelErrorNA;
            }
        }
        [ExcelFunction("DeDup Array", Category = "Array", Name = "Unique", IsExceptionSafe = true)]
        public static object Unique(
            [ExcelArgument(Description = "Excel input range", Name = "Unique")] object[,] range)
        {
            try
            {
                ArrayList list = new ArrayList();

                for (int i = range.GetLowerBound(0); i <= range.GetUpperBound(0); i++)
                {
                    if (!list.Contains(range[i, 0]))
                        list.Add(range[i, 0]);
                }

                object[,] array = new object[list.Count, 1];
                int j = 0;
                foreach (var e in list)
                {
                    array[j, 0] = e;
                    j++;
                }
                return array;
            }
            catch (Exception ex)
            {
                return ex.Message;  // ExcelError.ExcelErrorNA;
            }
        }
    }
}
