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
            [ExcelArgument(Description = "Excel input range", Name ="Input")] object[,] r,
            [ExcelArgument(Description = "Sort Order\n0 - Ascending (default)\n1 - descending", Name = "Order")] int sortOrder = 0,
            [ExcelArgument(Description = "Sort Index\n0 - First Column (default)\n1-n another column", Name = "Index")] int sortIndex = 0
        )
        {
            try
            {
                var list = new List<Tuple<IComparable, int>>();  // Sort should not DeDup, use a different method

                for (int i = r.GetLowerBound(sortIndex); i <= r.GetUpperBound(sortIndex); i++)
                {
                    list.Add(Tuple.Create(r[i,sortIndex] as IComparable, i));
                }

                int internalSortOrder = sortOrder == 0 ? 1 : -1;

                list.Sort((x, y) => {
                    IComparable x1 = x.Item1;
                    IComparable y1 = y.Item1;

                    if (x1 == null) return internalSortOrder;
                    if (y1 == null) return -internalSortOrder;

                    return x1.CompareTo(y1) * internalSortOrder;
                });

                object[,] array = new object[r.GetUpperBound(0) - r.GetLowerBound(0)+1, r.GetUpperBound(1) - r.GetLowerBound(1) + 1];
                int j = r.GetLowerBound(0);
                for (int i = r.GetLowerBound(sortIndex); i <= r.GetUpperBound(sortIndex); i++)
                {
                    array[j, 0] = list[i].Item1;
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
