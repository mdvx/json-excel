using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Newtonsoft.Json;

namespace JsonExcel
{
    public static class Functions
    {
        private static Dictionary<string, Dictionary<object, object>> _deserializeCache = new Dictionary<string, Dictionary<object, object>>();

        [ExcelFunction(Category = "JSON", Description = "Convert an Excel Range to a JSON string", IsExceptionSafe = true)]
        public static object JsonFromCells(object[,] range)
        {
            try
            {
                Dictionary<object, object> dic = new Dictionary<object, object>();
                for (int i = range.GetLowerBound(0); i < range.GetUpperBound(0); i++)
                {

                    dic[range[i, 0]] = range[i, 1];

                    for (int j = 1; j < range.GetUpperBound(1); j++)
                    {
                        dic[range[i, 0]] = range[i, j];
                    }
                }

                return JsonConvert.SerializeObject(dic);
            }
            catch (Exception ex)
            {
                return ex.Message;  // ExcelError.ExcelErrorNA;
            }
        }
        [ExcelFunction( Category ="JSON",Description ="Convert a JSON string to and Excel Array", IsExceptionSafe = true)]
        public static object JsonToArray(string json, int orientation=0)
        {
            try
            {
                if (!_deserializeCache.TryGetValue(json,out Dictionary<object, object> dic))
                {
                    dic = JsonConvert.DeserializeObject<Dictionary<object, object>>(json);
                    _deserializeCache[json] = dic;
                }

                var arr = orientation == 0 ? new object[dic.Keys.Count, 2] :  new object[2,dic.Keys.Count];
                int i = 0;
                foreach(var e in dic)
                {
                    if (orientation == 0)
                    {
                        arr[i, 0] = e.Key;
                        arr[i, 1] = e.Value;
                    }
                    else
                    {
                        arr[0, i] = e.Key;
                        arr[1, i] = e.Value;
                    }
                    i++;
                }
                return arr;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        [ExcelFunction(Category = "JSON", Description = "lookup a JSON key in a string", IsExceptionSafe=true)]
        public static object JsonLookup(string json, string key)
        {
            try
            {
                if (!_deserializeCache.TryGetValue(json, out Dictionary<object, object> dic))
                {
                    dic = JsonConvert.DeserializeObject<Dictionary<object, object>>(json);
                    _deserializeCache[json] = dic;
                }

                return dic[key];
            }
            catch (KeyNotFoundException)
            {
                return ExcelError.ExcelErrorNA;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
