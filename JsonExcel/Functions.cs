using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonExcel
{
    public static class Functions
    {
        private static Dictionary<string, JObject> _deserializeCache = new Dictionary<string, JObject>();

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
                if (!_deserializeCache.TryGetValue(json, out JObject jsonLinq))
                {
                    jsonLinq = JObject.Parse(json);
                    _deserializeCache[json] = jsonLinq;
                }

                IEnumerable<JToken> jTokens = jsonLinq.Descendants().Where(p => p.Count() == 0);
                Dictionary<string, object> results = jTokens.Aggregate(new Dictionary<string, object>(), (properties, jToken) =>
                {
                    properties.Add(jToken.Path, jToken);
                    return properties;
                });

                object[,] array = new object[results.Count, 2];
                int i = 0;
                foreach (var e in results) {
                    array[i, 0] = e.Key;
                    array[i, 1] = e.Value.ToString();
                    i++;
                }

                return orientation == 0 ? array : array.Transpose();
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
                if (!_deserializeCache.TryGetValue(json, out JObject jo))
                {
                    jo = JObject.Parse(json);
                    _deserializeCache[json] = jo;
                }

                return jo[key].ToString();
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
