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
        private static Dictionary<JObject, Dictionary<string,object>> _flatternCache = new Dictionary<JObject, Dictionary<string, object>>();

        [ExcelFunction("Convert an Excel Range to a JSON string", 
            Category = "Json Excel",
            IsExceptionSafe = true)]
        public static object JsonFromCells(
            [ExcelArgument(Description = "Excel input range")] object[,] range)
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
        [ExcelFunction( "Convert a JSON string to an Excel Array", 
            Category = "Json Excel", 
            IsExceptionSafe = true)]
        public static object JsonToArray(
            [ExcelArgument(Description = "JSON input string")] string json,
            [ExcelArgument(Description = "0 - Vertical (default)\n1 - Horizontal")] int orientation=0)
        {
            try
            {
                if (!_deserializeCache.TryGetValue(json, out JObject jo))
                {
                    jo = JObject.Parse(json);
                    _deserializeCache[json] = jo;
                }

                if (!_flatternCache.TryGetValue(jo, out Dictionary<string, object> results))
                {
                    IEnumerable<JToken> jTokens = jo.Descendants().Where(p => p.Count() == 0);
                    results = jTokens.Aggregate(new Dictionary<string, object>(), (properties, jToken) =>
                    {
                        properties.Add(jToken.Path, jToken);
                        return properties;
                    });
                    _flatternCache[jo] = results;
                }

                object[,] array = new object[results.Count, 2];
                int i = 0;
                foreach (var e in results) {
                    array[i, 0] = e.Key;
                    array[i, 1] = ToExcelVal(e.Value);
                    i++;
                }

                return orientation == 0 ? array : array.Transpose();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        [ExcelFunction( "Lookup a JSON key in a string", 
            Category = "Json Excel", 
            IsExceptionSafe = true)]
        public static object JsonLookup(
            [ExcelArgument(Description = "JSON input string")] string json,
            [ExcelArgument(Description = "JSON lookup key")] string key)
        {
            try
            {
                if (!_deserializeCache.TryGetValue(json, out JObject jo))
                {
                    jo = JObject.Parse(json);
                    _deserializeCache[json] = jo;
                }
                if (jo.ContainsKey(key))
                    return ToExcelVal(jo[key]);

                if (!_flatternCache.TryGetValue(jo, out Dictionary<string, object> results))
                {
                    IEnumerable<JToken> jTokens = jo.Descendants().Where(p => p.Count() == 0);
                    results = jTokens.Aggregate(new Dictionary<string, object>(), (properties, jToken) =>
                    {
                        properties.Add(jToken.Path, jToken);
                        return properties;
                    });
                    _flatternCache[jo] = results;
                }

                return ToExcelVal(results[key]);
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

        private static object ToExcelVal(object val)
        {
            if (val is JValue)
                return (val as JValue).Value;

            if (val is JToken)
                return (val as JToken).ToObject<object>();

            return val.ToString();
        }

    }
}
