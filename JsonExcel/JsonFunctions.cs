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
    public static class JsonFunctions
    {
        private static Dictionary<string, JContainer> _deserializeCache = new Dictionary<string, JContainer>();
        private static Dictionary<JContainer, Dictionary<string,object>> _flatternCache = new Dictionary<JContainer, Dictionary<string, object>>();

        [ExcelFunction("Convert an Excel Range to a JSON string", 
            Category = "Json Excel",
            IsExceptionSafe = true)]
        public static object JsonFromCells(
            [ExcelArgument(Description = "Excel input range")] object[,] range)
        {
            try
            {
                Dictionary<object, object> dic = new Dictionary<object, object>();
                for (int i = range.GetLowerBound(0); i <= range.GetUpperBound(0); i++)
                {
                    dic[range[i, 0]] = range[i, 1];

                    for (int j = range.GetLowerBound(1); j <= range.GetUpperBound(1); j++)
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
                if (!_deserializeCache.TryGetValue(json, out JContainer jc))
                {
                    if (json.StartsWith("{"))
                        jc = JObject.Parse(json);
                    else if (json.StartsWith("["))
                        jc = JArray.Parse(json);
                    else
                        throw new ArgumentException("Not JSON");

                    _deserializeCache[json] = jc;
                }

                if (!_flatternCache.TryGetValue(jc, out Dictionary<string, object> results))
                {
                    IEnumerable<JToken> jTokens = jc.Descendants().Where(p => p.Count() == 0);
                    results = jTokens.Aggregate(new Dictionary<string, object>(), (properties, jToken) =>
                    {
                        properties.Add(jToken.Path, jToken);
                        return properties;
                    });
                    _flatternCache[jc] = results;
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
                if (!_deserializeCache.TryGetValue(json, out JContainer jc))
                {
                    if (json.StartsWith("["))
                        jc = JArray.Parse(json);
                    else if (json.StartsWith("{"))
                        jc = JObject.Parse(json);
                    else
                        throw new ArgumentException("Not JSON");

                    _deserializeCache[json] = jc;
                }

                if (jc is JObject && (jc as JObject).ContainsKey(key))
                    return ToExcelVal(jc[key]);

                if (jc is JArray && Convert.ToInt32(key) < (jc as JArray).Count)
                    return ToExcelVal(jc[Convert.ToInt32(key)]);

                if (!_flatternCache.TryGetValue(jc, out Dictionary<string, object> results))
                {
                    IEnumerable<JToken> jTokens = jc.Descendants().Where(p => p.Count() == 0);
                    results = jTokens.Aggregate(new Dictionary<string, object>(), (properties, jToken) =>
                    {
                        properties.Add(jToken.Path, jToken);
                        return properties;
                    });
                    _flatternCache[jc] = results;
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
            {
                var jt = val as JToken;
                if (!jt.HasValues)  // no children
                    return jt.ToObject<object>();
            }

            return val.ToString();
        }

    }
}
