using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonExcel
{
    static class MatrixUtils
    {
        public static object[,] Transpose(this object[,] matrix)
        {
            var rows = matrix.GetLength(0);
            var columns = matrix.GetLength(1);

            var result = new object[columns, rows];

            for (var c = 0; c < columns; c++)
            {
                for (var r = 0; r < rows; r++)
                {
                    result[c, r] = matrix[r, c];
                }
            }

            return result;
        }
    }
}
