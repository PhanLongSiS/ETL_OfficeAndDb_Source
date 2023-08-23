using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelExcelJson
{
    public class CellPosition
    {
        public int RowIndex;
        public int ColumnIndex;

        public CellPosition()
        {

        }

        public CellPosition(string position)
        {
            StringAddressToNumber(position, ref this.ColumnIndex, ref this.RowIndex);
        }

        /// <summary>
        ///    Chuyển đổi vị trí cell về tọa độ cột, dòng
        /// </summary>
        /// <param name="colrow_name">Ví dụ cell dạng gợi nhớ. Ví dụ A7.</param>
        /// <param name="ColumnIndex">Chỉ sô cột. Ví dụ 1</param>
        /// <param name="RowIndex">Chỉ số dòng, Ví dụ 7. </param>
        static public void StringAddressToNumber(string colrow_name, ref int ColumnIndex, ref int RowIndex)
        {
            if (colrow_name == null || colrow_name == string.Empty) return;
            int splitPos = colrow_name.IndexOfAny("0123456789".ToCharArray());
            string ColumnName = colrow_name.Substring(0, splitPos);
            string RowName = colrow_name.Substring(ColumnName.Length, colrow_name.Length - splitPos);
            ColumnIndex = ColumnNameToNumber(ColumnName);
            RowIndex = Convert.ToInt32(RowName);
        }

        // Return the column number for this column name.
        static int ColumnNameToNumber(string col_name)
        {
            int result = 0;
            // Process each letter.
            for (int i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result;
        }
    }
}
