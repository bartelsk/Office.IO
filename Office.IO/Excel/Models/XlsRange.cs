namespace BartelsOnline.Office.IO.Excel.Models
{
   /// <summary>
   /// This represents a single Excel range.
   /// </summary>
   public class XlsRange
   {
      /// <summary>
      /// The worksheet column name of this range, e.g. 'A'.
      /// </summary>
      public string ColumnName { get; set; }

      /// <summary>
      /// The worksheet column index of this range.
      /// </summary>
      public int ColumnIndex { get; set; }

      /// <summary>
      /// The worksheet row index of this range.
      /// </summary>
      public int RowIndex { get; set; }

      /// <summary>
      /// The worksheet range address.
      /// </summary>
      public string Address { get; set; }

      /// <summary>
      /// The range value as a string type.
      /// </summary>
      public string Value { get; set; }

      /// <summary>
      /// Initializes a new instance of the BartelsOnline.Office.IO.Excel.Models.XlsRange class.
      /// </summary>
      /// <param name="columnName">The worksheet column name of this range, e.g. 'A'.</param>
      /// <param name="columnIndex">The worksheet column index of this range.</param>
      /// <param name="rowIndex">The worksheet row index of this range.</param>
      /// <param name="address">The worksheet range address.</param>
      /// <param name="value">The range value as a string type.</param>
      public XlsRange(string columnName, int columnIndex, int rowIndex, string address, string value)
      {
         ColumnName = columnName;
         ColumnIndex = columnIndex;
         RowIndex = rowIndex;
         Address = address;
         Value = value;
      }
   }
}