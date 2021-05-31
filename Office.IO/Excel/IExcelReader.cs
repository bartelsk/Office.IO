using BartelsOnline.Office.IO.Excel.Models;
using System.Collections.Generic;

namespace BartelsOnline.Office.IO.Excel
{
   /// <summary>
   /// This is the interface for the ExcelReader class.
   /// </summary>
   public interface IExcelReader
   {
      /// <summary>
      /// Returns a cell value from a sheet.
      /// </summary>
      /// <param name="sheetName">The sheet name.</param>
      /// <param name="cellAddress">The cell address.</param>
      /// <returns>The cell value as a string type.</returns>
      string ReadCell(string sheetName, string cellAddress);

      /// <summary>
      /// Returns a cell value from a sheet.
      /// </summary>
      /// <param name="sheetNumber">The sheet number.</param>
      /// <param name="cellAddress">The cell address.</param>
      /// <returns>The cell value as a string type.</returns>
      string ReadCell(int sheetNumber, string cellAddress);

      /// <summary>
      /// Returns range values from a sheet.
      /// </summary>
      /// <param name="sheetName">The sheet name.</param>
      /// <param name="rangeAddress">The range address in 'A1:B2' format.</param>
      /// <returns>A List with row and column data of the specified address.</returns>
      List<List<XlsRange>> ReadRange(string sheetName, string rangeAddress);

      /// <summary>
      /// Returns range values from a sheet.
      /// </summary>
      /// <param name="sheetNumber">The sheet number.</param>
      /// <param name="rangeAddress">The range address in 'A1:B2' format.</param>
      /// <returns>A List with row and column data of the specified address.</returns>
      List<List<XlsRange>> ReadRange(int sheetNumber, string rangeAddress);
   }
}
