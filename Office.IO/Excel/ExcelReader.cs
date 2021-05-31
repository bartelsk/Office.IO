using BartelsOnline.Office.IO.Excel.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BartelsOnline.Office.IO.Excel
{
   /// <summary>
   /// This class contains methods to easily read Excel cell and range values.
   /// </summary>
   public class ExcelReader : IExcelReader, IDisposable
   {
      private readonly SpreadsheetDocument _document;
      private readonly WorkbookPart _workbookPart;

      private bool disposedValue;

      #region Constructor / Destructor

      /// <summary>
      /// Initializes a new instance of the BartelsOnline.Office.IO.Excel.ExcelReader class.
      /// </summary>
      /// <param name="fileName">The filename of the Excel workbook.</param>
      public ExcelReader(string fileName)
      {
         _document = SpreadsheetDocument.Open(fileName, false);
         _workbookPart = _document.WorkbookPart;
      }

      #pragma warning disable CS1591 
      ~ExcelReader()
      
      {
         Dispose(disposing: false);
      }

      protected virtual void Dispose(bool disposing)
      {
         if (!disposedValue)
         {
            if (disposing)
            {
               if (_document != null)
               {
                  _document.Close();
                  _document.Dispose();
               }
            }
            disposedValue = true;
         }
      }

      public void Dispose()
      {
         Dispose(disposing: true);
         GC.SuppressFinalize(this);
      }
      #pragma warning restore CS1591

      #endregion

      #region Public methods

      /// <summary>
      /// Returns a cell value from a sheet.
      /// </summary>
      /// <param name="sheetName">The sheet name.</param>
      /// <param name="cellAddress">The cell address.</param>
      /// <returns>The cell value as a string type.</returns>
      public string ReadCell(string sheetName, string cellAddress)
      {
         string cellValue = string.Empty;
         Sheet theSheet = GetSheet(sheetName);
         if (theSheet != null)
         {
            cellValue = GetCellValue(theSheet, cellAddress);
         }
         return cellValue;
      }

      /// <summary>
      /// Returns a cell value from a sheet.
      /// </summary>
      /// <param name="sheetNumber">The sheet number.</param>
      /// <param name="cellAddress">The cell address.</param>
      /// <returns>The cell value as a string type.</returns>
      public string ReadCell(int sheetNumber, string cellAddress)
      {
         string cellValue = string.Empty;
         Sheet theSheet = GetSheet(sheetNumber);
         if (theSheet != null)
         {
            cellValue = GetCellValue(theSheet, cellAddress);
         }
         return cellValue;
      }

      /// <summary>
      /// Returns range values from a sheet.
      /// </summary>
      /// <param name="sheetName">The sheet name.</param>
      /// <param name="rangeAddress">The range address in 'A1:B2' format.</param>
      /// <returns>A List with row and column data of the specified address.</returns>
      public List<List<XlsRange>> ReadRange(string sheetName, string rangeAddress)
      {
         List<List<XlsRange>> rowData = null;
         Sheet theSheet = GetSheet(sheetName);
         if (theSheet != null)
         {
            rowData = ReadWorksheetRange(theSheet, rangeAddress);
         }                 
         return rowData;
      }

      /// <summary>
      /// Returns range values from a sheet.
      /// </summary>
      /// <param name="sheetNumber">The sheet number.</param>
      /// <param name="rangeAddress">The range address in 'A1:B2' format.</param>
      /// <returns>A List with row and column data of the specified address.</returns>
      public List<List<XlsRange>> ReadRange(int sheetNumber, string rangeAddress)
      {
         List<List<XlsRange>> rowData = null;
         Sheet theSheet = GetSheet(sheetNumber);
         if (theSheet != null)
         {
            rowData = ReadWorksheetRange(theSheet, rangeAddress);
         }         
         return rowData;
      }

      #endregion

      #region Private methods

      /// <summary>
      /// Get a sheet reference.
      /// </summary>
      /// <param name="id">The sheet number (first sheet is 1).</param>
      /// <returns>A Sheet type.</returns>
      private Sheet GetSheet(int id)
      {
         return _workbookPart.Workbook.Descendants<Sheet>().ElementAt(id - 1);
      }

      /// <summary>
      /// Get a sheet reference.
      /// </summary>
      /// <param name="name">The sheet name.</param>
      /// <returns>A Sheet type.</returns>
      private Sheet GetSheet(string name)
      {
         return _workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == name).FirstOrDefault();
      }

      /// <summary>
      /// Get a row reference.
      /// </summary>
      /// <param name="sheet">A sheet type.</param>
      /// <param name="rowIndex">The row index.</param>
      /// <returns>A Row type.</returns>
      private Row GetRow(Sheet sheet, int rowIndex)
      {
         return sheet.GetFirstChild<SheetData>().Elements<Row>()
                .Where(r => r.RowIndex == rowIndex).FirstOrDefault();
      }

      /// <summary>
      /// Gets a cell reference.
      /// </summary>
      /// <param name="sheet">A sheet type.</param>
      /// <param name="address">The cell address.</param>
      /// <returns>A Cell type.</returns>
      private Cell GetCell(Sheet sheet, string address)
      {
         WorksheetPart wsPart = (WorksheetPart)(_workbookPart.GetPartById(sheet.Id));
         return wsPart.Worksheet.Descendants<Cell>()
                .Where(c => c.CellReference == address).FirstOrDefault();
      }

      /// <summary>
      /// Gets a cell reference.
      /// </summary>
      /// <param name="sheet">A sheet type.</param>
      /// <param name="rowIndex">The row index (first row is 1).</param>
      /// <param name="columnIndex">The column index (first column is 1).</param>
      /// <returns>A Cell type.</returns>
      private Cell GetCell(Sheet sheet, int rowIndex, int columnIndex)
      {
         WorksheetPart wsPart = (WorksheetPart)(_workbookPart.GetPartById(sheet.Id));
         return wsPart.Worksheet.Descendants<Cell>()
                .Where(c => c.CellReference == GetColumnName(columnIndex) + rowIndex).FirstOrDefault();
      }

      /// <summary>
      /// Returns a cell value.
      /// </summary>
      /// <param name="sheet">A sheet.</param>
      /// <param name="address">The cell address.</param>
      /// <returns>The cell value as a string type.</returns>
      private string GetCellValue(Sheet sheet, string address)
      {
         Cell cell = GetCell(sheet, address);
         return GetCellValue(cell);
      }      

      /// <summary>
      /// Returns the column name.
      /// </summary>
      /// <param name="columnIndex">The column index.</param>
      /// <returns>The column name.</returns>
      private static string GetColumnName(int columnIndex)
      {
         int modifier;
         int dividend = columnIndex;
         string columnName = string.Empty;

         while (dividend > 0)
         {
            modifier = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
            dividend = (dividend - modifier) / 26;
         }

         return columnName;
      }

      /// <summary>
      /// Returns the column index.
      /// </summary>
      /// <param name="columnName">The column name.</param>
      /// <returns>The column index.</returns>
      private static int GetColumnIndex(string columnName)
      {
         int sum = 0;
         columnName = columnName.ToUpperInvariant();

         for (int i = 0; i < columnName.Length; i++)
         {
            sum *= 26;
            sum += (columnName[i] - 'A' + 1);
         }

         return sum;
      }

      /// <summary>
      /// Converts a cell address like 'B12' into a CellAddress type.
      /// </summary>
      /// <param name="address">The cell address.</param>
      /// <returns>A CellAdress type.</returns>
      private static CellAddress GetCellAddress(string address)
      {
         int startIndex = address.IndexOfAny("0123456789".ToCharArray());
         if (startIndex > -1)
         {
            return new CellAddress()
            {
               ColumnIndex = GetColumnIndex(address.Substring(0, startIndex)),
               ColumnName = address.Substring(0, startIndex),
               RowIndex = int.Parse(address.Substring(startIndex))
            };
         }
         else
         {
            return new CellAddress() { ColumnName = address, ColumnIndex = -1, RowIndex = 1 };
         }
      }

      /// <summary>
      /// Returns a cell value.
      /// </summary>      
      /// <param name="cell">A cell.</param>
      /// <returns>The cell value as a string type.</returns>
      private string GetCellValue(Cell cell)
      {
         string cellValue = string.Empty;
         if (cell != null && cell.InnerText.Length > 0)
         {
            // integer number
            cellValue = cell.InnerText;

            // not an integet
            if (cell.DataType != null)
            {
               switch (cell.DataType.Value)
               {
                  case CellValues.Boolean:
                     switch (cellValue)
                     {
                        case "0":
                           cellValue = "FALSE";
                           break;
                        default:
                           cellValue = "TRUE";
                           break;
                     }
                     break;
                  case CellValues.SharedString:
                     var stringTable = _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                     if (stringTable != null)
                     {
                        cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
                     }
                     break;
               }
            }
         }
         return cellValue;
      }

      /// <summary>
      /// Returns range values from a sheet.
      /// </summary>
      /// <param name="sheet">A sheet.</param>
      /// <param name="rangeAddress">A range address.</param>
      /// <returns>A List with row and column data of the specified address.</returns>
      private List<List<XlsRange>> ReadWorksheetRange(Sheet sheet, string rangeAddress)
      {
         List<List<XlsRange>> rowData = null;
         if (sheet != null)
         {
            int sepPos = rangeAddress.IndexOf(":");
            if (sepPos > 0)
            {
               rowData = new List<List<XlsRange>>();
               
               // Get range reference
               CellAddress topLeft = GetCellAddress(rangeAddress.Substring(0, sepPos));
               CellAddress bottomRight = GetCellAddress(rangeAddress.Substring(sepPos + 1));

               if (topLeft.ColumnIndex > -1)
               {
                  // Range of type "A1:B2"
                  for (int row = topLeft.RowIndex; row <= bottomRight.RowIndex; row++)
                  {
                     List<XlsRange> columnData = new();
                     for (int col = topLeft.ColumnIndex; col <= bottomRight.ColumnIndex; col++)
                     {
                        Cell theCell = GetCell(sheet, row, col);
                        XlsRange xlsRange = new(GetColumnName(col), col, row, GetColumnName(col) + row, null);

                        if (theCell != null)
                        {
                           xlsRange.Value = GetCellValue(theCell);
                        }

                        columnData.Add(xlsRange);
                     }
                     rowData.Add(columnData);
                  }
               }
               else
               {
                  // Range of type "A:B"
                  for (int row = topLeft.RowIndex; row <= 1048576; row++)
                  {
                     List<XlsRange> columnData = new();
                     for (int col = 1; col <= 200; col++)
                     {
                        Cell theCell = GetCell(sheet, row, col);
                        if (theCell != null)
                        {
                           XlsRange xlsRange = new(GetColumnName(col), col, row, GetColumnName(col) + row, GetCellValue(theCell));                           
                           columnData.Add(xlsRange);
                        }
                        else
                        {
                           // No more columns in the current region
                           break;
                        }
                     }
                     if (columnData != null && columnData.Count > 0)
                     {
                        rowData.Add(columnData);
                     }
                     else
                     {
                        // Current region is now null, so no more rows
                        break;
                     }
                  }
               }
            }
            else
            {
               throw new ArgumentException($"Invalid range address.");
            }
         }
         return rowData;
      }    

      /// <summary>
      /// Checks whether the specified string contains numbers only. 
      /// </summary>
      /// <param name="text">The text to check.</param>
      /// <returns>True if the entire string has numeric values.</returns>
      private static bool IsNumeric(string text)
      {
         return System.Text.RegularExpressions.Regex.IsMatch(text, "^[0-9]+$");
      }

      #endregion
   }
}
