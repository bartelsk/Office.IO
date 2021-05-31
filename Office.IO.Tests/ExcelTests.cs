using BartelsOnline.Office.IO.Excel;
using BartelsOnline.Office.IO.Excel.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Office.IO.Tests
{
   [TestClass]
   public class ExcelTests
   {
      [TestMethod]
      public void ReadSingleCellValueFromSheetName()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            string result = xlsReader.ReadCell("Sheet1", "E9");
            Assert.IsTrue(result == "hh");
         }
      }

      [TestMethod]
      public void ReadSingleCellValueFromSheetNumber()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            string result = xlsReader.ReadCell(1, "B3");
            Assert.IsTrue(result == "TRUE");
         }
      }

      [TestMethod]
      public void ReadRangeValuesFromSheetName()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> cells = xlsReader.ReadRange("Sheet1", "A1:B4");
                        
            // 4 rows?
            Assert.IsTrue(cells.Count == 4);

            // loop rows in range
            foreach (var row in cells)
            {
               XlsRange firstColumn = row.Where(c => c.ColumnName == "A").Single();
               XlsRange secondColumn = row.Where(c => c.ColumnName == "B").Single();
               Assert.IsTrue(firstColumn.ColumnName == "A");
               Assert.IsTrue(secondColumn.ColumnName == "B");               
            }  
         }
      }

      [TestMethod]
      public void ReadRangeValuesFromSheetNumber()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> cells = xlsReader.ReadRange(1, "D6:E10");

            // 5 rows?
            Assert.IsTrue(cells.Count == 5);

            // loop rows in range 
            foreach (var row in cells)
            {
               string colA = row[0].Address;              

               XlsRange firstColumn = row.Where(c => c.ColumnName == "D").Single();               
               Assert.IsTrue(firstColumn.ColumnName == "D");               
            }
         }
      }

      private static string GetTestFileName()
      {
         string folder = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), @"..\..\.."));
         return Path.Combine(folder, "Tests.xlsx");
      }
   }
}
