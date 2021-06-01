using BartelsOnline.Office.IO.Excel;
using BartelsOnline.Office.IO.Excel.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
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
            string result = xlsReader.ReadCell("Sheet1", "L4");
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
            List<List<XlsRange>> rows = xlsReader.ReadRange("Sheet1", "H9:I10");
                        
            // 2 rows?
            Assert.IsTrue(rows.Count == 2);

            // loop rows in range
            foreach (var row in rows)
            {
               XlsRange firstColumn = row.Where(c => c.ColumnName == "H").Single();
               XlsRange secondColumn = row.Where(c => c.ColumnName == "I").Single();
               Assert.IsTrue(firstColumn.ColumnName == "H");
               Assert.IsTrue(secondColumn.ColumnName == "I");               
            }  
         }
      }

      [TestMethod]
      public void ReadRangeValuesFromSheetNumber()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange(1, "A1:B4");

            // 4 rows?
            Assert.IsTrue(rows.Count == 4);

            // loop rows in range 
            foreach (var row in rows)
            {
               string colA = row[0].Address;              

               XlsRange firstColumn = row.Where(c => c.ColumnName == "A").Single();               
               Assert.IsTrue(firstColumn.ColumnName == "A");               
            }
         }
      }

      [TestMethod]
      public void ReadAllRowsInColumnRangeFromSheetName()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange("Sheet1", "F:G");

            // loop rows in range 
            foreach (var row in rows)
            {
               string colA = row[0].Address;

               XlsRange firstColumn = row.Where(c => c.ColumnName == "F").Single();
               Assert.IsTrue(firstColumn.ColumnName == "F");
            }
         }
      }

      [TestMethod]
      public void ReadAllRowsInColumnRangeFromSheetNumber()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange(1, "K:L");

            // loop rows in range 
            foreach (var row in rows)
            {
               string colA = row[0].Address;

               XlsRange firstColumn = row.Where(c => c.ColumnName == "K").Single();
               Assert.IsTrue(firstColumn.ColumnName == "K");
            }
         }
      }

      [TestMethod]
      [ExpectedException(typeof(ArgumentException))]
      public void InvalidRangeAddress1()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange(1, "A3:M");
         }
      }

      [TestMethod]
      [ExpectedException(typeof(ArgumentException))]
      public void InvalidRangeAddress2()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange(1, "B:I23");
         }
      }

      [TestMethod]
      [ExpectedException(typeof(ArgumentException))]
      public void InvalidRangeAddress3()
      {
         string fileName = GetTestFileName();
         using (ExcelReader xlsReader = new(fileName))
         {
            List<List<XlsRange>> rows = xlsReader.ReadRange(1, "D4");
         }
      }

      private static string GetTestFileName()
      {
         string folder = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), @"..\..\.."));
         return Path.Combine(folder, "Tests.xlsx");
      }
   }
}
