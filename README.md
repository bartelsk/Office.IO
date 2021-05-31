# Office.IO
A utility library that makes it easier to work with Office files using C#. 

### Currently only features an Excel reader implementation.

## Usage

Read the value of cell **E9** from sheet **Customers**:

```csharp
using (ExcelReader xlsReader = new(myExcelFile))
{
  string value = xlsReader.ReadCell("Customers", "E9");  
}
```

Return all cell values from **range D6:E10** from **sheet 1**:

```csharp
using (ExcelReader xlsReader = new(myExcelFile))
{
  List<List<XlsRange>> cells = xlsReader.ReadRange(1, "D6:E10");

  // loop rows in range 
  foreach (var row in cells)
  {
     XlsRange firstColumnOfCurrentRow = row.Where(c => c.ColumnName == "D").Single(); 
     
     // cell address
     //firstColumnOfCurrentRow.Address
     
     // cell value
     //firstColumnOfCurrentRow.Value
  }
}
```

The **ReadRange** method returns an **XlsRange** type which looks like this:

```csharp
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
}

```

