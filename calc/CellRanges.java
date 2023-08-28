
// CellRanges.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Illustrates:
       - range data block
       - array formula
       - cell cursor
       - used area

    Based on code in Dev Guide's SpreadSheetSample.java example.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class CellRanges
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    doCellRange(sheet);
    doCellCursor(sheet);

    Lo.saveDoc(doc, "cellRanges.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void doCellRange(XSpreadsheet sheet)
  // examples using the SheetCellRange service
  {
    try {
      Calc.highlightRange(sheet, "A2:C23", "Range Data Example");

      Object[][] vals = {
                  { "Name", "Fruit", "Quantity" }, { "Alice", "Apples", 3 },
                  { "Alice", "Oranges", 7 }, { "Bob", "Apples", 3 },
                  { "Alice", "Apples", 9 }, { "Bob", "Apples", 5 },
                  { "Bob", "Oranges", 6 }, { "Alice", "Oranges", 3 },
                  { "Alice", "Apples", 8 }, { "Alice", "Oranges", 1 },
                  { "Bob", "Oranges", 2 }, { "Bob", "Oranges", 7 },
                  { "Bob", "Apples", 1 }, { "Alice", "Apples", 8 },
                  { "Alice", "Oranges", 8 }, { "Alice", "Apples", 7 },
                  { "Bob", "Apples", 1 }, { "Bob", "Oranges", 9 },
                  { "Bob", "Oranges", 3 }, { "Alice", "Oranges", 4 },
                  { "Alice", "Apples", 9 }
                };
      Calc.setArray(sheet, "A3:C23", vals);    // or just "A1"


      XCellRange cellRange = Calc.getCellRange(sheet, "A3:C23");
      XCellRangeData crData = Lo.qi(XCellRangeData.class, cellRange);
      Calc.printAddress(cellRange);

      // Sheet operation using the range of the crData
      XSheetOperation sheetOp = Lo.qi(XSheetOperation.class, crData);
      double avg = sheetOp.computeFunction(GeneralFunction.AVERAGE);
      System.out.println("Average of A3:C23: " + avg + "\n");


      // ----- Array formulas
      XCellRange range3Rows = Calc.getCellRange(sheet, "E3:G5");
      Calc.highlightRange(sheet, "E2:G5", " Array Formula Example");
      XArrayFormulaRange afRange = Lo.qi(XArrayFormulaRange.class, range3Rows);
      // Insert a 3x3 unit matrix
      afRange.setArrayFormula("=A3:C5");
      System.out.println("Array formula: " + afRange.getArrayFormula() + "\n");


      // ------ Cell Ranges Query
      XCellRangesQuery crQuery = Lo.qi(XCellRangesQuery.class, cellRange);
      XSheetCellRanges cellRanges = 
                crQuery.queryContentCells((short) CellFlags.STRING);
      System.out.println("Text Cells in A3:C23: " +
                        cellRanges.getRangeAddressesAsString() + "\n");
    }
    catch(Exception e)
    {  System.out.println("doCellRange() exception: " + e);  }
  }  // end of doCellRange()




  private static void doCellCursor(XSpreadsheet sheet)
  {
    // Find the array formula using a cell cursor
    XCellRange xRange = Calc.getCellRange(sheet, "E4");
    XSheetCellRange cellRange = Lo.qi(XSheetCellRange.class, xRange);
    XSheetCellCursor cursor = sheet.createCursorByRange(cellRange);
    cursor.collapseToCurrentArray();

    XArrayFormulaRange xArray = Lo.qi(XArrayFormulaRange.class, cursor);
    System.out.println("Array formula in " +
                       Calc.getRangeStr(cellRange) +
                       ": \"" + xArray.getArrayFormula() + "\"");

    // Find the used area
    XUsedAreaCursor uaCursor = Lo.qi(XUsedAreaCursor.class, cursor);
    uaCursor.gotoStartOfUsedArea(false);
    uaCursor.gotoEndOfUsedArea(true);

    // uaCursor and cursor are interfaces of the same object -
    // so modifying uaCursor takes effect on cursor and its cellRange:
    System.out.println("The used area is: " +
                           Calc.getRangeStr(cellRange) + "\n");
  }   // end of doCellCursor()


}  // end of CellRanges class
