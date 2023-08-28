
// Ranger.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Find out the filled cells in a disjoint range.

   Based on code in Dev Guide's SpreadSheetSample.java example.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class Ranger
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
    createTable(sheet);
    doCellRanges();

    Lo.saveDoc(doc, "ranger.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void createTable(XSpreadsheet sheet)
  {
    Calc.highlightRange(sheet, "A9:C30", "Example");
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
    Calc.setArray(sheet, "A10:C30", vals);
  }  // end of createTable()



  private static void doCellRanges()
  // examples using cell range collections
  {
    try {
      XSheetCellRangeContainer scrCon =
                Lo.createInstanceMSF(XSheetCellRangeContainer.class,
                                    "com.sun.star.sheet.SheetCellRanges");
      // collection of non-empty cells in the supplied cell ranges

      // supply four ranges
      insertRange(scrCon, 0, 0, 0, 0, false);  // A1:A1  -- empty
      insertRange(scrCon, 0, 1, 0, 2, true);   // A2:A3  -- empty
      insertRange(scrCon, 1, 0, 1, 2, false);  // B1:B3   -- empty

      insertRange(scrCon, 0, 9, 1, 29, false); // A10:B30  -- not empty

      // Query the list of filled cells
      System.out.print("All the filled cells: ");
      XEnumerationAccess cellsEA = scrCon.getCells();
      XEnumeration cellEnum = cellsEA.createEnumeration();
      while (cellEnum.hasMoreElements()) {
        XCellAddressable xAddr = Lo.qi(XCellAddressable.class, 
                                                  cellEnum.nextElement());
        Calc.printCellAddress( xAddr.getCellAddress() );
      }
      System.out.println();
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("doCellRanges() exception: " + e);  }
  }  // end of doCellRanges()



  private static void insertRange(XSheetCellRangeContainer scrCon,
                   int startCol, int startRow, int endCol, int endRow, boolean isMerged)
  // Inserts a cell range address into a cell range container and prints a message
  {
    CellRangeAddress addr = new CellRangeAddress();
    addr.Sheet = (short) 0;
    addr.StartColumn = startCol;
    addr.StartRow = startRow;
    addr.EndColumn = endCol;
    addr.EndRow = endRow;

    scrCon.addRangeAddress(addr, isMerged);

    System.out.println("Inserting " + Calc.getRangeStr(addr) + " " +
                   (isMerged ? "   with" : "without") + " merge," +
                   " resulting list: " + scrCon.getRangeAddressesAsString() + "\n");
  }  // end of insertRange()


}  // end of Ranger class
