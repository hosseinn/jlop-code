
// CopyTests.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Create a new sheet. Copy a data range.
    Report page breaks, and print some doc properties.

  Based on code in Dev Guide's SpreadSheetSample.java example.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.util.*;


public class CopyTests
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

    // Insert a new spreadsheet, and bring to front
    XSpreadsheet sheet = Calc.insertSheet(doc, "A new sheet", (short) 0x7FFF);
    Calc.setActiveSheet(doc, sheet);

    fillSheet(sheet);

    // print some document properties
    System.out.println("IsIterationEnabled: " +
                          Props.getProperty(doc, "IsIterationEnabled"));
    System.out.println("IterationCount: " +
                          Props.getProperty(doc, "IterationCount"));
    Date d = (Date) Props.getProperty(doc, "NullDate");
    System.out.println("NullDate: " + d.Day + "-" + d.Month + "-" + d.Year);


    Lo.saveDoc(doc, "copyTests.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void fillSheet(XSpreadsheet sheet)
  {
    // Copy a cell range
    Calc.highlightRange(sheet, "A1:B3", "Copy from");
    Calc.highlightRange(sheet, "D1:E3", "To");
    Calc.setVal(sheet, "A2", 123);
    Calc.setVal(sheet, "B2", 345);
    Calc.setVal(sheet, "A3", "=SUM(A2:B2)");
    Calc.setVal(sheet, "B3", "=FORMULA(A3)");

    XCellRangeMovement crMove = Lo.qi(XCellRangeMovement.class, sheet);
    CellRangeAddress sourceRange = Calc.getAddress(sheet, "A2:B3");
    CellAddress destCell = Calc.getCellAddress(sheet, "D2");
    crMove.copyRange(destCell, sourceRange);

    // Print automatic column page breaks
    XSheetPageBreak spBreak = Lo.qi(XSheetPageBreak.class, sheet);
    TablePageBreakData[] pageBreaks = spBreak.getColumnPageBreaks();
    System.out.print("Automatic column page breaks:");
    for (int i = 0; i < pageBreaks.length; i++) {
      if (!pageBreaks[i].ManualBreak)
        System.out.print(" " + pageBreaks[i].Position);
    }
    System.out.println("\n");
  }  // end of fillSheet()



}  // end of CopyTests class
