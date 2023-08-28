
// Multiples.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Two multiples repeated ops examples, used to fill
    table from starting values.

    Based on code in Dev Guide's SpreadSheetSample.java example.
    See
       https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Multiple_Operations
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class Multiples
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
    doMultiples(sheet);

    Lo.saveDoc(doc, "multiples.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void doMultiples(XSpreadsheet sheet)
  // examples using the SheetCellRange service
  {
    // fill top-left corner
    Calc.setVal(sheet, "A4", "=A5^B4");
    Calc.setVal(sheet, "A5", 1);
    Calc.setVal(sheet, "B4", 1);

    // fill down left edge and to the right along top
    Calc.getCellSeries(sheet, "A5:A9").fillAuto(FillDirection.TO_BOTTOM, 1);
    Calc.getCellSeries(sheet, "B4:F4").fillAuto(FillDirection.TO_RIGHT, 1);

    CellRangeAddress formulaRange = Calc.getAddress(sheet, "A4");
                     // specify where the formula is located

    CellAddress collCell = Calc.getCellAddress(sheet, "A5");  // starting column cell
    CellAddress rowCell = Calc.getCellAddress(sheet, "B4");  // starting row cell

    XCellRange cellRange = Calc.getCellRange(sheet, "A4:F9");  // multiple ops data

    XMultipleOperation multOp = Lo.qi(XMultipleOperation.class, cellRange);
    multOp.setTableOperation(formulaRange, 
                       TableOperationMode.BOTH, collCell, rowCell);
                        // fill in the table along both axes

    // a row of trig functions  ------------------
    Calc.setVal(sheet, "B11", "=SIN(A11)");
    Calc.setVal(sheet, "C11", "=COS(A11)");
    Calc.setVal(sheet, "D11", "=TAN(A11)");

    // initial values on the left, going down
    Calc.setVal(sheet, "A12", 0);
    Calc.setVal(sheet, "A13", 0.2);
    // finish off by filling down
    Calc.getCellSeries(sheet, "A12:A16").fillAuto(FillDirection.TO_BOTTOM, 2);

    formulaRange = Calc.getAddress(sheet, "B11:D11");
                     // specify where the formulas are located

    collCell = Calc.getCellAddress(sheet, "A11");  // rowCell not needed

    cellRange = Calc.getCellRange(sheet, "A12:D16");

    multOp = Lo.qi(XMultipleOperation.class, cellRange);
    multOp.setTableOperation(formulaRange, 
                 TableOperationMode.COLUMN, collCell, rowCell);

    Calc.highlightRange(sheet, "A3:F16", " Two Multiple Ops Examples");
  }  // end of doMultiples()


}  // end of Multiples class
