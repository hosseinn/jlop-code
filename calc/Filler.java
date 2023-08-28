
// Filler.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Various forms of auto-filling:
      dates, geometric, linear, 
      left, up, right

    Based on code in Dev Guide's SpreadSheetSample.java example.
    See
      https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Fill_Series
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class Filler
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
    fillSeries(sheet);
    // Calc.highlightRange(sheet, "A9:G10", "Bill Series Examples");

    Lo.saveDoc(doc, "filler.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void fillSeries(XSpreadsheet sheet)
  // six series filling examples
  {
    XCellSeries series;

    // set first two values of three rows
    Calc.setVal(sheet, "A7", 1);    // 1. ascending
    Calc.setVal(sheet, "B7", 2);

    Calc.setDate(sheet, "A8", 28, 2, 2015);   //2. dates, descending
    Calc.setDate(sheet, "B8", 28, 1, 2015);

    Calc.setVal(sheet, "A9", 6);     // 3. descending
    Calc.setVal(sheet, "B9", 4);

    // Autofill, using first 2 cells to right to determine progressions
    series = Calc.getCellSeries(sheet, "A7:G9");
    series.fillAuto(FillDirection.TO_RIGHT, 2);

    // ----------------------------------------
    Calc.setVal(sheet, "A2", 1);
    Calc.setVal(sheet, "A3", 4);

    // Fill 2 rows; 2nd row is not filled completely since
    // the end value is reached
    series = Calc.getCellSeries(sheet, "A2:E3");
    series.fillSeries(FillDirection.TO_RIGHT, FillMode.LINEAR,
                                                Calc.NO_DATE, 2, 9);   
                      // ignore date mode; step == 2; end at 9


    // ----------------------------------------
    Calc.setDate(sheet, "A4", 20, 11, 2015);    // day, month, year

    // fill by adding one month to date; day is unchanged
    series = Calc.getCellSeries(sheet, "A4:E4");
    series.fillSeries(FillDirection.TO_RIGHT, FillMode.DATE,
                       FillDateMode.FILL_DATE_MONTH, 1, Calc.MAX_VALUE);
//   series.fillAuto(FillDirection.TO_RIGHT, 1);  // increments day not month


    // ----------------------------------------
    Calc.setVal(sheet, "E5", "Text 10");   // start in the middle of a row

    // Fill from right to left with text+value in steps of +10
    series = Calc.getCellSeries(sheet, "A5:E5");
    series.fillSeries(FillDirection.TO_LEFT, FillMode.LINEAR,
                       Calc.NO_DATE, 10, Calc.MAX_VALUE);


    // ----------------------------------------
    Calc.setVal(sheet, "A6", "Jan");

    // Fill with values generated automatically from first entry
    series = Calc.getCellSeries(sheet, "A6:E6");

    series.fillSeries(FillDirection.TO_RIGHT, FillMode.AUTO,
                                  Calc.NO_DATE, 1, Calc.MAX_VALUE);
    // series.fillAuto(FillDirection.TO_RIGHT, 1);  // does the same


    // ----------------------------------------
    Calc.setVal(sheet, "G6", 10); 

    // Fill from  bottom to top with a geometric series (*2)
    series = Calc.getCellSeries(sheet, "G2:G6");
    series.fillSeries(FillDirection.TO_TOP, FillMode.GROWTH,
                       Calc.NO_DATE, 2, Calc.MAX_VALUE);


  }  // end of fillSeries()


}  // end of Filler class
