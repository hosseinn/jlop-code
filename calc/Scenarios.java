
// Scenarios.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Shows three alternative scenarios.
    Starts with the second.

  Based on code in Dev Guide's SpreadSheetSample.java example.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class Scenarios
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

    // create three scenarios
    Object[][] vals = new Object[][] {{11,12},{"Test13","Test14"}};
    Calc.insertScenario(sheet, "B10:C11", vals, "First Scenario", "1st scenario.");

    vals[0][0] = "Test21";
    vals[0][1] = "Test22";
    vals[1][0] = 23;
    vals[1][1] = 24;
    Calc.insertScenario(sheet, "B10:C11", vals, "Second Scenario", "Visible scenario.");

    vals[0][0] = 31;
    vals[0][1] = 32;
    vals[1][0] = "Test33";
    vals[1][1] = "Test34";
    Calc.insertScenario(sheet, "B10:C11", vals, "Third Scenario", "Last scenario.");


    Calc.applyScenario(sheet, "Second Scenario");

    Lo.saveDoc(doc, "scenarios.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of Scenarios class
