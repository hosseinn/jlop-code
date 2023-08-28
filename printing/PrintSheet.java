
// PrintSheet.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Open the tables.ods spreadsheet,
   which has 3 sheets called Sheet1, Sheet2, and Sheet3.

   Print Sheet2 using the "FinePrint" printer
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

// import com.sun.star.util.*;



public class PrintSheet
{

  public static void main(String args[])
  {  
    String fnm = "tables.ods";   // hardwired
    String pName = "FinePrint";

    // load the spreadsheet
    XComponentLoader loader = Lo.loadOffice();
    XComponent cDoc = Lo.openReadOnlyDoc(fnm, loader);
    XSpreadsheetDocument doc = Calc.getSSDoc(cDoc);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    // what are the sheets called?
    String[] sheetNms = Calc.getSheetNames(doc);
    System.out.println("Names of sheets (" + sheetNms.length + "):");
    for(String sheetNm : sheetNms)
      System.out.println("  " + sheetNm);

    // make "Sheet2" active
    XSpreadsheet sheet = Calc.getSheet(doc, "Sheet2");
    Calc.setActiveSheet(doc, sheet);

    // set Global Sheet settings
    // changes are remembered
    XPropertySet gsProps = 
                Lo.createInstanceMCF(XPropertySet.class,
                     "com.sun.star.sheet.GlobalSheetSettings");
    Props.setProperty(gsProps, "PrintAllSheets", false);
    System.out.println();
    Props.showProps("Global Sheet Settings", gsProps);

    // set printer settings
    XPrintable xp = Lo.qi(XPrintable.class, doc);
    Print.usePrinter(xp, pName);
    PropertyValue[] printProps =
        Props.makeProps("PaperOrientation", PaperOrientation.LANDSCAPE,
                        "PageFormat", PaperFormat.A4);
    xp.setPrinter(printProps);
    // Print.reportPrinterProps(xp);

    Print.print(xp);

    // reset global settings property
    Props.setProperty(gsProps, "PrintAllSheets", true);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of PrintSheet class
