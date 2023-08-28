
// SheetPrinter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Print the first sheet of a spreadsheet scaled so that
   it uses two pieces of paper along the y-axis.

   There is some commented out code that sets the
   printing area to just the "E" column.

   Change the sheet header.

   Usage:
      run SheetPrinter totals.ods
      run SheetPrinter totals.ods phantom

    Get printer names by using
      java ListPrinters
    or
      printersList.bat
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class SheetPrinter
{

  public static void main(String args[])
  {
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: SheetPrinter fnm  [(partial)printer-name]");
      return;
    }

    String fnm = args[0];

    String pName = JPrint.getDefaultPrinterName();  // default
    if (args.length == 2) { 
      String[] printerNames = JPrint.findPrinterNames(args[1]);
      if (printerNames != null) {
        System.out.println("Using first matching printer: \"" + printerNames[0] + "\"");
        pName = printerNames[0];
      }
      else
        System.out.println("Using default printer: \"" + pName + "\"");
    }


    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    int docType = Info.reportDocType(doc);
    if (docType != Lo.CALC) {
      System.out.println("Not a spreadsheet document");
      Lo.closeDoc(doc);
      Lo.closeOffice();
      return;
    }

    XSpreadsheet sheet = Calc.getSheet(doc, 0);
    String styleName = (String) Props.getProperty(sheet, "PageStyle");
    System.out.println("PageStyle of first sheet: \"" + styleName + "\"");

    // get the properties set for the sheet's page style
    XPropertySet props = Info.getStyleProps(doc, "PageStyles", styleName);
    Props.setProperty(props, "ScaleToPagesY", (short)2);
                 // use a max of 2 pages along the y-axis
    Props.showProps(styleName, props);

    showTotalsHeader(props);

/*
    // print only the "E" column
    XPrintAreas printAreas = Lo.qi(XPrintAreas.class, sheet);
    printAreas.setPrintAreas(new CellRangeAddress[] {});   // reset print areas

    CellRangeAddress addr = Calc.getAddress(sheet, "E1:E111");
    printAreas.setPrintAreas(new CellRangeAddress[]{ addr });   // set area
*/

    // set printer settings
    XPrintable xp = Lo.qi(XPrintable.class, doc);
    Print.usePrinter(xp, pName);
    PropertyValue[] printProps =
        Props.makeProps("PaperOrientation", PaperOrientation.LANDSCAPE,
                        "PageFormat", PaperFormat.A4);
    xp.setPrinter(printProps);
    Print.reportPrinterProps(xp);

    Print.print(xp);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void showTotalsHeader(XPropertySet props)
  // change the header of the sheet
  {
    //  get the right-hand header and footer
    XHeaderFooterContent header = Calc.getHeadFoot(props, "RightPageHeaderContent");
    XHeaderFooterContent footer = Calc.getHeadFoot(props, "RightPageFooterContent");

    // print details about them
    Calc.printHeadFoot("Right Header", header);
    Calc.printHeadFoot("Right Footer", footer);
    System.out.println();

    // modify the header center text to be "Totals"
    Calc.setHeadFoot(header, Calc.HF_CENTER, "Totals");

    // turn on headers and make left and right page headers the same
    Props.setProperty(props, "HeaderIsOn", true);
    Props.setProperty(props, "HeaderIsShared", true);
                              // from style.PageProperties 
    Props.setProperty(props, "RightPageHeaderContent", header);
  }  // end of showTotalsHeader()


}  // end of SheetPrinter class
