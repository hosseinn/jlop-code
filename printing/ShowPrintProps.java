
// ShowPrintProps.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* 
    usage:
      run ShowPrintProps skinner.png
      run ShowPrintProps storyStart.doc
      run ShowPrintProps classic_letter.ott
      run ShowPrintProps python-outline.pdf
      run ShowPrintProps algs.odp
      run ShowPrintProps "small totals.ods"
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;
import com.sun.star.beans.*;

import com.sun.star.awt.*;

import com.sun.star.text.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class ShowPrintProps
{

  public static void main(String args[])
  {  
    if (args.length != 1) {
      System.out.println("Usage: run ShowPrintProps fnm");
      return;
    }


    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }
    int docType = Info.reportDocType(doc);

    System.out.println();
    XPropertySet docProps = Print.getDocSettings(docType);
    Props.showProps("Document Settings", docProps);


    System.out.println();
    if (docType == Lo.WRITER) {
      XPagePrintable xpp = Lo.qi(XPagePrintable.class, doc);
      PropertyValue[] printProps = xpp.getPagePrintSettings();
      Props.showProps("Page print settings", printProps);
    }
    else if (docType == Lo.CALC) {
      XPropertySet globalSheetProps = 
                Lo.createInstanceMCF(XPropertySet.class,
                               "com.sun.star.sheet.GlobalSheetSettings");
      Props.showProps("Global Sheet Settings", globalSheetProps);

      XSpreadsheetDocument ssDoc = Calc.getSSDoc(doc);
      XSpreadsheet sheet = Calc.getSheet(ssDoc, 0);
      if (sheet == null) 
        System.out.println("Could not find a sheet");
      else {
        String styleName = (String) Props.getProperty(sheet, "PageStyle");
        System.out.println("\nPageStyle of first sheet: " + styleName);
        
        // get the properties set for the sheet's page style
        XPropertySet props = Info.getStyleProps(doc, "PageStyles", styleName);
        Props.showProps(styleName + " PageStyles", props);
          // see Calc Tests/AllStyleInfo.java

        System.out.println("\nPrinting areas of first sheet");
        XPrintAreas printAreas = Lo.qi(XPrintAreas.class, sheet);
        CellRangeAddress[] crAddrs = printAreas.getPrintAreas();
        Calc.printAddresses(crAddrs);
      }
    }



    Lo.closeDoc(args[0]);
    Lo.closeOffice();
  }  // end of main()


}  // end of ShowPrintProps class
