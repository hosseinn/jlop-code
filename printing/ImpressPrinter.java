
// ImpressPrinter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/*  Print slides as as handouts (6 slides / sheet), in
    black & white

    The handout setting is ignored, and the file is printed
    as normal, one slide/sheet, in color

    Usage:
      run ImpressPrinter algs.odp
      run ImpressPrinter algs.odp phantom

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

import com.sun.star.text.*;
import com.sun.star.drawing.*;


public class ImpressPrinter
{

  public static void main(String args[])
  {  
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: ImpressPrinter fnm  [(partial)printer-name]");
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
    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    int docType = Info.reportDocType(doc);
    if (docType != Lo.IMPRESS) {
      System.out.println("Not a slides document");
      Lo.closeDoc(doc);
      Lo.closeOffice();
      return;
    }


    XPropertySet props = Lo.createInstanceMSF(XPropertySet.class,
                             "com.sun.star.presentation.DocumentSettings");
    Props.setProperties(props,
          new String[] { "IsPrintHandout", "SlidesPerHandout", "IsPrintFitPage",
                         "IsPrintDate", "PrintQuality", "PrinterName" },
          new Object[] { true, (short)6, true, true, 2, pName}   // 2 == B&W
      );     // 6, ?? 6, 8, 4, 4

    Props.showProps("Document Settings", props);

    // fails: http://openoffice.2283327.n4.nabble.com/api-dev-simpress-printing-handouts-programmatically-howto-td2774228.html
    // XDrawPage dp = Draw.getHandoutMasterPage(doc);
    // Props.showObjProps("Handout Master Page", dp);

    // GUI.setVisible(doc, true);
    // Lo.dispatchCmd("Print");
    // Lo.delay(500);

    // set printer props
    XPrintable xp = Lo.qi(XPrintable.class, doc);
    Print.usePrinter(xp, pName);
    xp.setPrinter( Props.makeProps("PaperOrientation", PaperOrientation.LANDSCAPE,
                                   "PaperFormat", PaperFormat.A4) );
                                         // must set paper format again
    Print.reportPrinterProps(xp);

    Print.print(xp);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of ImpressPrinter class
