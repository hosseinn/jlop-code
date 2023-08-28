
// DocPrinter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* 
    usage:
      run DocPrinter skinner.png FinePrint
      run DocPrinter skinner.png "1200"

      run DocPrinter storyStart.doc FinePrint
      run DocPrinter chapter.ott phantom 3-4

      run DocPrinter python-outline.pdf phantom 2-3   

      run DocPrinter Books.odb

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

import com.sun.star.awt.*;


public class DocPrinter
{

  public static void main(String args[])
  {  
    if ((args.length < 1) || (args.length > 3)) {
      System.out.println("Usage: DocPrinter fnm [(partial)printer-name [no-of-pages]]");
      return;
    }

    String fnm = args[0];

    String pName = JPrint.getDefaultPrinterName();  // default

    if (args.length > 1) {   // 2 or 3 args
      String[] printerNames = JPrint.findPrinterNames(args[1]);
      if (printerNames == null)
        System.out.println("Using default printer: \"" + pName + "\"");
      else {
        pName = printerNames[0];
        System.out.println("Using first matching printer: \"" + pName + "\"");
      }
    }

    String pagesStr = "1-";   // default is print all pages
    if (args.length == 3) 
      pagesStr = args[2];


    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    int docType = Info.reportDocType(doc);

    // XPrintableBroadcaster pb = Lo.qi(XPrintableBroadcaster.class, doc);
    // Print.setDocListener(pb);    // pb is null

    XPrintable xp = Lo.qi(XPrintable.class, doc);
    if (xp == null)
      System.out.println("Cannot print; XPrintable is null");
    else if (!Print.isPrintable(docType))
      System.out.println("Cannot print that document type");
    else {
      Print.usePrinter(xp, pName);
      Print.reportPrinterProps(xp);
      Print.print(xp, pagesStr);
    }
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of DocPrinter class
