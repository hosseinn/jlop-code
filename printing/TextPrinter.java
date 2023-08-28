
// TextPrinter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/*  Print a file 2 pages to a sheet on a specified printer.

    The page/sheet setting is ignored, and the file is printed
    as normal, one page / sheet.

    Usage:
      run TextPrinter storyStart.doc FinePrint
      run TextPrinter storyStart.doc

      run TextPrinter python-outline.pdf 1200

      run TextPrinter classic_letter.ott
      run TextPrinter classic_letter.ott phantom


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


public class TextPrinter
{

  public static void main(String args[])
  {  
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: TextPrinter fnm  [(partial)printer-name]");
      return;
    }

    String fnm = args[0];
    int pagesPerSheet = 2;

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
    if (docType != Lo.WRITER) {
      System.out.println("Not a text document");
      Lo.closeDoc(doc);
      Lo.closeOffice();
      return;
    }

    XPrintable xp = Lo.qi(XPrintable.class, doc);
    Print.usePrinter(xp, pName);
    Print.reportPrinterProps(xp);

    XPagePrintable xpp = Lo.qi(XPagePrintable.class, doc);
    PropertyValue[] props = xpp.getPagePrintSettings();

    Props.setProp(props, "IsLandscape", true);
    Props.setProp(props, "PageColumns", (short)2);

    Props.showProps("Page print settings", props);


  // fails: see http://en.libreofficeforum.org/node/12145 (Dec 2015)
  // http://comments.gmane.org/gmane.comp.documentfoundation.libreoffice.bugs/332836

   xpp.setPagePrintSettings(props);
   xpp.printPages(new PropertyValue[1]);    

  // two approaches: http://openoffice3.web.fc2.com/OOoBasic_Writer_No2.html
  //  xpp.printPages(props);    


    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



}  // end of TextPrinter class
