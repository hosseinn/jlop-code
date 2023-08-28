
// DocInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/*  Pring document properties and document info

    Print all the fonts supported by Office.

    Usage:
      run DocInfo build.odt
      run DocInfo story.doc     // first manually set Author in its properties
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class DocInfo
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run DocInfo <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    Props.showObjProps("Document", doc);

    Info.printDocProperties(doc);

    // System.out.println("Available Fonts:");
    // Lo.printNames( Info.getFontNames() );

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of DocInfo class
