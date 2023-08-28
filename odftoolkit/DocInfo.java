
// DocInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan 2017

/*  Print document properties for the specified file.
    Usage:
      run DocInfo algs.odp
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
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    System.out.println();
    Props.showObjProps("Document", doc);
    Info.printDocProperties(doc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()

}  // end of DocInfo class
