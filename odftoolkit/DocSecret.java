
// DocSecret.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan 2017

/*  Add a user defined property to the specified file's
    document properities. It will be 
      Secret="<input msg>"

    Usage:
      run DocSecret algs.odp "Made in Thailand"
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
// import com.sun.star.text.*;
import com.sun.star.beans.*;


public class DocSecret
{

  public static void main(String args[])
  {
    if (args.length != 2) {
      System.out.println("Usage: run DocSecret <fnm> <secret msg>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    XPropertyContainer pCon = Info.getUserDefinedProps(doc);
    Props.setProperty(pCon, "Secret", args[1]);

    Info.printDocProperties(doc);

    Lo.save(doc);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of DocSecret class
