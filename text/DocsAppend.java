
// DocsAppend.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/*  Append text documents together, creating <fnm>APPENDED.<ext>
    single text document. <fnm> and <ext> are from the first file.

    Usage:
      run DocsAppend oneLiner.odt story.doc
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;

import com.sun.star.document.*;
import com.sun.star.beans.*;


public class DocsAppend
{

  public static void main(String args[])
  {
    if (args.length < 2) {
      System.out.println("Usage: DocsAppend fnm1 fnm2 ...");
      return;
    }
    else
      System.out.println("Appending " + args.length + " files");

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      Lo.closeOffice();
      return;
    }

    appendTextFiles(doc, args);

    Lo.saveDoc(doc, Info.getName(args[0]) + "_APPENDED." + Info.getExt(args[0]));
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static void appendTextFiles(XTextDocument doc, String[] args)
  {
    XTextCursor cursor = Write.getCursor(doc);
    for (int i=1; i < args.length; i++) {  
      // start at 1 to skip the first file, which is doc
      try {
        cursor.gotoEnd(false);
        // Write.pageBreak(cursor);

        System.out.println("Appending " + args[i]);
        XDocumentInsertable inserter = 
                          UnoRuntime.queryInterface(XDocumentInsertable.class, cursor);
                          // XDocumentInsertable only works with text files
        if (inserter == null)
          System.out.println("Document inserter could not be created");
        else
          inserter.insertDocumentFromURL( FileIO.fnmToURL(args[i]), new PropertyValue[0]);
      }
      catch (java.lang.Exception e)
      {  System.out.println("Could not append " + args[i] + ": " + e); }
    }
  }  // end of appendTextFiles()


}  // end of DocsAppend class
