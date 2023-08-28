
// ItalicsStyler.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Search for all occurrences of a specified string (case-insensitive)
   and set their style to be in red italics. Save changed
   text to "italicized.doc".

   Usage:
      run ItalicsStyler story.doc scandal
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.container.*;

import com.sun.star.beans.*;
import com.sun.star.text.*;
import com.sun.star.util.*;


public class ItalicsStyler
{

  public static void main(String args[])
  {
    if (args.length < 2) {
      System.out.println("Usage: run ItalicsStyler <fnm> <phrase>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument textDoc = Write.openDoc(args[0], loader);
    if (textDoc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    italicizeAll(textDoc, args[1]);

    Lo.saveDoc(textDoc, "italicized.doc");
    Lo.closeDoc(textDoc);
    Lo.closeOffice();
  }  // end of main()



  private static void italicizeAll(XTextDocument doc, String phrase)
  {
    // get the text view cursor and link the page cursor to it
    XTextViewCursor tvc = Write.getViewCursor(doc);
    tvc.gotoStart(false);
    XPageCursor pageCursor = UnoRuntime.queryInterface(XPageCursor.class, tvc);
    // System.out.println("The current page number is " + pageCursor.getPage());

    try {
      XSearchable xSearchable = UnoRuntime.queryInterface(XSearchable.class, doc);
      XSearchDescriptor srchDesc = xSearchable.createSearchDescriptor();

      System.out.println("Searching for all occurrence of \"" + phrase + "\"");
      int phraseLen = phrase.length();
      srchDesc.setSearchString(phrase);
      Props.setProperty(srchDesc, "SearchCaseSensitive", false);

      XIndexAccess matches = xSearchable.findAll(srchDesc);
      System.out.println("No. of matches: " + matches.getCount());
      for (int i = 0; i < matches.getCount(); i++) {
        XTextRange matchTR = UnoRuntime.queryInterface(XTextRange.class, matches.getByIndex(i));
        if (matchTR != null) {
          tvc.gotoRange(matchTR, false);
          System.out.println("  - found \"" + matchTR.getString() + "\"");
          System.out.println("    - on page " + pageCursor.getPage());
          // System.out.println("    - at cursor point " + Write.getCoordStr(tvc));
          tvc.gotoStart(true);
          System.out.println("    - starting at char position: " + 
                                        (tvc.getString().length() - phraseLen));
          Props.setProperties(matchTR, 
                         new String[] {"CharColor", "CharPosture"},
                         new Object[] { 0xFF0000, com.sun.star.awt.FontSlant.ITALIC} );
        }
      }
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of italicizeAll()


}  // end of ItalicsStyler class
