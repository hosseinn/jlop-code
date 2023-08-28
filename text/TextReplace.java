
// TextReplace.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Search for the first occurrence of some words and/or
   replace some English spelled words with US spelled versions.
   Saved the changed document in "replaced.doc"

   Usage:
      run TextReplace story.doc
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.beans.*;
import com.sun.star.text.*;
import com.sun.star.util.*;


public class TextReplace
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run TextReplace <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

/*
    String words[] = {"(G|g)rit", "colou?r"};
    // String words[] = {"the", "questionable", "dubious", "silhouette"};
    findWords(doc, words);
*/

    String ukWords[] = {
            "colour", "neighbour", "centre", "behaviour", "metre", "through" };
    String usWords[] = {
            "color", "neighbor", "center", "behavior", "meter", "thru" };
    replaceWords(doc, ukWords, usWords);


    Lo.saveDoc(doc, "replaced.doc");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void findWords(XTextDocument doc, String[] words)
  {
    // get the text view cursor and link the page cursor to it
    XTextViewCursor tvc = Write.getViewCursor(doc);
    tvc.gotoStart(false);
    XPageCursor pageCursor = UnoRuntime.queryInterface(XPageCursor.class, tvc);
    // System.out.println("The current page number is " + pageCursor.getPage());

    try {
      XSearchable searchable = UnoRuntime.queryInterface(XSearchable.class, doc);
      XSearchDescriptor srchDesc = searchable.createSearchDescriptor();

      for(int i = 0; i < words.length; i++ ) {
        System.out.println("Searching for first occurrence of \"" + words[i] + "\"");
        srchDesc.setSearchString(words[i]);
        XPropertySet srchProps = UnoRuntime.queryInterface(XPropertySet.class, srchDesc);
        srchProps.setPropertyValue("SearchRegularExpression", true);
        // srchProps.setPropertyValue("SearchCaseSensitive", false);

        XInterface srch = (XInterface) searchable.findFirst(srchDesc);
        if (srch != null) {
          XTextRange matchTR = UnoRuntime.queryInterface(XTextRange.class, srch);

          tvc.gotoRange(matchTR, false); 
          System.out.println("  - found \"" + matchTR.getString() + "\"");
          System.out.println("    - on page " + pageCursor.getPage());
          //System.out.println("    - at cursor point " + Write.getCoordStr(tvc));

          tvc.gotoStart(true);  
          System.out.println("    - at char position: " + tvc.getString().length()); 
          // tvc.gotoRange(tvc.getEnd(), false); 
        }
        else
          System.out.println("  - not found");
      }
    }
    catch(com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of findWords()




  private static void replaceWords(XTextDocument doc, String[] oldWords, String[] newWords)
  {
    XReplaceable replaceable = UnoRuntime.queryInterface(XReplaceable.class, doc);
    XReplaceDescriptor replaceDesc = replaceable.createReplaceDescriptor();

    System.out.println("Change all occurrences of ...");
    for (int i = 0; i < oldWords.length; i++) {
      System.out.println("  " + oldWords[i] + " -> " + newWords[i]);
      replaceDesc.setSearchString(oldWords[i]);
      replaceDesc.setReplaceString(newWords[i]);
      int numChanges = replaceable.replaceAll(replaceDesc);     // Replace all words
      System.out.println("    - no. of changes: " + numChanges);
    }
  }  // end of replaceWords()


}  // end of TextReplace class
