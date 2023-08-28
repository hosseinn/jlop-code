
// CopyPasteText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Copy/paste a specified sentence form the file using
   the Clip functions and using "Copy"/"Paste" dispatches.

   "Copy" stores the data in multiple formats on the clipboard
    which are read by the Clip.getXXX() methods.

   Based on HighlightText.java in Text Tests/
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;
import com.sun.star.view.*;


public class CopyPasteText
{

  private static final String FNM = "storyStart.doc";


  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(FNM, loader);
    if (doc == null) {
      System.out.println("Could not open " + FNM);
      Lo.closeOffice();
      return;
    }

    // useClipUtils(doc, 5);
    useDispatches(doc, 4);


    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void useClipUtils(XTextDocument doc, int n)
  {
    XSentenceCursor senCursor = Write.getSentenceCursor(doc);
    senCursor.gotoStart(false);     // start of text; no selection

    gotoSentence(senCursor, n);

    // copy to clipboard
    Clip.setText(senCursor.getString());
    System.out.println("Copied \"" + senCursor.getString() + "\"");

    Lo.wait(2000);
    XTextCursor cursor = Write.getCursor(doc);
    cursor.gotoEnd(false);   // go to end of doc

    // paste into doc
    Write.append(cursor, Clip.getText());

    GUI.setVisible(doc, true);   // so we can see change
    Lo.waitEnter();
  }  // end of useClipUtils()





  private static void useDispatches(XTextDocument doc, int n)
  {
    GUI.setVisible(doc, true);   // *must* be made visible
    Lo.wait(2000);   // give Office time to appear

    XTextViewCursor tvc = Write.getViewCursor(doc);
    XSentenceCursor senCursor = Write.getSentenceCursor(doc);
    senCursor.gotoStart(false);     // go to start of text; no selection

    gotoSentence(senCursor, n);

    // move the text view cursor to highlight the current paragraph
    tvc.gotoRange(senCursor.getStart(), false);
    tvc.gotoRange(senCursor.getEnd(), true);

    Lo.dispatchCmd("Copy");

    Clip.listFlavors();

    // save copied text as RTF & HTML
    FileIO.saveString("storyFrag.rtf", Clip.getRTF());
    FileIO.saveString("storyFrag.html", Clip.getHTML());

    // save embedded text fragment as ODT
    FileIO.saveBytes("storyFrag.odt", Clip.getEmbedSource());

    Lo.wait(2000);
    tvc.gotoEnd(false);    // go to end of doc
    Lo.dispatchCmd("Paste");

    Lo.waitEnter();
  }  // end of useDispatches()



  private static void gotoSentence(XSentenceCursor senCursor, int n)
  {
    String currSen;
    do {
      senCursor.gotoEndOfSentence(true);   // select all of sentence
      currSen = senCursor.getString();
      if (currSen.length() > 0) {
        // System.out.println("P " + n + " :<" + currSen + ">");
        n--;
      }
    } while ((n > 0) && senCursor.gotoNextSentence(false));
  }  // end of gotoSentence()



}  // end of CopyPasteText class
