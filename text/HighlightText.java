
// HighlightText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Highlight the text in various ways

   Usage:
     > compile HighlightText.java

     > run HighlightText storyStart.doc
     > run HighlightText scandalStart.txt
     > run HighlightText story.doc

  If you interrupt this program, make sure that LibreOffice is 
  killed. Use lolist.bat and lokill.bat.

*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;
import com.sun.star.view.*;


public class HighlightText
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run HighlightText <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    showParagraphs(doc);
    System.out.println("Word count: " + countWords(doc));
    showLines(doc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void showParagraphs(XTextDocument doc)
  {
    XTextViewCursor tvc = Write.getViewCursor(doc);
    XParagraphCursor paraCursor = Write.getParagraphCursor(doc);
    paraCursor.gotoStart(false);     // go to start of text; no selection

    String currPara;
    do {
      paraCursor.gotoEndOfParagraph(true);   // select all of paragraph
      currPara = paraCursor.getString();
      if (currPara.length() > 0) {
        // move the text view cursor to highlight the current paragraph
        tvc.gotoRange(paraCursor.getStart(), false);
        tvc.gotoRange(paraCursor.getEnd(), true);

        System.out.println("P<" + currPara + ">");
        Lo.wait(500);   // to slow down the paragraph changing speed
      }
    } while (paraCursor.gotoNextParagraph(false));
  }  // end of showParagraphs()





  private static int countWords(XTextDocument doc)
  {
    XWordCursor wordCursor = Write.getWordCursor(doc);
    wordCursor.gotoStart(false);     // go to start of text

    int wordCount = 0;
    String currWord;
    do {
      wordCursor.gotoEndOfWord(true);
      currWord = wordCursor.getString();
      if (currWord.length() > 0) {
        // System.out.println("<" + currWord + ">");
        wordCount++;
      }
    } while( wordCursor.gotoNextWord(false));
    return wordCount;
  }  // end of countWords()



  private static void showLines(XTextDocument doc)
  {
    XTextViewCursor tvc = Write.getViewCursor(doc);
    tvc.gotoStart(false);     // go to start of text

    XLineCursor lineCursor = UnoRuntime.queryInterface(XLineCursor.class, tvc); 

    boolean haveText = true;
    do {
      lineCursor.gotoStartOfLine(false);
		  lineCursor.gotoEndOfLine(true); 
      System.out.println("L<" + tvc.getString() + ">");    // no text selection in line cursor
      Lo.wait(500);   // to slow down the line changing speed
      tvc.collapseToEnd();
      haveText = tvc.goRight((short) 1, true); 
    } while (haveText);
  }  // end of showLines()


}  // end of HighlightText class
