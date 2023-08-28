
// TalkingBook.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Speak the text in the loaded file, a word at a time

   Speech library: FreeTTS (http://freetts.sourceforge.net/) and
   its MBROLA voice.

   Use compileTalk.bat and runTalk.bat to compile and run
   this example:
     > compileTalk TalkingBook.java

     > runTalk TalkingBook storyStart.doc
               // first few paragraphs of story -- short

     > runTalk TalkingBook scandalStart.txt
              // the speaking breaks at new lines with text files.

     > runTalk TalkingBook story.doc
              // complete story -- long!

  If you interrupt this program, make sure that LibreOffice is 
  killed. Use lolist.bat and lokill.bat to check.

*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class TalkingBook
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: runTalk TalkingBook <fnm>");
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

    speakSentences(doc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void speakSentences(XTextDocument doc)
  {
    Speaker speaker = new Speaker();

    XTextViewCursor tvc = Write.getViewCursor(doc);

    XParagraphCursor paraCursor = Write.getParagraphCursor(doc);
    paraCursor.gotoStart(false);     // go to start of text

    // create range comparer for the entire document
    XTextRangeCompare comparer = 
             UnoRuntime.queryInterface(XTextRangeCompare.class, doc.getText());

    String currParaStr, currSentStr;
    do {
      paraCursor.gotoEndOfParagraph(true);   // select all of paragraph
      XTextRange endPara = paraCursor.getEnd();

      currParaStr = paraCursor.getString();
      System.out.println("P<" + currParaStr + ">");

      if (currParaStr.length() > 0) {

        // set sentence cursor pointing to start of this paragraph
        XTextCursor cursor = paraCursor.getText().createTextCursorByRange(
                                                           paraCursor.getStart());

        XSentenceCursor sc =  UnoRuntime.queryInterface(XSentenceCursor.class, cursor); 
        sc.gotoStartOfSentence(false);   // goto start
        do {
          sc.gotoEndOfSentence(true);   // select all of sentence
          if (comparer.compareRegionEnds(endPara, sc.getEnd()) > 0) {
            System.out.println("Sentence cursor passed end of current paragraph");
            break;
          }

          // move the text view cursor to highlight the current sentence
          tvc.gotoRange(sc.getStart(), false);
          tvc.gotoRange(sc.getEnd(), true);

          currSentStr = stripNonWordChars(sc.getString());
          System.out.println("  S<" + currSentStr + ">");
          if (currSentStr.length() > 0)
            speaker.say(currSentStr);
        } while (sc.gotoNextSentence(false));
      }
    } while (paraCursor.gotoNextParagraph(false));

    speaker.dispose();
  }  // end of speakSentences()



  private static String stripNonWordChars(String s)
  {  return s.replaceAll("[^\\p{L}\\p{Nd} ,\\-]+", "");  }


}  // end of TalkingBook class
