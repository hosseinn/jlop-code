
// LingoFile.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Apply the spell checker and proof reader (grammar checker) to the supplied
   file.

   The proof reader is LanguageTool, not LightProof, on one of
   my test machines.

   Usage:
     run LingoFile badGrammar.odt
     run LingoFile bigStory.doc > spellInfo.txt
*/

import java.text.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.text.*;
import com.sun.star.linguistic2.*;



public class LingoFile
{


  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run LingoFile <fnm>");
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
    GUI.setVisible(doc, true);   // Office must be visible...
    Lo.wait(2000);
    Write.openSentCheckOptions();  // for the dialog to appear
       // openThesaurusDialog()
       // openSentCheckOptions()
       // openSpellGrammarDialog()
    Lo.waitEnter();
*/

    // load spell checker, proof reader
    XSpellChecker speller = Write.loadSpellChecker();
    XProofreader proofreader = Write.loadProofreader();

    BreakIterator bi = BreakIterator.getSentenceInstance(java.util.Locale.US);

    // iterate through the doc by paragraphs
    XParagraphCursor paraCursor = Write.getParagraphCursor(doc);
    paraCursor.gotoStart(false);     // go to start of text
    String currPara;
    do {
      paraCursor.gotoEndOfParagraph(true);   // select all of paragraph
      currPara = paraCursor.getString();
      if (currPara.length() > 0)
        checkSentences(currPara, bi, speller, proofreader);
    } while (paraCursor.gotoNextParagraph(false));


    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void checkSentences(String currPara, BreakIterator bi, 
                           XSpellChecker speller, XProofreader proofreader) 
  /* Spell check and proof read each sentence in the current paragraph.
     Use a Java BreakIterator to improve slightly on the
     sentence division done by Office
  */
  { System.out.println("\n>> " + currPara);
    bi.setText(currPara);
    int lastIdx = bi.first();
    while (lastIdx != BreakIterator.DONE) {
      int firstIdx = lastIdx;
      lastIdx = bi.next();
      if (lastIdx != BreakIterator.DONE) {
        String sentence = currPara.substring(firstIdx, lastIdx);
        // System.out.println("s >> \"" + sentence + "\"");
        Write.proofSentence(sentence, proofreader);
        Write.spellSentence(sentence, speller);
      }
    }
  }  // end of checkSentences()




}  // end of LingoFile class
