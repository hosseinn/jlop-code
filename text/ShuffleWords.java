
// ShuffleWords.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Each word in the input file is mid-shuffled. This causes the
   middle letters of the word to be rearranged, but not the first
   and last letters. Words of <= 3 characters are unaffected.

   The words are highlighted as they are shuffled.

   The shuffled result is saved in "shuffled.doc".

   Usage:
     run ShuffleWords storyStart.doc
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;


public class ShuffleWords
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run ShuffleWords <fnm>");
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

    applyShuffle(doc);

    Lo.saveDoc(doc, "shuffled.doc");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void applyShuffle(XTextDocument doc)
  {
    XText docText = doc.getText();

    XWordCursor wordCursor = Write.getWordCursor(doc);
    wordCursor.gotoStart(false);     // go to start of text

    XTextViewCursor tvc = Write.getViewCursor(doc);

    String currWord;
    do {
      wordCursor.gotoEndOfWord(true);

      // move the text view cursor, and highlight the current word
      tvc.gotoRange(wordCursor.getStart(), false);
      tvc.gotoRange(wordCursor.getEnd(), true);
      currWord = wordCursor.getString().trim();
      // System.out.println(currWord);
      if (currWord.length() > 0) {
        Lo.wait(250);   // slow down so user can see selection before change
        docText.insertString(wordCursor, midShuffle(currWord), true);
      }
    } while( wordCursor.gotoNextWord(false));
  }  // end of applyShuffle()



  private static String midShuffle(String s)
  // shuffle the middle letters of a string
  {
    if (s.length() <= 3)   // not long enough
      return s;

    // extract middle of the string for shuffling
    String mid = s.substring(1, s.length()-1);

    // shuffle a list of characters  made from the middle
    ArrayList<Character> midList = new ArrayList<>();
    for(char ch :  mid.toCharArray())
        midList.add(ch); 
    Collections.shuffle(midList);
    
    // rebuild string, adding back the first and last letters
    StringBuilder sb = new StringBuilder();
    sb.append(s.charAt(0));
    for(char ch : midList)
      sb.append(ch);
    sb.append(s.charAt(s.length()-1));

    return sb.toString();
  }  // end of midShuffle()


}  // end of ShuffleWords class
