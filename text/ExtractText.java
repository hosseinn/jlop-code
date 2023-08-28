
// ExtractText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Feb. 2017

/* Attempt to extract the text from the document. 

   Usage:
     run ExtractText storyStart.doc
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;



public class ExtractText
{
  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: ExtractText fnm");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    if (Info.isDocType(doc, Lo.WRITER_SERVICE)) {
      XTextDocument textDoc = Write.getTextDoc(doc);
      XTextCursor cursor = Write.getCursor(textDoc);
      String text = Write.getAllText(cursor);
      System.out.println("----------------- Text Content ---------------");
      System.out.println(text);
      System.out.println("----------------------------------------------");
    }
    else
      System.out.println("Text extraction unsupported for this document type");

    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of ExtractText class

