
// ExtractText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Feb, 2017

/* Attempt to extract the text from the slide deck. 

   The order of the text extracted may not be the same as the order
   that it appears in the file; it depends
   on the order that the text shapes are saved inside the file.

   Usage:
     run ExtractText algs.ppt
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

    if (Draw.isShapesBased(doc)) {
      String text = Draw.getShapesText(doc);
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

