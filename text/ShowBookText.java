
// ShowBookText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Print all the text in every paragraph using enumeration access.

   Usage:
     > compile ShowBookText.java

     > run ShowBookText storyStart.doc
     > run ShowBookText simpleTextTest.odt
     > run ShowBookText table.odt

*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.container.*;


public class ShowBookText
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run ShowBookText <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    // GUI.setVisible(doc, true);
    printParas(doc.getText());

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void printParas(XText xText)
  /* iterate through the document contents, printing all 
     the text portions in each paragraph */
  {
    try { 
      XEnumeration textEnum = Write.getEnumeration(xText);
      while (textEnum.hasMoreElements()) {
        XTextContent tc = UnoRuntime.queryInterface(
                                      XTextContent.class, textEnum.nextElement());
               // return a paragraph (or text table)
        if (!Info.supportService(tc, "com.sun.star.text.TextTable")) {
          // print all the text portions in the paragraph
          System.out.println("P--");
          XEnumeration paraEnum = Write.getEnumeration(tc);
          while (paraEnum.hasMoreElements()) {
            XTextRange txtRange = UnoRuntime.queryInterface(
                                        XTextRange.class, paraEnum.nextElement());
               // return a text portion
            System.out.println("  " + Props.getProperty(txtRange, "TextPortionType") +
                               " = \"" + txtRange.getString() + "\"");
          }
        }
        else
          System.out.println("Text table");
      }
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of printParas()


}  // end of ShowBookText class
