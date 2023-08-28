
// HelloText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Create a new swriter documents, add a few lines, and save it
   as the file "hello" and the specified extension.

   Tested extensions are: doc, docx, rtf, odt, pdf, txt,
   and other extensions are treated like text by default (e.g. "java")
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;



public class HelloText
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);

    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);    // to make the construction visible

    XTextCursor cursor = Write.getCursor(doc);
    cursor.gotoEnd(false);   // make sure at end of doc before appending

    Write.appendPara(cursor, "Hello LibreOffice.\n");
    Lo.wait(1000);   // to slow things down so they can be seen

    Write.appendPara(cursor, "How are you?.");
    Lo.wait(2000);

    Lo.saveDoc(doc, "hello.odt");
                   // doc, docx, rtf, odt, pdf, txt
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of HelloText class

