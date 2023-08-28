
// TextMacro.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Use Macros.execute() to execute a variety of macros coded in 
   different languages, all of thm added variants of "Hello World"
   to the open text document.

   Usage:  run TextMacro
      - the document remains open until the user types enter
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;


public class TextMacro
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

    if (Macros.getSecurity() == Macros.LOW)
      Macros.setSecurity(Macros.MEDIUM);

    XTextCursor cursor = Write.getCursor(doc);
    GUI.setVisible(doc, true);
    Lo.wait(1000);   // make sure the document is visible before sending it dispatches

    Write.appendPara(cursor, "Hello LibreOffice");

    /* I used FindMacros.java to find the names of these scripts which
        all mention "hello"
    */
    Macros.execute("HelloWorld.helloworld.bsh", "BeanShell", "share");
    Write.endParagraph(cursor);

    Macros.execute("HelloWorld.py$HelloWorldPython", "Python", "share");
    Write.endParagraph(cursor);

    Macros.execute("HelloWorld.helloworld.js", "JavaScript", "share");
    Write.endParagraph(cursor);

    Macros.execute("HelloWorld.org.libreoffice.example.java_scripts.HelloWorld.printHW", 
                                                          "Java", "share");

    Write.endParagraph(cursor);
    Write.appendPara(cursor, "Timestamp: " + Lo.getTimeStamp());

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of TextMacro class

