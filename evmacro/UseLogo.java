
// UseLogo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov 2016

/* Create a new text document and use the Python LibreLogo macro
   to draw a pretty pattern. 

   Press <enter> to close the document.

   Usage:  run UseLogo
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;

import com.sun.star.ui.*;


public class UseLogo
{
  private static final String LOGO_BAR = 
           "private:resource/toolbar/addon_LibreLogo.OfficeToolBar";


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
    Lo.wait(1000);   // make sure doc is visible before using dispatches

    // XLayoutManager lm = GUI.getLayoutManager(doc);
    // GUI.showUIElement(lm, LOGO_BAR, true);

    XTextCursor cursor = Write.getCursor(doc);

    String logoCmds = "repeat 88 [ fd 200 left 89 ] fill";
    Write.appendPara(cursor, logoCmds);
    Macros.executeLogoCmds(logoCmds);

    // GUI.showUIElement(lm, LOGO_BAR, false);

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of UseLogo class

