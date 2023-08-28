
// GetText.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Macro function: GetText.show()

   show() is triggered when a button is pressed, and then displays 
   the text in the "Text Box 1" textfield in an Office message box.

   Calls Lo.scriptInitialize() to initialize my Utils classes.

   Also shows the use of logging to LOG_FNM and the creation of 
   a Console window, both useful for debugging.
*/


import com.sun.star.uno.*;
import com.sun.star.awt.*;
import com.sun.star.lang.*;
import com.sun.star.script.provider.XScriptContext;



public class GetText
{
  private static final String LOG_FNM = "c://macrosInfo.txt";
              // log file for storing debugging output



  public static void show(XScriptContext sc, ActionEvent e)
  // Called when a button pressed
  { 
    String controlName = Forms.getEventSourceName(e);

    FileIO.appendTo(LOG_FNM, "\"" + controlName + 
                 "\" sent ActionEvent at " + Lo.getTimeStamp());

    XComponent doc = Lo.scriptInitialize(sc);
    if (doc == null)
      return;

    // for debugging
    Console console = new Console();
    console.setVisible(true);

    Forms.listForms(doc);

    XControlModel textBox = Forms.getControlModel(doc, "Text Box 1");
    // Props.showObjProps("TextBox", textBox);

    String textContents = (String)Props.getProperty(textBox, "Text");
    GUI.showXMessageBox("Textbox text", textContents);

    console.setVisible(false);
    console.closeDown();
  } // end of show() for ActionEvent


}  // end of GetText class
