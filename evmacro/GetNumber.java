
// GetNumber.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Macro function: GetNumber.get()

   Extract an integer from a form textfield.
   Display the extracted value in a dialog which allows the user
   to accept or decline the number.

   The dialog is either loaded from the extension (NumExtractor.xdl)
   by calling loadXDLDialog(), or built at run time by calling
   runtimeDialog().
*/


import com.sun.star.uno.*;
import com.sun.star.awt.*;
import com.sun.star.lang.*;

import com.sun.star.script.provider.XScriptContext;



public class GetNumber
{
  private static final String LOG_FNM = "c://macrosInfo.txt";
              // log file for debugging output


  public static void get(XScriptContext sc, KeyEvent e)
  // Called for key events
  { 
    String controlName = Forms.getEventSourceName(e);
    // log(controlName + " sent '" + e.KeyChar + "'");
    if (e.KeyCode == Key.RETURN) {  // return typed
      XComponent doc = Lo.scriptInitialize(sc);
      if (doc != null) {
        XControlModel cModel = Forms.getControlModel(doc, controlName);
        if (Forms.isTextField(cModel))  // the form control is a textfield
          // loadXDLDialog(cModel);
          runtimeDialog(cModel);
      }
    }
  } // end of get() & KeyEvent




  private static void loadXDLDialog(XControlModel cModel)
  // load dialog from XDL file
  /* start the dialog that checks if the number extracted from
     the text field control (cModel) is acceptable to the user
  */
  {
    XDialog dialog = Dialogs.loadAddonDialog("org.openoffice.formmacros", 
                                         "dialogLibrary/NumExtractor.xdl");
    // log("dialog: " + dialog);
    if (dialog == null)
      return;

    initDialog(dialog, cModel);
    dialog.execute();
 }  // end of loadXDLDialog()


  private static void log(String msg)
  {   FileIO.appendTo(LOG_FNM, msg);  }


  private static void initDialog(XDialog dialog, XControlModel cModel)
  {
    XControl dialogCtrl = Dialogs.getDialogControl(dialog);
    if (dialogCtrl == null)
      return;

    int val = extractDigits( (String)Props.getProperty(cModel, "Text"));

    // store extracted number in dialog's read-only text field;
    // the names of the controls are hardwired
    XTextComponent numFieldTB = Lo.qi(XTextComponent.class,
                              Dialogs.findControl(dialogCtrl, "TextField1"));
    numFieldTB.setText(""+val);

    // assign same listener to both buttons
    NumActionListener naListener = new NumActionListener(dialog, cModel, val);

    XButton okButton = Lo.qi(XButton.class,
                              Dialogs.findControl(dialogCtrl, "CommandButton1"));
    okButton.addActionListener(naListener);

    XButton cancelButton = Lo.qi(XButton.class,
                              Dialogs.findControl(dialogCtrl, "CommandButton2"));
    cancelButton.addActionListener(naListener);

  } // end of initDialog()



  private static int extractDigits(String s)
  // extract the digits from s, returning them as a single int
  {
    StringBuilder sb = new StringBuilder(s.length());
    for(int i = 0; i < s.length(); i++){
        char ch = s.charAt(i);
        if(ch > 47 && ch < 58)   // between '0' and '9'
            sb.append(ch);
    }
    return Lo.parseInt(sb.toString());
  }  // end of extractDigits()



  // ---------------------------------------------------------------------


  private static void runtimeDialog(XControlModel cModel)
  // build the dialog at run time instead of loading it
  /* 
     start the dialog that checks if the number extracted from
     the text field control (cModel) is acceptable to the user
  */
  {
    XControl dialogCtrl = makeDialogControl();
    if (dialogCtrl == null)
      return;

    XDialog dialog = Dialogs.createDialogPeer(dialogCtrl);
    // log("dialog: " + dialog);
    if (dialog == null)
      return;

    initDialog(dialog, cModel);
    dialog.execute();
  }  // end of runtimeDialog()



  private static XControl makeDialogControl()
  // (x,y), width and heights taken from XDL file
  {
    XControl dialogCtrl = Dialogs.createDialogControl(109, 73, 94, 44, 
                                                          "Number Extractor");
    if (dialogCtrl == null)
      System.out.println("dialog control is null");
    // log("Dialog name:" + Dialogs.getControlName(dialogCtrl));
          // OfficeDialog1

    XControl xc = Dialogs.insertLabel(dialogCtrl, 8, 11, 48, "Extracted Number: ");
    // log("Label name:" + Dialogs.getControlName(xc));
          // FixedText1

    xc = Dialogs.insertTextField(dialogCtrl, 61, 9, 24, "");
    // log("Text field name:" + Dialogs.getControlName(xc));
          // TextField1

    xc = Dialogs.insertButton(dialogCtrl, 9, 27, 33, "Ok");
    // log("Ok button name:" + Dialogs.getControlName(xc));
         // CommandButton1

    xc = Dialogs.insertButton(dialogCtrl, 52, 27, 33, "Cancel");
    // log("Cancel button name:" + Dialogs.getControlName(xc));
         // CommandButton2

    return dialogCtrl;
  }  // end of makeDialogControl()


}  // end of GetNumber class

