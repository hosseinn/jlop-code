
// ShowEvent.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Macro function: ShowEvent.show()

   Report when an event is detected by displaying a modeless Java dialog
   giving brief information about the event.

   I've not used any of my Util classes here so the code is standalone.

   Installation:
     > compile ShowEvent.java
     > jar cvf ShowEvent.jar ShowEvent.class

     > copy ShowEvent.jar into ShowEvent\
     > copy ShowEvent\ to <OFFICE>\share\Scripts\java\ShowEvent
*/


import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.beans.*;
// import com.sun.star.document.*;
import com.sun.star.awt.*;
import com.sun.star.script.provider.XScriptContext;


import java.awt.*;
import javax.swing.*;
import java.util.Random;



public class ShowEvent 
{

  public static void show(XScriptContext sc, ActionEvent e) 
  // Called from a button with an action
  {  display("action", getSource(e));  }


  public static void show(XScriptContext sc, TextEvent e)
  // Called when text changes inside a text component
  {  display("text", getSource(e));  }


  public static void show(XScriptContext sc, FocusEvent e)
  // Called when focus changes
  {  display("focus");  }


  public static void show(XScriptContext sc, Short val)
  // Called from a toolbar
  { display("toolbar (" + val + ")");  } 


  public static void show(XScriptContext sc, KeyEvent e)
  // Called from a key
  {  display("key " + e.KeyChar, getSource(e)); }


  public static void show(XScriptContext sc, MouseEvent e)
  // Called from a button with the mouse
  { display("mouse"); }



  public static void show(XScriptContext sc, 
                        // com.sun.star.document.EventObject e) 
                        com.sun.star.document.DocumentEvent e) 
  // Triggered by a document event
  {  display("document", e.EventName);  }



  public static void show(XScriptContext sc, EventObject e)
  /* Even though EventObject is the superclass of the event classes,
     this method is not called if there's no more specific version
     of show() that matches the triggering event. 
     In other words, this method is of no use.
  */
  {
    if (e != null)
      display("object", getSource(e)); 
    else
      display("object ()"); 
  }


  public static void show(XScriptContext sc)
  // Called from a menu or the "Run Macro..." menu
  // or used if there's no suitable event version of show()
  {  display("menu/run");   }


  // --------------------------------------------

  private static void display(String msg)
  {  
    showDialog(msg + " event"); 
    // JOptionPane.showMessageDialog(null, msg + " event");  
  }


  private static void display(String msg, String info)
  {  
    showDialog(msg + " event: " + info); 
    // JOptionPane.showMessageDialog(null, msg + " event: " + info);  
  }


  private static void showDialog(String msg)
  /* JDialog is used so the dialog is modeless, which means that
     execution will *not* suspend until the user dismisses the window.
     JOptionPane.showMessageDialog() creates a moded dialog which has 
     this suspension property.

     The dialog is created at a slightly random position, so that 
     dialogs created by other events will hopefully still be visible.

     The dialog's parent is null which means that it may or may not 
     appear in front of the Office window. 
  */
  {
    JDialog dlg = new JDialog((java.awt.Frame)null, "Show Event");
    dlg.getContentPane().setLayout(new GridLayout(3,1));
    dlg.add(new JLabel(""));
    dlg.add(new JLabel(msg, SwingConstants.CENTER));    // centered text
    dlg.pack();

    Random r = new Random();
    dlg.setLocation(100 + r.nextInt(50), 100 + r.nextInt(50));

    dlg.setVisible(true);
    dlg.setResizable(false);
    dlg.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
  }   // end of showDialog()




  private static String getSource(EventObject event)
  // get the name of the GUI component that sent this event (if possible)
  {
    XControl control = UnoRuntime.queryInterface(XControl.class, event.Source);
    XPropertySet xProps =  UnoRuntime.queryInterface(
                                  XPropertySet.class, control.getModel());
    try {
      return (String)xProps.getPropertyValue("Name");
    }
    catch (com.sun.star.uno.Exception e) {
      return "Exception?";
    }
    catch (com.sun.star.uno.RuntimeException e) {
      return "RuntimeException?";
    }
  } // end of getSource()


}  // end of ShowEvent class
