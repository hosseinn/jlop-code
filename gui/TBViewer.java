
// TBViewer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016

/* Display a Writer document in an Office window (like in
   OffViewer), *and* add a new item to the beginning of the
   'find' toolbar.

   The new item is the word "Hello" and an icon (which does
   not appear for some reason). When "Hello" is clicked, an 
   Office dialog window appears.

   When "Hello" is clicked, the ToolbarItemListener.clicked() 
   method is invoked, which is defined in this class.


   ToolbarItemListener is my listener (in Utils/) which is 
   invoked by my ItemInterceptor object (also defined in Utils/), 
   which is registered as a 
   dispatch interceptor in the TBViewer() constructor.

   Usage:
       run TBViewer classic_letter.ott


   See MenuViewer.java for how to add an item to a menu bar.
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;

import com.sun.star.beans.*;
import com.sun.star.util.*;



public class TBViewer implements ToolbarItemListener,
                                 XWindowListener
{
  private XComponent doc = null;


  public TBViewer(String fnm)
  {
    XComponentLoader loader = Lo.loadOffice();
    doc = Lo.openReadOnlyDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    XWindow win = GUI.getWindow(doc);
    win.addWindowListener(this);
    win.setVisible(true);
    Lo.delay(500);   // give window time to appear, or dispatches may fail

    GUI.printUIs(doc);
    GUI.showOne(doc, GUI.FIND_BAR);

    // call dispatches *after* the window is visible
    Lo.dispatchCmd("Sidebar",
                 Props.makeProps("Sidebar", false));    // hide sidebar


    GUI.addItemToToolbar(doc, GUI.FIND_BAR, "Hello", "h.png");

    // add an dispatch interceptor for the "Hello" item
    XDispatchProviderInterception dpi = GUI.getDPI(doc);
    if (dpi != null)
      dpi.registerDispatchProviderInterceptor(
                                new ItemInterceptor(this, "Hello"));
  }  // end of TBViewer()




  // ------------------- ToolbarItemListener method ---------------------------


  public void clicked(String itemName, URL cmdURL, PropertyValue[] props)
  // triggered when the user-defined toobar item is clicked
  {  
     System.out.println("Clicked on " + itemName);
     // GUI.showJMessageBox("Item Dialog", "Processing the \"" + itemName + "\" command");
              // may not appear in front of Office window
     GUI.showMessageBox("Item Dialog", "Processing the \"" + itemName + "\" command");
  }  // end of clicked()




  // ---------------------- XWindow listener methods -----------------------


  public void disposing(com.sun.star.lang.EventObject e)
  {  System.out.println("Doc window is closing");  
     if (doc != null)
        Lo.closeDoc(doc);
     // Lo.closeOffice();    // Office hangs, so kill it instead
     Lo.killOffice(); 
     System.exit(0);
  } 

  public void windowShown(com.sun.star.lang.EventObject e) {} 

  public void windowHidden(com.sun.star.lang.EventObject e)  {} 

  public void windowResized(WindowEvent e){} 

  public void windowMoved(WindowEvent e) {} 


  // ------------------------------------------------------

  public static void main(String[] args) 
  {
    if (args.length != 1)
      System.out.println("Usage: TBViewer fnm");
    else
      new TBViewer(args[0]);
  } // end of main()

}  // end of TBViewer class
