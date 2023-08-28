
// OffViewer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016

/*  Display a document in an undecorated Office window.
    Slightly different toolbars are shown depending on the
    document type.

    Include mouse, key, and window listeners

    Some times there is a SEH exception on closing: see
        https://ask.libreoffice.org/en/question/64870/
                    seh-exception/?answer=65020#post-id-65020

    After closing the SEH dialog, no Office process is left running.

  Usage:
    run OffViewer classic_letter.ott

    run OffViewer skinner.png
    run OffViewer algs.odp
    run OffViewer tables.ods
        -- read-only mode yellow banner

    run OffViewer liangTables.odb
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;
import com.sun.star.document.*;

import com.sun.star.ui.*;



public class OffViewer implements XMouseClickHandler, XKeyHandler,
                                  XWindowListener,
                                  com.sun.star.document.XEventListener
{
  private XComponent doc = null;


  public OffViewer(String fnm)
  {
    XComponentLoader loader = Lo.loadOffice();
    doc = Lo.openReadOnlyDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    // attach listeners
    doc.addEventListener(this);

    XWindow win = GUI.getWindow(doc);
    win.addWindowListener(this);
    win.setVisible(true);
    Lo.delay(500);   // give window time to appear, or dispatches will fail

    XUserInputInterception uii = GUI.getUII(doc);
    uii.addMouseClickHandler(this);
    uii.addKeyHandler(this);

    // modify UI visibility
    ArrayList<String> showElems = new ArrayList<String>();
    showElems.add(GUI.FIND_BAR);
    showElems.add(GUI.STATUS_BAR);
    GUI.showOnly(doc, showElems);

    // call dispatches *after* the window is visible
    Lo.dispatchCmd("Sidebar",
                 Props.makeProps("Sidebar", false));    // hide sidebar


    // modify UI based on the document type
    int docType = Info.reportDocType(doc);
    if (docType == Lo.DRAW)
      Lo.dispatchCmd("LeftPaneDraw",
                  Props.makeProps("LeftPaneDraw", false) );    // hide Pages pane

    if (docType == Lo.IMPRESS)
      Lo.dispatchCmd("LeftPaneImpress",
                  Props.makeProps("LeftPaneImpress", true) );    // show Slides pane

    if (docType == Lo.CALC)
      Lo.dispatchCmd("InputLineVisible",
                    Props.makeProps("InputLineVisible", false) );    // hide formula bar

    if (docType == Lo.BASE)
      Lo.dispatchCmd("DBViewTables");    // switch to tables view

    // report on toolbar commands
    // GUI.printUICmds(doc, GUI.STATUS_BAR);
    // GUI.printAllUICommands(doc);
  }  // end of OffViewer()



  // ------------------ XMouseClickHandler methods -------------------


  public boolean mousePressed(MouseEvent e)
  { System.out.println("Mouse pressed (" + e.X + ", " + e.Y + ")");
    return false;    // send event on, or use true not to send
  }

  public boolean mouseReleased(MouseEvent e)
  { System.out.println("Mouse released (" + e.X + ", " + e.Y + ")");
    return false;
  }


  // ---------------------- XKeyHandler methods -------------------------

  public boolean keyPressed(KeyEvent e)
  { System.out.println("Key pressed: " + e.KeyCode + "/" + e.KeyChar);
    return true;
  }

  public boolean keyReleased(KeyEvent e)
  { System.out.println("Key released: " + e.KeyCode + "/" + e.KeyChar);
    return true;
  }


  // ---------------------- XWindow listener methods -----------------------


  public void windowShown(com.sun.star.lang.EventObject e)
  {  System.out.println("Doc window has become visible");  }

  public void windowHidden(com.sun.star.lang.EventObject e)
  {  System.out.println("Doc window has been hidden");  }

  public void windowResized(WindowEvent e)
  {  System.out.println("Resized to " + e.Width + " x " + e.Height);  }

  public void windowMoved(WindowEvent e)
  {  System.out.println("Moved to (" + e.X + ", " + e.Y + ")");  }



  // ---------------------- event listener methods -----------------------

  public void notifyEvent(com.sun.star.document.EventObject e)
  {  System.out.println("Document is changed: " + e.EventName); }



  public void disposing(com.sun.star.lang.EventObject e)
  {
     System.out.println("Document is closing");
     if (doc != null)
       Lo.closeDoc(doc);
     // Lo.closeOffice();   // Office hangs, so kill it instead
     Lo.killOffice();
     System.exit(0);
  }



  // ------------------------------------------------------

  public static void main(String[] args)
  {
    if (args.length != 1)
      System.out.println("Usage: OffViewer fnm");
    else
      new OffViewer(args[0]);
  } // end of main()

}  // end of OffViewer class
