
// MenuViewer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016

/* Display a Writer document in an Office window (like in
   OffViewer), *and* extend the menu bar with a new menu,
   *or* create a new menu bar with the menu.

   The "My Menu" menu bar includes a "Hello" item 
   with an icon, a "QUIT" item for quitting Office, 
   and some checkbox and radio button menu items that do nothing.

   When the "Hello" item is clicked, an Office dialog window
   appears. 

   See TBViewer.java for adding this item to a toolbar.

   Usage:
      run MenuViewer classic_letter.ott
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;
import com.sun.star.document.*;



public class MenuViewer implements XMenuListener,
                                   XWindowListener
{
  private static final String CMD_HELLO = "Cmd_Hello";
  private static final String CMD_QUIT = "Cmd_Quit";

  private XComponent doc = null;

  private short id = 0;   // used for menu IDs



  public MenuViewer(String fnm)
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
    Lo.delay(500);   // wait for window to appear


    // create new menu bar
    GUI.showNone(doc);
    makeMenubar(win);
/*
   // or modify the existing menubar
    GUI.showOne(doc, GUI.MENU_BAR);
    setMenubar(doc);
*/
  }  // end of MenuViewer()




  private void setMenubar(XComponent doc)
  // add menu to end of existing menu bar
  {
    XLayoutManager lm = GUI.getLayoutManager(doc);
    XMenuBar menubar = GUI.getMenubar(lm);
    if (menubar == null)
      System.out.println("Menubar not found"); 
    else {
      short menuCount = menubar.getItemCount();  
              // count along the menu bar to get no. of menus
      short maxID = GUI.getMenuMaxID(menubar);   
              // find the largest menu item ID
      id = (short)(maxID + (short)100);    // hacky
        /* add a big number to account for the IDs used by
           menu items in the menu
        */
      addMenu(menubar, menuCount);
    }
  }  // end of setMenubar()



  private void makeMenubar(XWindow win)
  // add menu to a new menu bar
  {
    XTopWindow topWin = Lo.qi(XTopWindow.class, win);
    XMenuBar menubar = Lo.createInstanceMCF(XMenuBar.class, 
                                         "com.sun.star.awt.MenuBar");
    if (menubar == null)
      System.out.println("Could not create a menu bar"); 
    else {
      addMenu(menubar, (short)0);
      topWin.setMenuBar(menubar);
    }
  }  // end of makeMenubar()
  



  private void addMenu(XMenuBar bar, short mPos)
  // add menu to the menu bar at position mPos
  {
    short menuID = id;
    bar.insertItem(menuID, "My ~Menu",    // 'F' may be used by "File" menu
                         MenuItemStyle.AUTOCHECK, mPos);
        // first short is an ID for the new menu's name; 
        // the second (last arg) is the menu's position on the menubar
    id++;
    bar.setPopupMenu(menuID, makePopupMenu());   
                          // add menu to menubar
  }  // end of addMenu()




  public XPopupMenu makePopupMenu()
  /* Popup menu consists of:
            * Hello + image
            * three radio buttons (only one on)
            * a separator
            * two check boxes
            * Quit
  */
  {
    XPopupMenu pum = Lo.createInstanceMCF(XPopupMenu.class, "com.sun.star.awt.PopupMenu");
    if (pum == null) {
      System.out.println("Could not create a menu");
      return null;
    }

    short mPos = 0;
    pum.insertItem(id, "Hello", (short)0, mPos++);
        // first short is an ID for the new menu item; 
        // the second is the style for the item, 0 means none/ordinary;
        // the third is the items's position on the menu

    pum.setItemImage(id, Images.loadGraphicFile("H.png"), false);
    pum.setCommand(id++, CMD_HELLO);

    pum.insertItem(id, "First Radio",
               (short) (MenuItemStyle.RADIOCHECK + 
                        MenuItemStyle.AUTOCHECK), mPos++);
    pum.enableItem(id++, false);   // grayed out

    pum.insertItem(id, "Second Radio",
               (short) (MenuItemStyle.RADIOCHECK + 
                        MenuItemStyle.AUTOCHECK), mPos++);
    pum.checkItem(id++, true);  // selected


    pum.insertItem(id++, "Third Radio",
               (short) (MenuItemStyle.RADIOCHECK + 
                        MenuItemStyle.AUTOCHECK), mPos++);

    pum.insertSeparator(mPos++);

    pum.insertItem(id++, "Check 1",
               (short) (MenuItemStyle.CHECKABLE + 
                        MenuItemStyle.AUTOCHECK), mPos++);

    pum.insertItem(id++, "Check 2",
               (short) (MenuItemStyle.CHECKABLE + 
                        MenuItemStyle.AUTOCHECK), mPos++);

    pum.insertItem(id, "Quit", (short)0, mPos++);
    pum.setCommand(id++, CMD_QUIT);


    pum.addMenuListener(this);

    return pum;
  }  // end of makePopupMenu()





  // -------------------- XMenuListener methods ------------------------
  // listen to the menu


  public void itemSelected(MenuEvent menuEvent)
  {  
    short id = menuEvent.MenuId;
    XPopupMenu pum = Lo.qi(XPopupMenu.class, menuEvent.Source);
    if (pum == null)
      System.out.println("Menu item " + id + " selected; popupmenu is null");
    else {
      String itemName = pum.getItemText(id);
      String cmd = pum.getCommand(id);
      System.out.println("Menu item \"" + pum.getItemText(id) + 
                                    "\" selected");
      System.out.println("  is checked? " + pum.isItemChecked(id));
      processCmd(cmd);
    }
  }  // end of itemSelected()



  public void processCmd(String cmd)
  {
    if ((cmd == null) || cmd.equals("")) {
      System.out.println("  No command");
      return;
    }

    if (cmd.equals(CMD_HELLO))
      GUI.showMessageBox("Item Dialog", 
                         "Processing the \"Hello\" command");
    else if (cmd.equals(CMD_QUIT)) {
      System.out.println("  Quiting the Application");  
      if (doc != null)
        Lo.closeDoc(doc);    // this will trigger a call to disposing()
    }
    else
      System.out.println("  Got command: " + cmd);  
  }  // end of processCmd()



  public void itemHighlighted(MenuEvent menuEvent)
  { // System.out.println("Menu " + menuEvent.MenuId + " highlighted");
  } 


  public void itemDeactivated(MenuEvent menuEvent)
  { // System.out.println("Menu " + menuEvent.MenuId + " deactivated");
  }


  public void itemActivated(MenuEvent menuEvent)
  { // System.out.println("Menu " + menuEvent.MenuId + " activated");
  }



  // ---------------------- XWindow listener methods -----------------------


  public void disposing(com.sun.star.lang.EventObject e)
  {  System.out.println("Doc window is closing");  
     if (doc != null)
        Lo.closeDoc(doc);
     // Lo.closeOffice();  // Office hangs, so kill it instead
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
      System.out.println("Usage: MenuViewer fnm");
    else
      new MenuViewer(args[0]);
  } // end of main()

}  // end of MenuViewer class
