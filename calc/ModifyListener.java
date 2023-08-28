
// ModifyListener.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Creates a table and attaches a modify listener
    to it. When the user chamges a cell, the change
    is printed (twice).
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.view.*;

import com.sun.star.util.*;
import com.sun.star.awt.*;



public class ModifyListener implements XModifyListener
{
  private XSpreadsheetDocument doc;
  private XSpreadsheet sheet;


  public ModifyListener()
  {
    XComponentLoader loader = Lo.loadOffice();
    doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);
    sheet = Calc.getSheet(doc, 0);


    // insert some data
    Calc.setCol(sheet, "A1",
            new Object[] {"Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3});

    XModifyBroadcaster mb = Lo.qi(XModifyBroadcaster.class, doc);
    mb.addModifyListener(this);

    // close down when window closes
    XExtendedToolkit tk = Lo.createInstanceMCF(XExtendedToolkit.class, 
                                       "com.sun.star.awt.Toolkit");
    if (tk != null)
      tk.addTopWindowListener( new XTopWindowAdapter() {
        public void windowClosing(EventObject eo)
        { System.out.println("Closing");  
          Lo.saveDoc(doc, "modify.ods");
          Lo.closeDoc(doc);
          Lo.closeOffice();
        }
      });

    // Lo.saveDoc(doc, "modify.ods");
    // Lo.waitEnter();
    // Lo.closeDoc(doc);
    // Lo.closeOffice();
  }  // end of ModifyListener()





  // ---------------- XModifyListener methods ---------------


  public void disposing(EventObject event)
  {  System.out.println("Disposing"); }  



  public void modified(EventObject event)
  /* is fired twice for each change to a cell, even if the cell is
     edited to be the same as before ??
  */
  { System.out.println("Modified");
    // System.out.println("Modified: " + event.Source);
    // Info.showServices("Event source", event.Source);
    XSpreadsheetDocument doc = Lo.qi(XSpreadsheetDocument.class, event.Source);

    CellAddress addr = Calc.getSelectedCellAddr(doc);
    System.out.println("  " + Calc.getCellStr(addr) + " = " + 
                                      Calc.getVal(sheet, addr));
  }  // end of modified()



  // -----------------------------------------------

  public static void main(String args[])
  {  new ModifyListener();  }

}  // end of ModifyListener class
