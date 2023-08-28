
// SelectListener.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Creates a table and attaches a selection change listener
    to it. 
     - Report when the selection changes by printing the name of the 
       previously selected cell and the new one;
     - Report if cell just left has a new or changed numerical value;
     - Report if cell just entered has a numerical value.

*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.view.*;
import com.sun.star.awt.*;



public class SelectListener implements XSelectionChangeListener
{
  private XSpreadsheetDocument doc;
  private XSpreadsheet sheet;

  private CellAddress currAddr;
  private Double currVal = null; 


  public SelectListener()
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

    currAddr = Calc.getSelectedCellAddr(doc);
    currVal = getCellDouble(sheet, currAddr);   // may be null


    attachListener(doc);

    // insert some data
    Calc.setCol(sheet, "A1",
            new Object[] {"Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3});

    // close down when window closes
    XExtendedToolkit tk = Lo.createInstanceMCF(XExtendedToolkit.class, 
                                       "com.sun.star.awt.Toolkit");
    if (tk != null)
      tk.addTopWindowListener( new XTopWindowAdapter() {
        public void windowClosing(EventObject eo)
        { System.out.println("Closing");  
          Lo.saveDoc(doc, "select.ods");
          Lo.closeDoc(doc);
          Lo.closeOffice();
        }
      });
  }  // end of SelectListener()



  private void attachListener(XSpreadsheetDocument doc)
  {
    // start listening for selection changes
    XController ctrl = Calc.getController(doc);
    XSelectionSupplier supp = Lo.qi(XSelectionSupplier.class, ctrl);
    if (supp == null)
      System.out.println("Could not attach selection change listener");
    else {
      supp.addSelectionChangeListener(this);
      System.out.println("Successfully attached selection change listener");
    }
  }  // end of attachListener()




  // ---------------- XSelectionChangeListener methods ---------------


  public void disposing(EventObject event)
  {  System.out.println("Disposing"); } 



  public void selectionChanged(EventObject event)
  /* is fired four times for every click, and twice for shift arrow keys (?)
     Once for arrow keys.
     - Report when the selection changes by printing the name of the 
       previously selected cell and the new one;
     - Report if cell just left has a new or changed numerical value;
     - Report if cell just entered has a numerical value.
  */
  {
    //System.out.println("Modified: " + event.Source);
    //Info.showServices("Event source", event.Source);
    XController ctrl = Lo.qi(XController.class, event.Source);
    if (ctrl == null){
      System.out.println("No ctrl for event source");
      return;
    }

    CellAddress addr = Calc.getSelectedCellAddr(doc);
    if (addr == null)
      return;
     
    if (!Calc.isEqualAddresses(addr, currAddr)) {
      System.out.println( Calc.getCellStr(currAddr) + " --> " + 
                          Calc.getCellStr(addr));

      // check if currAddr value has changed
      Double d = getCellDouble(sheet, currAddr);    // value right now
      if (d != null) {
        if (currVal == null)  // so previously stored value was null
          System.out.println( Calc.getCellStr(currAddr) +
                                                 " new value: " + d);
        else {   // currVal has a value; is it different from d?
          if (currVal.doubleValue() != d.doubleValue())
            System.out.println( Calc.getCellStr(currAddr) +
                       " has changed from " + currVal + " to " + d);
        }
      }

      // update current address and value
      currAddr = addr;
      currVal = getCellDouble(sheet, addr);
      if (currVal != null)
        System.out.println( Calc.getCellStr(currAddr) + " value: " + currVal);
    }
  }  // end of selectionChanged()



  private Double getCellDouble(XSpreadsheet sheet, CellAddress addr)
  {
    Object obj = Calc.getVal(sheet, addr);
    if (obj instanceof Double)
      return (Double)obj;
    else
      return null;
  } // end of getCellDouble()





  // -----------------------------------------------

  public static void main(String args[])
  {  new SelectListener();  }

}  // end of SelectListener class
