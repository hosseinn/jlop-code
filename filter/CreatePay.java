
// CreatePay.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Read 2D data from an XML file (pay.xml), and create a sheet from it.
*/ 


import org.w3c.dom.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class CreatePay 
{

  public static void main(String args[])
  {
    Document xdoc = XML.loadDoc("pay.xml");

    NodeList pays = xdoc.getElementsByTagName("payment");
    if (pays == null)
      return;

    Object[][] data = XML.getAllNodeValues(pays, 
                  new String[]{"purpose", "amount", "tax", "maturity"}); 
    Lo.printTable("payments", data);

    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);
    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    Calc.setArray(sheet, "A1", data);

    // Lo.saveDoc(doc, "payCreated.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



}  // end of CreatePay class
