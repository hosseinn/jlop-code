
// CreateAssoc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Read 2D data from an XML file (clubs.xml), and create a sheet from it.
*/ 


import org.w3c.dom.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import java.util.*;



public class CreateAssoc 
{

  public static void main(String args[])
  {
    Document xdoc = XML.loadDoc("clubs.xml");
    NodeList root = xdoc.getChildNodes();

    // get the first association
    Node cdb = XML.getNode("club-database", root);
    Node assoc1 = XML.getNode("association", cdb.getChildNodes());
    NodeList clubs = assoc1.getChildNodes();

    Object[][] data = XML.getAllNodeValues(clubs, 
       new String[]{"name", "contact", "location", "phone", "email"}); 

    Lo.printTable("clubs", data);


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

    // Lo.saveDoc(doc, "clubsCreated.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of CreateAssoc class
