
// ShowSheet.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Open a spreadsheet and display it on-screen.

   Optional extras: open read-only, protect the document/sheet
                    read in a password, save to a new file format

   Usage:
      run ShowSheet totals.ods
      run ShowSheet sorted.csv
      run ShowSheet results.xlsx

      run ShowSheet totals.ods show.pdf
      run ShowSheet totals.ods show.html   // or xhtml
      run ShowSheet totals.ods show.xlsx
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.util.*;


public class ShowSheet
{

  public static void main(String args[])
  {  
    String outFnm = null;
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: run ShowSheet fnm [out-fnm]");
      return;
    }
    if (args.length == 2)
      outFnm = args[1];
    
    XComponentLoader loader = Lo.loadOffice();

    //XComponent cDoc = Lo.openReadOnlyDoc(args[0], loader);
    //XSpreadsheetDocument doc = Calc.getSSDoc(cDoc);
    XSpreadsheetDocument doc = Calc.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }
/*
    XProtectable pro = Lo.qi(XProtectable.class, doc);
    pro.protect("foobar");
    System.out.println("Is protected: " + pro.isProtected());
*/

    GUI.setVisible(doc, true);
    // Calc.zoom(doc, Calc.ENTIRE_PAGE);
    Calc.gotoCell(doc, "A1");

    // XSpreadsheet sheet = Calc.getSheet(doc, 0);

    String[] sheetNames = Calc.getSheetNames(doc);
    System.out.println("Names of sheets (" + sheetNames.length + "):");
    for(String sheetName : sheetNames)
      System.out.println("  " + sheetName);

    XSpreadsheet sheet = Calc.getSheet(doc, "Sheet1");
    Calc.setActiveSheet(doc, sheet);


    XProtectable pro = Lo.qi(XProtectable.class, sheet);
    pro.protect("foobar");
    System.out.println("Is protected: " + pro.isProtected());

    String pwd = GUI.getPassword("Sheet protection", "Supply password:");
    if ((pwd != null) && pwd.equals("foobar")) {
      System.out.println("Password correct");
      pro.unprotect("foobar");
    }
    else {
      System.out.println("Password incorrect");
      GUI.showJMessageBox("Password Status", "Password incorrect");
    }


    Lo.waitEnter();

    if (outFnm != null)
      Lo.saveDoc(doc, outFnm);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of ShowSheet class
