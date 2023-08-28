
// StoreInCalc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016


/* Open a database as a data source rather than as a document.
   Copy data over to a Calc document.

*/

import java.util.*;

import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.uno.Exception;



public class StoreInCalc
{
  private static final String FNM = "liangTables.odb";
  private static final String CALC_FNM = "liangCalc.ods";



  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    System.out.println("Created a Calc document");

    XSpreadsheet sheet = Calc.getSheet(doc, 0);
    if (sheet == null) {
      System.out.println("No spreadsheet found");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    Object[][] rsa = queryDB(FNM,
                     "SELECT \"firstName\", \"lastName\" FROM \"Student\"");
    if (rsa != null)
      fillSheet(sheet, rsa);
    Lo.delay(3000);

    Lo.saveDoc(doc, CALC_FNM);
    Lo.closeDoc(doc);
    Lo.closeOffice();

    // Lo.killOffice();   // needed when getFileDataSource() used 
  }  // end of main()



  private static Object[][] queryDB(String fnm, String query)
  // open fnm as a data source, and execute query
  {
    Object[][] rsa = null;
    XConnection conn = null;
    try {
      XDataSource dataSource = Base.getFileDataSource(fnm);
      conn = dataSource.getConnection("", ""); 

      XResultSet rs = Base.executeQuery(query, conn);
      rsa = Base.getResultSetArr(rs);
      // BaseTablePrinter.printResultSet(rs);  
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    return rsa;
  }  // end of queryDB()



  private static void fillSheet(XSpreadsheet sheet, Object[][] rsa)
  // put the data in rsa into the sheet starting at cell (0,0)
  {
    try {
      XCellRange cellRange = sheet.getCellRangeByPosition(0, 0,
                                       rsa[0].length-1, rsa.length-1);
      XCellRangeData xData = Lo.qi(XCellRangeData.class, cellRange);
      xData.setDataArray(rsa);
    }
    catch (Exception e) {
      System.out.println("Could not writes values to cells");
    }
  }  // end of fillSheet()


}  // end of StoreInCalc class
