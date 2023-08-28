
// CopyResultSet.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Use SQL command to query a database, and copy the 
   result to the clipboard as a 2D array


   Usage:
     > run CopyResultSet

   Based on DBQuery.java in Base Tests/
*/

import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;




public class CopyResultSet
{
  private static final String FNM = "liangTables.odb";



  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(FNM, loader);
    if (dbDoc == null) {
      System.out.println("Could not open database " + FNM);
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", ""); // no login/password

      XResultSet rs = Base.executeQuery("SELECT * FROM \"Course\"", conn);

      // BaseTablePrinter.printResultSet(rs); 
      // rs.beforeFirst();   // fails if ResultSet is TYPE_FORWARD_ONLY, 
                          // which is the default

      Object[][] rsArr = Base.getResultSetArr(rs);
      Base.printResultSetArr(rsArr);

      JClip.setArray(rsArr);

      // Clip.listFlavors();
      JClip.listFlavors();

      System.out.println("Saving array from clipboard");
      FileIO.saveArray("queryResults.txt", JClip.getArray());
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()



}  // end of CopyResultSet class
