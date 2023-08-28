
// AccessQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use the DriverManager to query a MS Access file
   Note: using a database protocol that supports accdb and mdb
   files.
*/


import java.util.*;

import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;
import com.sun.star.sdbcx.*;

import com.sun.star.uno.Exception;



public class AccessQuery
{
  //private static final String FNM = "Books.mdb"; 
  private static final String FNM = "Books.accdb"; 



  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    XDriverManager dm = Base.getDriverManager();
    if (dm == null) {
      System.out.println("Could not access Driver manager");
      Lo.closeOffice();
      return;
    }
        
    String url = "sdbc:ado:PROVIDER=Microsoft.ACE.OLEDB.12.0;DATA SOURCE=" + 
                                                       FileIO.getAbsolutePath(FNM);
                                                    // for accdb *and* mdb files
    System.out.println("Using URL: " + url);

    XConnection conn = null;
    try {
      conn = dm.getConnectionWithInfo(url, null);

      //ArrayList<String> tableNames = Base.getTablesNames(conn);
      ArrayList<String> tableNames = Base.getTablesNamesMD(conn);
      System.out.println("No. of tables: " + tableNames.size());
      System.out.println( Arrays.toString(tableNames.toArray()));
               // table list includes an empty string

      XResultSet rs = Base.executeQuery("SELECT * FROM \"Authors\"", conn);
      BaseTablePrinter.printResultSet(rs);  
    }
    catch(Exception e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Lo.closeOffice();
  }  // end of main()



}  // end of AccessQuery class
