
// ThunderbirdQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use the DriverManager to access the Thunderbird e-mail
   reader's address book.

   The table and column names do not match those used in
   the Thurderbird GUI.
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



public class ThunderbirdQuery
{


  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    XDriverManager dm = Base.getDriverManager();
    if (dm == null) {
      System.out.println("Could not access Driver manager");
      Lo.closeOffice();
      return;
    }
        
    XConnection conn = null;
    try {
      conn = dm.getConnectionWithInfo("sdbc:address:thunderbird", null);
                // "sdbc:address:mozilla" for Seamonkey ??
                // https://wiki.openoffice.org/wiki/Documentation/DevGuide/Database/The_SDBC_Driver_for_address_books

      ArrayList<String> tableNames = Base.getTablesNamesMD(conn);
      System.out.println("No. of tables: " + tableNames.size());
      System.out.println( Arrays.toString(tableNames.toArray()));
               // table list is: [AddressBook, CollectedAddressBook]
        /* My ThunderBird address book has two folders:
                 Personal Address Book
                 Collected Addresses
        */

      Base.displayTablesSchema(conn, false); // use getTablesNamesMD()

      XResultSet rs = Base.executeQuery("SELECT \"E-mail\" FROM \"AddressBook\"", conn);
      BaseTablePrinter.printResultSet(rs);  
    }
    catch(Exception e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Lo.closeOffice();
  }  // end of main()



}  // end of ThunderbirdQuery class
