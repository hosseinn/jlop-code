
// DBCreate.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use SQL commands to create an embedded Firebird or HSQLDB database.

   Base.createBaseDoc() creates a new file containing an empty
   database. 

   XFlushable.flush() is needed to store table data in the database.

   Base.refreshTables() is needed if Base's GUI is visible and you
   want the Tables view to be updated.

   Base.showTables() open the tables in the Tables view, but if
   these windows are open when the programs exits, then Office hangs.

   Base.displayDatabase() is a better way to show all the tables in
   a series of JFrames.

   Usage:
     run DBCreate

   DBCmdsCreate.java is a more complex example of creating a database.
*/


import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class DBCreate
{
  private static final String FNM = "spies.odb";


  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.createBaseDoc(FNM,
                                               //  Base.HSQLDB, loader);
                                               Base.FIREBIRD, loader);
    if (dbDoc == null) {
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", ""); 

      //GUI.setVisible(dbDoc, true);
      //Lo.delay(500);  // wait for GUI to appear

      // create a table in the database
      Base.exec("CREATE TABLE SPIES " +
           "( FIRSTNAME VARCHAR(50), LASTNAME VARCHAR(50), ID VARCHAR(50), " +
           "PRIMARY KEY (ID) )", conn);

      Base.exec("INSERT INTO SPIES VALUES( 'James', 'Bond', '007')", conn);
      Base.exec("INSERT INTO SPIES VALUES( 'Johnny', 'English', '013')", conn);
      Base.exec("INSERT INTO SPIES VALUES( 'Maxwell', 'Smart', 'Agent 86')", conn);

      XFlushable flusher = Lo.qi(XFlushable.class, dataSource);
      flusher.flush();   // needed or data not saved to file; can only be called once

      // Base.refreshTables(conn);
            /* must refresh or new table(s) will not be visible 
               inside the Tables window of the Base app; this must
               be done after the call to flush() */
      // Base.showTablesView(dbDoc);

      // Base.displayDatabase(conn);


      System.out.println();
      XResultSet rs = Base.executeQuery("SELECT * FROM SPIES", conn);
      BaseTablePrinter.printResultSet(rs);

      // BaseTablePrinter.printTable(conn, "SPIES");
    }
    catch(SQLException e) {
      System.out.println(e);
    }

   Base.closeConnection(conn);
   Base.closeBaseDoc(dbDoc);
   Lo.closeOffice();
  }  // end of main()



}  // end of DBCreate class
